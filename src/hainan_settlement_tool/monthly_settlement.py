from __future__ import annotations

import argparse
import json
import math
import re
import shutil
from collections import defaultdict
from copy import copy
from dataclasses import dataclass, asdict
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils.datetime import to_excel
from openpyxl.utils import get_column_letter


LEDGER_SHEET = "海南2026年售电结算台账"
DATA_START_ROW = 5
YEAR = 2026
DEFAULT_PLATFORM_INVOICE = "走平台扣13%"
PAYMENT_PARTY_OVERRIDES = {
    ("海南精研科技有限公司", "代理费"): (3, "清能"),
}
THIN_BLACK_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


@dataclass
class DetailRow:
    ledger_row: int
    customer: str
    owner: str
    entity: str
    kind: str
    total: float
    sharp: float
    peak: float
    flat: float
    valley: float
    peak_flat: float
    valley_flat: float
    ratio: float
    unit_price: float
    tax_rate: float
    expected_net: float


@dataclass
class GroupTotal:
    kind: str
    owner: str
    entity: str
    display_entity: str
    rows: int
    expected_net: float
    output_file: str


def n(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return 0.0
        if text.startswith("="):
            return 0.0
        text = text.replace(",", "")
        try:
            return float(text)
        except ValueError:
            return 0.0
    try:
        if math.isnan(value):
            return 0.0
    except TypeError:
        pass
    return float(value)


def s(value: Any) -> str:
    return "" if value is None else str(value).strip()


def norm_name(value: str) -> str:
    text = s(value)
    outer = re.fullmatch(r"[\u4e00-\u9fa5]{2,4}（(.+)）", text)
    if outer:
        text = outer.group(1)
    text = re.sub(r"\s+", "", text)
    text = text.replace("（个体工商户）", "")
    text = text.replace("(个体工商户)", "")
    text = text.replace("绿洲森焱", "绿舟森焱")
    return text


def safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def payment_party_for(entity: str, kind: str, month: int, default_party: str = "清辉") -> str:
    override = PAYMENT_PARTY_OVERRIDES.get((norm_name(entity), kind))
    if override and month >= override[0]:
        return override[1]
    return default_party or "清辉"


def month_start_col(ws, month: int) -> int:
    for col in range(1, ws.max_column + 1):
        if s(ws.cell(1, col).value) == f"{month}月":
            return col
    raise ValueError(f"未找到 {month}月 的台账区块")


def ledger_sheet(wb):
    if LEDGER_SHEET in wb.sheetnames:
        return wb[LEDGER_SHEET]
    if "Sheet1" in wb.sheetnames:
        return wb["Sheet1"]
    for sheet in wb.worksheets:
        if "售电结算台账" in s(sheet["A1"].value):
            return sheet
    raise KeyError(f"找不到台账主表：{LEDGER_SHEET}")


def read_ledger_rows(ledger_path: Path, month: int) -> tuple[list[DetailRow], list[DetailRow]]:
    wb_values = load_workbook(ledger_path, data_only=True)
    wb_formulas = load_workbook(ledger_path, data_only=False)
    ws = ledger_sheet(wb_values)
    ws_formula = ledger_sheet(wb_formulas)
    start = month_start_col(ws, month)

    def get_numeric(row: int, col: int, seen: set[tuple[int, int]] | None = None) -> float:
        value = ws.cell(row, col).value
        if value not in (None, ""):
            return n(value)
        formula_value = ws_formula.cell(row, col).value
        if formula_value in (None, ""):
            return 0.0
        if not (isinstance(formula_value, str) and formula_value.startswith("=")):
            return n(formula_value)

        text = formula_value[1:].replace("$", "").strip()
        match = re.fullmatch(r"([A-Z]{1,3})(\d+)", text)
        if match:
            from openpyxl.utils.cell import column_index_from_string

            seen = seen or set()
            target = (int(match.group(2)), column_index_from_string(match.group(1)))
            if target in seen:
                return 0.0
            seen.add(target)
            return get_numeric(target[0], target[1], seen)
        return 0.0

    def net_amount(row: int, cached_col: int, ratio_col: int, unit_col: int, tax_col: int) -> float:
        cached = get_numeric(row, cached_col)
        if cached:
            return cached
        total = get_numeric(row, start)
        ratio = get_numeric(row, ratio_col)
        unit_price = get_numeric(row, unit_col)
        tax_rate = get_numeric(row, tax_col)
        gross = total * ratio * unit_price
        return round(gross - gross / 1.13 * tax_rate, 4)

    proxy_rows: list[DetailRow] = []
    inter_rows: list[DetailRow] = []
    for r in range(4, ws.max_row + 1):
        customer = s(ws.cell(r, 3).value)
        if not customer:
            continue
        total = get_numeric(r, start)
        if total <= 0:
            continue

        owner = s(ws.cell(r, 10).value)
        developer = s(ws.cell(r, 8).value)
        inter_name = s(ws.cell(r, 19).value)
        common = dict(
            ledger_row=r,
            customer=customer,
            owner=owner,
            total=total,
            sharp=get_numeric(r, start + 1),
            peak=get_numeric(r, start + 2),
            flat=get_numeric(r, start + 3),
            valley=get_numeric(r, start + 4),
            peak_flat=get_numeric(r, start + 5),
            valley_flat=get_numeric(r, start + 6),
        )

        inter_net = net_amount(r, start + 12, start + 7, start + 8, start + 10)
        if inter_name and inter_net:
            inter_rows.append(
                DetailRow(
                    **common,
                    entity=inter_name,
                    kind="居间",
                    ratio=get_numeric(r, start + 7),
                    unit_price=get_numeric(r, start + 8),
                    tax_rate=get_numeric(r, start + 10),
                    expected_net=inter_net,
                )
            )

        proxy_net = net_amount(r, start + 18, start + 13, start + 14, start + 16)
        if developer and proxy_net:
            proxy_rows.append(
                DetailRow(
                    **common,
                    entity=developer,
                    kind="代理",
                    ratio=get_numeric(r, start + 13),
                    unit_price=get_numeric(r, start + 14),
                    tax_rate=get_numeric(r, start + 16),
                    expected_net=proxy_net,
                )
            )
    return proxy_rows, inter_rows


def template_index(proxy_template_root: Path, inter_template_root: Path) -> dict[tuple[str, str, str], Path]:
    result: dict[tuple[str, str, str], Path] = {}
    roots = [("代理", proxy_template_root), ("居间", inter_template_root)]
    for kind, root in roots:
        if not root.exists():
            continue
        for path in root.rglob("*.xlsx"):
            if path.name.startswith("~$"):
                continue
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
            except Exception:
                continue
            month_sheets = [
                name
                for name in wb.sheetnames
                if re.fullmatch(r"\d+月", s(name))
            ]
            sheet_name = max(month_sheets, key=lambda name: int(name[:-1])) if month_sheets else wb.sheetnames[-1]
            ws = wb[sheet_name]
            raw_entity = s(ws["A2"].value).replace("代理名称:", "")
            owner = path.parent.name.replace(" - 海南2026", "")
            result[(kind, norm_name(owner), norm_name(raw_entity))] = path
    return result


def copy_row_format(ws, src_row: int, dst_row: int) -> None:
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(src_row, col)
        dst = ws.cell(dst_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.font:
            dst.font = copy(src.font)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.border:
            dst.border = copy(src.border)
        if src.protection:
            dst.protection = copy(src.protection)


def unmerge_overlapping_rows(ws, start_row: int, end_row: int, min_col: int = 1, max_col: int | None = None) -> None:
    max_col = max_col or ws.max_column
    for merged in list(ws.merged_cells.ranges):
        if (
            merged.min_row <= end_row
            and merged.max_row >= start_row
            and merged.min_col <= max_col
            and merged.max_col >= min_col
        ):
            try:
                ws.unmerge_cells(str(merged))
            except KeyError:
                try:
                    ws.merged_cells.ranges.remove(merged)
                except KeyError:
                    pass


def find_total_row(ws) -> int:
    for row in range(DATA_START_ROW, ws.max_row + 1):
        if s(ws.cell(row, 1).value) == "合计":
            return row
    raise ValueError(f"{ws.title} 未找到合计行")


def find_cell_contains(ws, needle: str) -> tuple[int, int] | None:
    for row in range(1, min(ws.max_row, 8) + 1):
        for col in range(1, min(ws.max_column, 24) + 1):
            if needle in s(ws.cell(row, col).value):
                return row, col
    return None


def set_top_titles(ws, kind: str, entity: str, month: int, display_entity: str | None = None) -> None:
    ws["A1"] = f"{kind}费用结算单"
    ws["A2"] = f"代理名称:{display_entity or entity}"
    period_cell = find_cell_contains(ws, "所属期")
    if period_cell:
        ws.cell(*period_cell).value = f"所属期：{YEAR} 年 {month:02d} 月"
    date_cell = find_cell_contains(ws, "结算日期")
    if date_cell:
        next_month = month + 1
        year = YEAR
        if next_month == 13:
            next_month = 1
            year += 1
        ws.cell(*date_cell).value = f"结算日期：{year} 年 {next_month:02d} 月 15 日"


def prepare_month_sheet(wb, month: int):
    title = f"{month}月"
    if title in wb.sheetnames:
        del wb[title]
    source_title = f"{month - 1}月" if f"{month - 1}月" in wb.sheetnames else "3月"
    if source_title not in wb.sheetnames:
        source_title = wb.sheetnames[-1]
    ws = wb.copy_worksheet(wb[source_title])
    ws.title = title
    return ws


def adjust_data_rows(ws, count: int) -> int:
    total_row = find_total_row(ws)
    existing = total_row - DATA_START_ROW
    if count > existing:
        insert_at = total_row
        ws.insert_rows(insert_at, count - existing)
        for row in range(insert_at, insert_at + count - existing):
            copy_row_format(ws, DATA_START_ROW, row)
    elif count < existing:
        ws.delete_rows(DATA_START_ROW + count, existing - count)
    total_row = find_total_row(ws)
    last_data_row = DATA_START_ROW + count - 1
    for merged in list(ws.merged_cells.ranges):
        if (
            merged.min_row <= total_row
            and merged.max_row >= DATA_START_ROW
            and merged.min_col <= 17
            and merged.max_col >= 1
        ):
            ws.unmerge_cells(str(merged))
    return total_row


def write_detail_sheet(
    ws,
    kind: str,
    entity: str,
    month: int,
    rows: list[DetailRow],
    display_entity: str | None = None,
) -> None:
    set_top_titles(ws, kind, entity, month, display_entity)
    total_row = adjust_data_rows(ws, len(rows))
    for idx, row in enumerate(rows, start=DATA_START_ROW):
        ws.cell(idx, 1).value = idx - DATA_START_ROW + 1
        ws.cell(idx, 2).value = row.customer
        values = [
            row.total,
            row.sharp,
            row.peak,
            row.flat,
            row.valley,
            row.peak_flat,
            row.valley_flat,
            row.ratio,
            row.unit_price,
        ]
        for offset, value in enumerate(values, start=3):
            ws.cell(idx, offset).value = round(value, 4) if isinstance(value, float) else value
        ws.cell(idx, 12).value = f"=C{idx}*J{idx}*K{idx}"
        ws.cell(idx, 13).value = f"=L{idx}-N{idx}"
        ws.cell(idx, 14).value = None
        ws.cell(idx, 15).value = f"=M{idx}/1.13*Q{idx}"
        ws.cell(idx, 16).value = f"=ROUND(M{idx}-O{idx},4)"
        ws.cell(idx, 17).value = row.tax_rate
        for col in range(18, ws.max_column + 1):
            ws.cell(idx, col).value = None

    if rows:
        last = DATA_START_ROW + len(rows) - 1
        ws.cell(total_row, 1).value = "合计"
        for col in range(3, 8):
            letter = get_column_letter(col)
            ws.cell(total_row, col).value = f"=SUM({letter}{DATA_START_ROW}:{letter}{last})"
        for col in range(12, 17):
            letter = get_column_letter(col)
            ws.cell(total_row, col).value = f"=SUM({letter}{DATA_START_ROW}:{letter}{last})"


def ensure_output_workbook(
    template_map: dict[tuple[str, str, str], Path],
    proxy_template_root: Path,
    inter_template_root: Path,
    output_root: Path,
    kind: str,
    owner: str,
    entity: str,
) -> tuple[Path, bool]:
    source = template_map.get((kind, norm_name(owner), norm_name(entity)))
    if source:
        if kind == "代理":
            rel = source.relative_to(proxy_template_root)
            target = output_root / "2026年代理 - 海南" / rel
        else:
            rel = source.relative_to(inter_template_root)
            target = output_root / "2026年居间 - 海南" / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        if not target.exists():
            shutil.copy2(source, target)
        return target, True

    base_root = output_root / ("2026年代理 - 海南" if kind == "代理" else "2026年居间 - 海南")
    folder = base_root / f"{safe_filename(owner)} - 海南2026"
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / f"{safe_filename(entity)} 2026海南.xlsx"
    if target.exists():
        return target, False

    # New entity: clone a same-kind workbook as a visual template, then keep only the new month sheet.
    candidates = [p for (k, _, _), p in template_map.items() if k == kind]
    if not candidates:
        raise ValueError(f"没有可用的{kind}分表模板")
    shutil.copy2(candidates[0], target)
    return target, False


def prior_sheet_display_entity(wb, month: int) -> str | None:
    for candidate in (f"{month - 1}月", "3月", "2月", "1月"):
        if candidate in wb.sheetnames:
            text = s(wb[candidate]["A2"].value)
            if text.startswith("代理名称:"):
                return text.replace("代理名称:", "", 1)
    return None


def build_split_files(
    template_root: Path,
    output_root: Path,
    month: int,
    proxy_rows: list[DetailRow],
    inter_rows: list[DetailRow],
    *,
    proxy_template_root: Path | None = None,
    inter_template_root: Path | None = None,
) -> list[GroupTotal]:
    proxy_template_root = proxy_template_root or template_root / "2026年代理 - 海南"
    inter_template_root = inter_template_root or template_root / "2026年居间 - 海南"
    template_map = template_index(proxy_template_root, inter_template_root)
    grouped: dict[tuple[str, str, str], list[DetailRow]] = defaultdict(list)
    for row in proxy_rows:
        grouped[("代理", row.owner, row.entity)].append(row)
    for row in inter_rows:
        grouped[("居间", row.owner, row.entity)].append(row)

    totals: list[GroupTotal] = []
    for (kind, owner, entity), rows in sorted(grouped.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])):
        path, matched_existing_template = ensure_output_workbook(
            template_map,
            proxy_template_root,
            inter_template_root,
            output_root,
            kind,
            owner,
            entity,
        )
        wb = load_workbook(path)
        display_entity = prior_sheet_display_entity(wb, month) if matched_existing_template else entity
        ws = prepare_month_sheet(wb, month)
        write_detail_sheet(ws, kind, entity, month, rows, display_entity)
        try:
            wb.calculation.fullCalcOnLoad = True
            wb.calculation.forceFullCalc = True
        except Exception:
            pass
        wb.save(path)
        totals.append(
            GroupTotal(
                kind=f"{kind}费",
                owner=owner,
                entity=entity,
                display_entity=display_entity or entity,
                rows=len(rows),
                expected_net=round(sum(row.expected_net for row in rows), 4),
                output_file=str(path),
            )
        )
    return totals


def summary_col(ws, header: str) -> int:
    for col in range(1, ws.max_column + 1):
        if s(ws.cell(2, col).value) == header:
            return col
    raise ValueError(f"{ws.title} 未找到列：{header}")


def value_at(values: list[Any], col: int, default: Any = None) -> Any:
    idx = col - 1
    if idx < 0 or idx >= len(values):
        return default
    return values[idx]


def summary_sheet_has_marker(ws) -> bool:
    try:
        summary_col(ws, "累计代理费总计")
        return True
    except ValueError:
        return False


def resolve_summary_sheet_name(wb, role: str = "main", *, required: bool = True) -> str | None:
    def clean(name: str) -> str:
        return re.sub(r"\s+", "", s(name))

    names = wb.sheetnames
    normalized = {clean(name): name for name in names}
    role_labels = {
        "main": "主汇总表",
        "qingneng": "清能汇总表",
        "qinghui": "清辉汇总表",
    }

    if role == "main":
        for exact in ("汇总表", "代理费汇总表"):
            if exact in normalized:
                return normalized[exact]
        candidates = [
            name
            for name in names
            if "汇总表" in clean(name)
            and "清能" not in clean(name)
            and "清辉" not in clean(name)
            and summary_sheet_has_marker(wb[name])
        ]
    elif role == "qingneng":
        candidates = [name for name in names if "清能" in clean(name) and "汇总" in clean(name) and summary_sheet_has_marker(wb[name])]
    elif role == "qinghui":
        candidates = [name for name in names if "清辉" in clean(name) and "汇总" in clean(name) and summary_sheet_has_marker(wb[name])]
    else:
        raise ValueError(f"未知汇总表角色：{role}")

    if candidates:
        return candidates[0]

    if role == "main":
        marker_sheets = [
            name
            for name in names
            if "清能" not in clean(name)
            and "清辉" not in clean(name)
            and summary_sheet_has_marker(wb[name])
        ]
        if len(marker_sheets) == 1:
            return marker_sheets[0]

    if required:
        label = role_labels.get(role, role)
        raise KeyError(
            f"选择的汇总表模板缺少{label}。当前工作表：{', '.join(names)}。"
            "请在“上月/修正版汇总表”选择代理费汇总表文件，不要选择售电结算台账或代理/居间分表。"
        )
    return None


def sheet_meta(summary_path: Path, sheet_name: str | None = None, *, role: str = "main") -> dict[tuple[str, str], dict[str, Any]]:
    wb = load_workbook(summary_path, data_only=False)
    actual_sheet = sheet_name if sheet_name in wb.sheetnames else resolve_summary_sheet_name(wb, role)
    ws = wb[actual_sheet]
    cumulative_col = summary_col(ws, "累计代理费总计")
    max_meta_col = max(ws.max_column, cumulative_col + 9, 40)
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for row in range(4, ws.max_row + 1):
        entity = s(ws.cell(row, 2).value)
        kind = s(ws.cell(row, 3).value)
        if not entity or "审核" in entity:
            continue
        result[(norm_name(entity), kind)] = {
            "row": row,
            "entity": entity,
            "kind": kind,
            "values": [ws.cell(row, col).value for col in range(1, max_meta_col + 1)],
            "cumulative_col": cumulative_col,
            "payment_party": ws.cell(row, cumulative_col + 8).value,
        }
    return result


def current_month_first_col(ws) -> int:
    # Last month block before cumulative columns.
    return summary_col(ws, "累计代理费总计") - 6


def ensure_summary_month_header_borders(ws, month_col: int) -> None:
    """Give the newly inserted month header real Excel borders.

    Historical templates sometimes rely on visible gridlines in row 3, but
    copied white fills hide those gridlines in generated month blocks.
    """
    for col in range(month_col, month_col + 6):
        reference = ws.cell(4, col).border
        border = copy(reference) if reference else copy(THIN_BLACK_BORDER)
        if not any((border.left.style, border.right.style, border.top.style, border.bottom.style)):
            border = copy(THIN_BLACK_BORDER)
        ws.cell(3, col).border = copy(border)

    ws.cell(2, month_col).border = copy(ws.cell(3, month_col).border)
    ws.cell(2, month_col + 5).border = copy(ws.cell(3, month_col + 5).border)


def insert_month_block(ws, month: int) -> int:
    cumulative_col = summary_col(ws, "累计代理费总计")
    insert_at = cumulative_col
    ws.insert_cols(insert_at, 6)

    # Copy width/style from previous month block.
    for i in range(6):
        src_col = insert_at - 6 + i
        dst_col = insert_at + i
        src_letter = get_column_letter(src_col)
        dst_letter = get_column_letter(dst_col)
        ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width
        ws.column_dimensions[dst_letter].hidden = False
        for row in range(1, ws.max_row + 1):
            src = ws.cell(row, src_col)
            dst = ws.cell(row, dst_col)
            if src.has_style:
                dst._style = copy(src._style)
            if src.number_format:
                dst.number_format = src.number_format
            dst.alignment = copy(src.alignment)
            dst.font = copy(src.font)
            dst.fill = copy(src.fill)
            dst.border = copy(src.border)

    for merged in list(ws.merged_cells.ranges):
        if (
            merged.min_row <= 3
            and merged.max_row >= 2
            and merged.min_col <= insert_at + 5
            and merged.max_col >= insert_at
        ):
            ws.unmerge_cells(str(merged))

    labels = ["代理费", "居间费", "退补电费", "当月抵扣", "费用合计"]
    ws.cell(2, insert_at).value = f"{YEAR}年{month}月"
    for i, label in enumerate(labels):
        ws.cell(3, insert_at + i).value = label
    ws.cell(2, insert_at + 5).value = "当月实际支付"
    ws.cell(3, insert_at + 5).value = None
    ensure_summary_month_header_borders(ws, insert_at)
    try:
        ws.merge_cells(start_row=2, start_column=insert_at, end_row=2, end_column=insert_at + 4)
        ws.merge_cells(start_row=2, start_column=insert_at + 5, end_row=3, end_column=insert_at + 5)
    except ValueError:
        pass
    return insert_at


def apply_summary_hidden_columns(ws, month_col: int, cumulative_col: int, main_sheet: bool) -> None:
    last_relevant_col = cumulative_col + 9
    for col in range(1, min(ws.max_column, last_relevant_col) + 1):
        ws.column_dimensions[get_column_letter(col)].hidden = False

    hidden: set[int] = {4, 5, 7}
    for start in range(12, month_col, 6):
        hidden.update({start, start + 2, start + 3, start + 4, start + 5})

    if main_sheet:
        hidden.update({cumulative_col + 1, cumulative_col + 5, cumulative_col + 6})
    else:
        hidden.update({cumulative_col, cumulative_col + 1, cumulative_col + 5, cumulative_col + 6})

    for col in hidden:
        ws.column_dimensions[get_column_letter(col)].hidden = True

    for merged in list(ws.merged_cells.ranges):
        if merged.min_row == 1 and merged.max_row == 1 and merged.min_col == 1:
            ws.unmerge_cells(str(merged))
    try:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_relevant_col)
    except ValueError:
        pass


def trim_summary_tail_columns(ws, cumulative_col: int) -> None:
    # Some historical sheets carry empty styled cells all the way to XFD.
    # Inserting a new month would push those phantom columns past Excel's limit.
    last_relevant_col = cumulative_col + 9
    if ws.max_column > last_relevant_col:
        ws.delete_cols(last_relevant_col + 1, ws.max_column - last_relevant_col)


def cell_in_merge(ws, row: int, col: int) -> bool:
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return True
    return False


def apply_summary_header_merges(ws, last_relevant_col: int) -> None:
    for col in range(1, last_relevant_col + 1):
        if s(ws.cell(2, col).value) and not s(ws.cell(3, col).value):
            if not cell_in_merge(ws, 2, col) and not cell_in_merge(ws, 3, col):
                try:
                    ws.merge_cells(start_row=2, start_column=col, end_row=3, end_column=col)
                except ValueError:
                    pass


def parse_monthly_deduction(text: str) -> float:
    text = s(text)
    m = re.search(r"每月扣除([0-9.]+)万", text)
    if m:
        return float(m.group(1))
    m = re.search(r"每月扣除([0-9.]+)元", text)
    if m:
        return round(float(m.group(1)) / 10000, 4)
    return 0.0


def find_summary_total_row(ws, start_row: int = 4) -> int:
    for row in range(start_row, ws.max_row + 1):
        if s(ws.cell(row, 1).value) == "合计":
            return row

    for row in range(start_row, ws.max_row + 1):
        row_values = [ws.cell(row, col).value for col in range(12, min(ws.max_column, 80) + 1)]
        if any(isinstance(value, str) and value.upper().startswith("=SUM(") for value in row_values):
            return row

    for row in range(start_row, ws.max_row + 1):
        labels = [s(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 12) + 1)]
        if any(("审批" in label or "审核" in label or "制表" in label) for label in labels):
            candidate = row - 1
            if candidate >= start_row:
                return candidate

    raise ValueError(f"{ws.title} 未找到合计行")


def write_summary_sheet(
    ws,
    totals: list[GroupTotal],
    meta: dict[tuple[str, str], dict[str, Any]],
    month: int,
    *,
    main_sheet: bool,
    allowed_keys: set[tuple[str, str]] | None = None,
) -> list[str]:
    old_cumulative_col = summary_col(ws, "累计代理费总计")
    month_col = insert_month_block(ws, month)
    cumulative_col = month_col + 6
    warnings: list[str] = []
    total_by_key = {(norm_name(t.entity), t.kind): t for t in totals}

    # Remove old approval/footer rows by detecting non-data names, then rewrite all known rows.
    data_rows = []
    for info in meta.values():
        if allowed_keys is not None and (norm_name(info["entity"]), info["kind"]) not in allowed_keys:
            continue
        data_rows.append(info)
    data_rows.sort(key=lambda item: item["row"])
    known_keys = {(norm_name(info["entity"]), info["kind"]) for info in data_rows}
    new_totals = [total for total in totals if (norm_name(total.entity), total.kind) not in known_keys]

    start_row = 4
    old_data_count = len(data_rows)
    target_data_count = old_data_count + len(new_totals)
    total_row = find_summary_total_row(ws, start_row)
    existing_data_count = total_row - start_row
    if target_data_count > existing_data_count:
        insert_at = total_row
        ws.insert_rows(insert_at, target_data_count - existing_data_count)
        unmerge_overlapping_rows(ws, insert_at, insert_at + target_data_count - existing_data_count - 1, 1, cumulative_col + 9)
    elif target_data_count < existing_data_count:
        ws.delete_rows(start_row + target_data_count, existing_data_count - target_data_count)
    total_row = start_row + target_data_count

    for idx, info in enumerate(data_rows, start=start_row):
        values = info["values"]
        for old_col, value in enumerate(values, start=1):
            target_col = old_col if old_col < month_col else old_col + 6
            ws.cell(idx, target_col).value = value
        key = (norm_name(info["entity"]), info["kind"])
        total = total_by_key.get(key)
        proxy_value = total.expected_net if total and total.kind == "代理费" else None
        inter_value = total.expected_net if total and total.kind == "居间费" else None
        ws.cell(idx, month_col).value = proxy_value
        ws.cell(idx, month_col + 1).value = inter_value
        ws.cell(idx, month_col + 2).value = None

        fee = n(proxy_value) + n(inter_value)
        loan_total = n(value_at(values, old_cumulative_col + 1, 0.0))
        previous_deducted = n(value_at(values, old_cumulative_col + 2, 0.0))
        monthly_deduction = parse_monthly_deduction(value_at(values, old_cumulative_col + 5, ""))
        remaining = max(loan_total - previous_deducted, 0.0)
        deduction = 0.0
        if remaining > 0 and fee > 0:
            deduction = min(fee, remaining, monthly_deduction or remaining)
            deduction = round(deduction, 4)
        ws.cell(idx, month_col + 3).value = deduction if deduction else None
        ws.cell(idx, month_col + 4).value = f"={get_column_letter(month_col)}{idx}+{get_column_letter(month_col+1)}{idx}+{get_column_letter(month_col+2)}{idx}"
        ws.cell(idx, month_col + 5).value = f"={get_column_letter(month_col)}{idx}+{get_column_letter(month_col+1)}{idx}+{get_column_letter(month_col+2)}{idx}-{get_column_letter(month_col+3)}{idx}"

        month_fee_cols = [col for col in range(16, cumulative_col, 6)]
        month_deduct_cols = [col for col in range(15, cumulative_col, 6)]
        ws.cell(idx, cumulative_col).value = "=" + "+".join(f"{get_column_letter(col)}{idx}" for col in month_fee_cols)
        ws.cell(idx, cumulative_col + 2).value = "=" + "+".join(f"{get_column_letter(col)}{idx}" for col in month_deduct_cols)
        ws.cell(idx, cumulative_col + 3).value = f"={get_column_letter(cumulative_col+1)}{idx}-{get_column_letter(cumulative_col+2)}{idx}"
        old_party = s(value_at(values, old_cumulative_col + 8, "清辉")) or "清辉"
        ws.cell(idx, cumulative_col + 8).value = payment_party_for(info["entity"], info["kind"], month, old_party)

    next_row = start_row + old_data_count
    if new_totals:
        template_row = next_row - 1
        for i, total in enumerate(new_totals):
            row = next_row + i
            copy_row_format(ws, template_row, row)
            ws.cell(row, 1).value = row - 3
            ws.cell(row, 2).value = total.entity
            ws.cell(row, 3).value = total.kind
            ws.cell(row, 4).value = "否"
            ws.cell(row, 6).value = total.entity
            ws.cell(row, 7).value = "平台"
            ws.cell(row, 8).value = 0
            ws.cell(row, 9).value = 0.13
            ws.cell(row, 10).value = 0.13
            ws.cell(row, 11).value = total.owner
            ws.cell(row, month_col).value = total.expected_net if total.kind == "代理费" else None
            ws.cell(row, month_col + 1).value = total.expected_net if total.kind == "居间费" else None
            ws.cell(row, month_col + 4).value = f"={get_column_letter(month_col)}{row}+{get_column_letter(month_col+1)}{row}+{get_column_letter(month_col+2)}{row}"
            ws.cell(row, month_col + 5).value = f"={get_column_letter(month_col)}{row}+{get_column_letter(month_col+1)}{row}+{get_column_letter(month_col+2)}{row}-{get_column_letter(month_col+3)}{row}"
            ws.cell(row, cumulative_col).value = "=" + "+".join(f"{get_column_letter(col)}{row}" for col in range(16, cumulative_col, 6))
            ws.cell(row, cumulative_col + 2).value = "=" + "+".join(f"{get_column_letter(col)}{row}" for col in range(15, cumulative_col, 6))
            ws.cell(row, cumulative_col + 3).value = f"={get_column_letter(cumulative_col+1)}{row}-{get_column_letter(cumulative_col+2)}{row}"
            ws.cell(row, cumulative_col + 7).value = f"{YEAR}{month:02d}"
            ws.cell(row, cumulative_col + 8).value = payment_party_for(total.entity, total.kind, month, "清辉")
            warnings.append(f"新增汇总主体：{total.kind} {total.entity}（负责人：{total.owner}）")

    if total_row:
        ws.cell(total_row, 1).value = "合计"
        for col in range(12, cumulative_col + 4):
            letter = get_column_letter(col)
            ws.cell(total_row, col).value = f"=SUM({letter}4:{letter}{total_row-1})"
        for col in range(cumulative_col + 4, min(cumulative_col + 10, ws.max_column + 1)):
            ws.cell(total_row, col).value = None

    apply_summary_hidden_columns(ws, month_col, cumulative_col, main_sheet)
    trim_summary_tail_columns(ws, cumulative_col)
    apply_summary_header_merges(ws, cumulative_col + 9)
    return warnings


def build_summary(summary_template: Path, output_root: Path, month: int, totals: list[GroupTotal]) -> Path:
    output = output_root / f"【2026年海南省代理费汇总表-{month}月自动生成】.xlsx"
    shutil.copy2(summary_template, output)
    wb = load_workbook(output)

    main_sheet_name = resolve_summary_sheet_name(wb, "main")
    qn_sheet_name = resolve_summary_sheet_name(wb, "qingneng", required=False)
    qh_sheet_name = resolve_summary_sheet_name(wb, "qinghui", required=False)

    main_meta = sheet_meta(summary_template, main_sheet_name, role="main")
    party_by_key = {
        key: payment_party_for(info["entity"], info["kind"], month, s(info.get("payment_party")) or "清辉")
        for key, info in main_meta.items()
    }
    main = wb[main_sheet_name]
    warnings = write_summary_sheet(main, totals, main_meta, month, main_sheet=True)

    # Payment sheets use the same template logic, filtered by payment party.
    if qn_sheet_name:
        meta = sheet_meta(summary_template, qn_sheet_name, role="qingneng")
        qn_totals = [
            total
            for total in totals
            if party_by_key.get((norm_name(total.entity), total.kind), "清辉") == "清能"
        ]
        qn_allowed = {key for key, party in party_by_key.items() if party == "清能"}
        qn_allowed.update((norm_name(total.entity), total.kind) for total in qn_totals)
        warnings.extend(write_summary_sheet(wb[qn_sheet_name], qn_totals, meta, month, main_sheet=False, allowed_keys=qn_allowed))
    if qh_sheet_name:
        meta = sheet_meta(summary_template, qh_sheet_name, role="qinghui")
        qh_totals = [
            total
            for total in totals
            if party_by_key.get((norm_name(total.entity), total.kind), "清辉") == "清辉"
        ]
        qh_allowed = {key for key, party in party_by_key.items() if party == "清辉"}
        qh_allowed.update((norm_name(total.entity), total.kind) for total in qh_totals)
        warnings.extend(write_summary_sheet(wb[qh_sheet_name], qh_totals, meta, month, main_sheet=False, allowed_keys=qh_allowed))

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass
    wb.save(output)

    warning_path = output_root / "自动生成汇总提示.txt"
    if warnings:
        warning_path.write_text("\n".join(warnings), encoding="utf-8")
    elif warning_path.exists():
        warning_path.unlink()
    return output


def write_report(output_root: Path, month: int, totals: list[GroupTotal], proxy_rows: list[DetailRow], inter_rows: list[DetailRow], summary_path: Path) -> Path:
    report = {
        "month": month,
        "proxy_groups": sum(1 for t in totals if t.kind == "代理费"),
        "intermediary_groups": sum(1 for t in totals if t.kind == "居间费"),
        "proxy_rows": len(proxy_rows),
        "intermediary_rows": len(inter_rows),
        "proxy_total": round(sum(t.expected_net for t in totals if t.kind == "代理费"), 4),
        "intermediary_total": round(sum(t.expected_net for t in totals if t.kind == "居间费"), 4),
        "summary_path": str(summary_path),
        "groups": [asdict(t) for t in totals],
    }
    path = output_root / f"{month}月自动生成校对报告.json"
    path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def run(args: argparse.Namespace) -> None:
    template_root = Path(args.template_root)
    output_root = Path(args.output_root)
    ledger_path = Path(args.ledger)
    summary_template = Path(args.summary_template)
    month = int(args.month)

    proxy_rows, inter_rows = read_ledger_rows(ledger_path, month)
    totals = build_split_files(template_root, output_root, month, proxy_rows, inter_rows)
    summary_path = build_summary(summary_template, output_root, month, totals)
    report_path = write_report(output_root, month, totals, proxy_rows, inter_rows, summary_path)
    print(json.dumps({
        "ok": True,
        "month": month,
        "split_files": len(totals),
        "summary": str(summary_path),
        "report": str(report_path),
        "proxy_total": round(sum(t.expected_net for t in totals if t.kind == "代理费"), 4),
        "intermediary_total": round(sum(t.expected_net for t in totals if t.kind == "居间费"), 4),
    }, ensure_ascii=False, indent=2))


def main() -> None:
    parser = argparse.ArgumentParser(description="海南售电月度结算分表和汇总表生成")
    parser.add_argument("--month", required=True, type=int)
    parser.add_argument("--ledger", required=True)
    parser.add_argument("--template-root", required=True)
    parser.add_argument("--output-root", required=True)
    parser.add_argument("--summary-template", required=True)
    run(parser.parse_args())


if __name__ == "__main__":
    main()
