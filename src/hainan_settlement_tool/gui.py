from __future__ import annotations

import json
import queue
import subprocess
import threading
import traceback
from pathlib import Path
from tkinter import (
    BooleanVar,
    Button,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    LabelFrame,
    StringVar,
    Text,
    Tk,
    filedialog,
    messagebox,
)
from tkinter.ttk import Combobox

from .workflow import (
    clean_power_data,
    run_settlement,
    update_ledger,
)


APP_TITLE = "海南售电结算自动化工具"


def as_path(text: str) -> Path | None:
    value = text.strip()
    return Path(value) if value else None


class WorkflowArgs:
    pass


class App:
    def __init__(self, root: Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1040x860")
        self.root.minsize(980, 780)

        self.queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.running = False

        self.month_var = IntVar(value=5)
        self.month_root_var = StringVar()
        self.output_dir_var = StringVar()
        self.raw_detail_var = StringVar()
        self.power_var = StringVar()
        self.base_ledger_var = StringVar()
        self.reference_ledger_var = StringVar()
        self.completed_ledger_var = StringVar()
        self.proxy_template_dir_var = StringVar()
        self.inter_template_dir_var = StringVar()
        self.summary_template_var = StringVar()
        self.copy_reference_existing_var = BooleanVar(value=False)
        self.allow_missing_owner_var = BooleanVar(value=False)
        self.no_recalc_var = BooleanVar(value=False)
        self.status_var = StringVar(value="空闲")

        self._build_ui()
        self.root.after(150, self._poll_queue)

    def _build_ui(self) -> None:
        top = LabelFrame(self.root, text="操作流程说明")
        top.pack(fill="x", padx=12, pady=(10, 6))
        instruction = (
            "阶段1：基础台账 +（电量处理表 或 原始零售侧明细） -> 输出待整理台账，只写电量、新增客户名称和户号。\n"
            "阶段2：人工整理后的台账 + 上月代理分表文件夹 + 上月居间分表文件夹 + 上月/修正版汇总表 -> 输出代理/居间分表和汇总表。\n"
            "当月结算文件夹、输出文件夹、参考台账都是辅助项。所有结果只写入输出文件夹，不覆盖原始文件。"
        )
        Label(top, text=instruction, justify="left", anchor="w", wraplength=990).pack(fill="x", padx=10, pady=8)

        common = LabelFrame(self.root, text="通用设置")
        common.pack(fill="x", padx=12, pady=6)
        common.grid_columnconfigure(1, weight=1)
        self._month_row(common, 0)
        self._path_row(common, "输出文件夹(可空)", self.output_dir_var, 1, self.pick_output_dir, kind="dir")

        stage1 = LabelFrame(self.root, text="阶段1：清洗并导入台账")
        stage1.pack(fill="x", padx=12, pady=6)
        stage1.grid_columnconfigure(1, weight=1)
        Label(stage1, text="必填：基础台账；再选择“电量处理表”或“原始零售侧明细”其中一个。", anchor="w").grid(
            row=0, column=0, columnspan=5, sticky="we", padx=10, pady=(8, 2)
        )
        self._path_row(stage1, "当月结算文件夹(可选)", self.month_root_var, 1, self.pick_month_root, kind="dir")
        self._path_row(stage1, "基础台账(必填)", self.base_ledger_var, 2, self.pick_base_ledger)
        self._path_row(stage1, "电量处理表(二选一)", self.power_var, 3, self.pick_power)
        self._path_row(stage1, "原始零售侧明细(二选一)", self.raw_detail_var, 4, self.pick_raw_detail)

        stage2 = LabelFrame(self.root, text="阶段2：生成结算成果")
        stage2.pack(fill="x", padx=12, pady=6)
        stage2.grid_columnconfigure(1, weight=1)
        Label(stage2, text="必填：人工整理后的台账、上月代理分表文件夹、上月居间分表文件夹、上月/修正版汇总表。", anchor="w").grid(
            row=0, column=0, columnspan=5, sticky="we", padx=10, pady=(8, 2)
        )
        self._path_row(stage2, "人工整理后的台账(必填)", self.completed_ledger_var, 1, self.pick_completed_ledger)
        self._path_row(stage2, "上月代理分表文件夹(必填)", self.proxy_template_dir_var, 2, self.pick_proxy_template_dir, kind="dir")
        self._path_row(stage2, "上月居间分表文件夹(必填)", self.inter_template_dir_var, 3, self.pick_inter_template_dir, kind="dir")
        self._path_row(stage2, "上月/修正版汇总表(必填)", self.summary_template_var, 4, self.pick_summary_template)

        advanced = LabelFrame(self.root, text="高级/可选")
        advanced.pack(fill="x", padx=12, pady=6)
        advanced.grid_columnconfigure(1, weight=1)
        self._path_row(advanced, "参考台账(可选)", self.reference_ledger_var, 0, self.pick_reference_ledger)
        opts = Frame(advanced)
        opts.grid(row=1, column=0, columnspan=5, sticky="we", padx=10, pady=(4, 8))
        Checkbutton(opts, text="用参考台账覆盖已有客户基础资料", variable=self.copy_reference_existing_var).pack(side="left", padx=(0, 18))
        Checkbutton(opts, text="允许负责人缺失继续生成", variable=self.allow_missing_owner_var).pack(side="left", padx=(0, 18))
        Checkbutton(opts, text="跳过 Excel 重算（更快，仅临时检查用）", variable=self.no_recalc_var).pack(side="left")

        actions = LabelFrame(self.root, text="执行")
        actions.pack(fill="x", padx=12, pady=6)
        self.buttons: list[Button] = []
        for text, command in [
            ("阶段1 清洗并导入台账", self.run_stage1),
            ("阶段2 生成结算成果", self.run_settlement),
            ("仅清洗电量", self.run_clean),
            ("仅更新台账", self.run_ledger),
            ("打开输出文件夹", self.open_output_dir),
        ]:
            btn = Button(actions, text=text, width=20, command=command)
            btn.pack(side="left", padx=8, pady=10)
            self.buttons.append(btn)
        Label(actions, textvariable=self.status_var, anchor="w", justify="left").pack(side="left", fill="x", expand=True, padx=12)

        log_frame = LabelFrame(self.root, text="运行日志")
        log_frame.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        self.log = Text(log_frame, height=16, wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)
        self._log("请选择文件后执行。建议先点“阶段1 清洗并导入台账”生成草稿，人工补齐后再点“阶段2 生成结算成果”。")

    def _month_row(self, parent: Frame, row: int) -> None:
        Label(parent, text="月份").grid(row=row, column=0, sticky="w", padx=10, pady=5)
        box = Combobox(parent, textvariable=self.month_var, values=list(range(2, 13)), width=8, state="readonly")
        box.grid(row=row, column=1, sticky="w", padx=8, pady=5)

    def _path_row(self, parent: Frame, label: str, var: StringVar, row: int, command, kind: str = "file") -> None:
        Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=10, pady=5)
        Entry(parent, textvariable=var).grid(row=row, column=1, columnspan=3, sticky="we", padx=8, pady=5)
        Button(parent, text="浏览", command=command, width=8).grid(row=row, column=4, sticky="e", padx=10, pady=5)

    def _reject_temp_file(self, path: Path, label: str) -> None:
        if path.name.startswith("~$"):
            raise ValueError(f"{label}选到了 Excel 临时文件：{path.name}\n请关闭/另选正式 .xlsx 文件。")

    def _required_file(self, var: StringVar, label: str) -> Path:
        path = as_path(var.get())
        if not path:
            raise ValueError(f"请选择{label}。")
        self._reject_temp_file(path, label)
        if not path.exists():
            raise FileNotFoundError(f"{label}不存在：{path}")
        if not path.is_file():
            raise ValueError(f"{label}不是文件：{path}")
        return path

    def _optional_file(self, var: StringVar, label: str) -> Path | None:
        path = as_path(var.get())
        if not path:
            return None
        self._reject_temp_file(path, label)
        if not path.exists():
            raise FileNotFoundError(f"{label}不存在：{path}")
        if not path.is_file():
            raise ValueError(f"{label}不是文件：{path}")
        return path

    def _required_dir(self, var: StringVar, label: str) -> Path:
        path = as_path(var.get())
        if not path:
            raise ValueError(f"请选择{label}。")
        if not path.exists():
            raise FileNotFoundError(f"{label}不存在：{path}")
        if not path.is_dir():
            raise ValueError(f"{label}不是文件夹：{path}")
        return path

    def _log(self, text: str) -> None:
        self.log.insert("end", text.rstrip() + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def _progress(self, text: str) -> None:
        message = f"{text}。请不要关闭应用窗口。"
        self.queue.put(("log", message))
        self.queue.put(("status", message))

    def _set_running(self, running: bool) -> None:
        self.running = running
        state = "disabled" if running else "normal"
        for btn in self.buttons:
            btn.configure(state=state)

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    self._log(str(payload))
                elif kind == "status":
                    self.status_var.set(str(payload))
                elif kind == "set_completed_ledger":
                    self.completed_ledger_var.set(str(payload))
                elif kind == "done":
                    self._set_running(False)
                    self.status_var.set("空闲")
                    self._log("完成。")
                    messagebox.showinfo("完成", str(payload))
                elif kind == "error":
                    self._set_running(False)
                    self.status_var.set("出错，请查看提示")
                    if isinstance(payload, dict):
                        self._log(str(payload.get("detail", payload.get("message", ""))))
                        messagebox.showerror("出错了", str(payload.get("message", "运行失败")))
                        if payload.get("file_locked"):
                            self._offer_close_excel(str(payload.get("message", "")))
                    else:
                        self._log(str(payload))
                        messagebox.showerror("出错了", str(payload))
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    def _run_background(self, title: str, func) -> None:
        if self.running:
            return
        self._set_running(True)
        self.status_var.set(f"正在执行：{title}。Excel 读取/重算时可能需要几十秒，请稍等。")
        self._log(f"开始：{title}")

        def worker() -> None:
            try:
                message = func()
                self.queue.put(("done", message))
            except Exception as exc:
                message = str(exc).strip("'")
                detail = f"{message}\n\n{traceback.format_exc()}"
                self.queue.put(("error", {"message": message, "detail": detail, "file_locked": self._is_file_locked_error(detail)}))

        threading.Thread(target=worker, daemon=True).start()

    def _is_file_locked_error(self, text: str) -> bool:
        markers = [
            "PermissionError",
            "being used by another process",
            "另一个程序正在使用此文件",
            "拒绝访问",
            "cannot access the file",
            "无法访问",
        ]
        return any(marker in text for marker in markers)

    def _offer_close_excel(self, message: str) -> None:
        prompt = (
            "看起来有 Excel 文件被打开或被后台 Excel 进程占用，导致工具无法继续写入/覆盖。\n\n"
            f"错误信息：{message}\n\n"
            "是否强制关闭所有 Excel 进程？\n"
            "注意：这会关闭当前所有 Excel 窗口，未保存的内容可能丢失。关闭后请重新点击刚才的步骤。"
        )
        if not messagebox.askyesno("文件可能被占用", prompt):
            return
        try:
            completed = subprocess.run(
                ["taskkill", "/F", "/IM", "EXCEL.EXE"],
                capture_output=True,
                text=True,
                encoding="gbk",
                errors="ignore",
                timeout=20,
            )
            detail = (completed.stdout or "") + (completed.stderr or "")
            self._log("已尝试关闭 Excel 进程：\n" + (detail.strip() or "taskkill 已执行"))
            messagebox.showinfo("已处理", "已尝试关闭 Excel。请重新点击刚才的步骤。")
        except Exception as exc:
            self._log(f"关闭 Excel 失败：{exc}")
            messagebox.showwarning("关闭失败", f"没能自动关闭 Excel：{exc}\n请手动关闭相关 Excel 文件后重试。")

    def pick_month_root(self) -> None:
        path = filedialog.askdirectory(title="选择当月结算文件夹")
        if path:
            self.month_root_var.set(path)
            month = self.month_var.get()
            root = Path(path)
            self.output_dir_var.set(str(root / f"{month}月自动化输出"))
            power = root / "零售侧用户电量数据处理表.xlsx"
            if power.exists():
                self.power_var.set(str(power))
            raw = self._find_raw(root)
            if raw:
                self.raw_detail_var.set(str(raw))

    def pick_output_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.output_dir_var.set(path)

    def pick_raw_detail(self) -> None:
        path = filedialog.askopenfilename(title="选择原始零售侧明细", filetypes=[("Excel/CSV", "*.xls *.xlsx *.csv")])
        if path:
            self.raw_detail_var.set(path)

    def pick_power(self) -> None:
        path = filedialog.askopenfilename(title="选择电量处理表", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.power_var.set(path)

    def pick_base_ledger(self) -> None:
        path = filedialog.askopenfilename(title="选择基础台账（上月/修正版）", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.base_ledger_var.set(path)

    def pick_reference_ledger(self) -> None:
        path = filedialog.askopenfilename(title="选择参考台账", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.reference_ledger_var.set(path)

    def pick_completed_ledger(self) -> None:
        path = filedialog.askopenfilename(title="选择人工整理后的台账", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.completed_ledger_var.set(path)

    def pick_proxy_template_dir(self) -> None:
        path = filedialog.askdirectory(title="选择上月代理分表文件夹（2026年代理 - 海南）")
        if path:
            self.proxy_template_dir_var.set(path)

    def pick_inter_template_dir(self) -> None:
        path = filedialog.askdirectory(title="选择上月居间分表文件夹（2026年居间 - 海南）")
        if path:
            self.inter_template_dir_var.set(path)

    def pick_summary_template(self) -> None:
        path = filedialog.askopenfilename(title="选择上月/修正版汇总表", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.summary_template_var.set(path)

    def _find_raw(self, root: Path) -> Path | None:
        for suffix in (".xls", ".xlsx", ".csv"):
            for path in root.glob(f"零售侧明细结果*{suffix}"):
                if "数据处理" not in path.name and "数据清洗" not in path.name:
                    return path
        return None

    def _output_dir(self) -> Path:
        output = as_path(self.output_dir_var.get())
        if output:
            if output.exists() and not output.is_dir():
                raise ValueError(f"输出文件夹位置不是文件夹：{output}")
            return output

        base_dir = as_path(self.month_root_var.get())
        if not base_dir:
            for candidate in (
                as_path(self.raw_detail_var.get()),
                as_path(self.power_var.get()),
                as_path(self.completed_ledger_var.get()),
                as_path(self.base_ledger_var.get()),
            ):
                if candidate:
                    base_dir = candidate.parent
                    break
        if not base_dir:
            raise ValueError("请选择输出文件夹；如果不填，请至少选择一个输入文件或当月结算文件夹。")
        output = base_dir / f"{self.month_var.get()}月自动化输出"
        self.output_dir_var.set(str(output))
        return output

    def _power_path(self) -> Path:
        power = as_path(self.power_var.get())
        if not power:
            raw = as_path(self.raw_detail_var.get())
            month_root = as_path(self.month_root_var.get())
            if raw:
                power = raw.parent / "零售侧用户电量数据处理表.xlsx"
            elif month_root:
                power = month_root / "零售侧用户电量数据处理表.xlsx"
            else:
                raise ValueError("请选择电量处理表，或选择原始零售侧明细让程序自动生成电量处理表。")
            self.power_var.set(str(power))
        return power

    def _prepare_stage1_inputs(self) -> tuple[Path | None, Path]:
        month_root = as_path(self.month_root_var.get())
        raw = as_path(self.raw_detail_var.get())
        power = as_path(self.power_var.get())

        if month_root:
            if not month_root.exists():
                raise FileNotFoundError(f"当月结算文件夹不存在：{month_root}")
            if not month_root.is_dir():
                raise ValueError(f"当月结算文件夹不是文件夹：{month_root}")
        if not raw and month_root:
            raw = self._find_raw(month_root)
            if raw:
                self.raw_detail_var.set(str(raw))
        if not power:
            power = self._power_path()
        if raw:
            self._reject_temp_file(raw, "原始零售侧明细")
            if not raw.exists():
                raise FileNotFoundError(f"原始零售侧明细不存在：{raw}")
            if not raw.is_file():
                raise ValueError(f"原始零售侧明细不是文件：{raw}")
        if power.name.startswith("~$"):
            raise ValueError(f"电量处理表选到了 Excel 临时文件：{power.name}\n请关闭/另选正式 .xlsx 文件。")
        if not power.exists() and not raw:
            raise FileNotFoundError(f"电量处理表不存在，且未选择可用于清洗的原始明细：{power}")
        return raw, power

    def run_clean(self) -> None:
        def task() -> str:
            raw = self._required_file(self.raw_detail_var, "原始零售侧明细")
            power = self._power_path()
            if power.name.startswith("~$"):
                raise ValueError(f"电量处理表选到了 Excel 临时文件：{power.name}\n请关闭/另选正式 .xlsx 文件。")
            result = clean_power_data(raw, power)
            self.queue.put(("log", json.dumps(result, ensure_ascii=False, indent=2)))
            return f"已生成电量处理表：\n{power}"

        self._run_background("清洗电量", task)

    def run_ledger(self) -> None:
        def task() -> str:
            base = self._required_file(self.base_ledger_var, "基础台账（上月/修正版）")
            power = self._power_path()
            if power.name.startswith("~$"):
                raise ValueError(f"电量处理表选到了 Excel 临时文件：{power.name}\n请关闭/另选正式 .xlsx 文件。")
            if not power.exists():
                raise FileNotFoundError(f"电量处理表不存在：{power}\n请先运行“仅清洗电量”，或选择已有电量处理表。")
            raw = self._optional_file(self.raw_detail_var, "原始零售侧明细")
            ref = self._optional_file(self.reference_ledger_var, "参考台账")
            out, report_path, report = update_ledger(
                month=self.month_var.get(),
                base_ledger=base,
                power_path=power,
                output_dir=self._output_dir(),
                raw_detail=raw,
                reference_ledger=ref,
                copy_reference_existing=self.copy_reference_existing_var.get(),
            )
            self.queue.put(("set_completed_ledger", str(out)))
            self.queue.put(("log", json.dumps(report, ensure_ascii=False, indent=2)))
            return f"已生成待整理台账：\n{out}\n\n请人工补齐/检查后，再用它执行阶段2。\n\n报告：\n{report_path}"

        self._run_background("更新台账", task)

    def run_stage1(self) -> None:
        def task() -> str:
            base = self._required_file(self.base_ledger_var, "基础台账（上月/修正版）")
            ref = self._optional_file(self.reference_ledger_var, "参考台账")

            self._progress("阶段1：正在准备电量和台账输入")
            raw, power = self._prepare_stage1_inputs()

            clean_report = None
            if not power.exists():
                if not raw:
                    raise FileNotFoundError("找不到电量处理表，也没有找到可清洗的原始零售侧明细。")
                self._progress("阶段1：正在清洗零售侧明细并生成电量处理表")
                clean_report = clean_power_data(raw, power)
                self.queue.put(("log", json.dumps(clean_report, ensure_ascii=False, indent=2)))
            else:
                self.queue.put(("log", f"已找到电量处理表，跳过清洗：{power}"))

            self._progress("阶段1：正在把电量、新增客户名称和户号导入台账")
            out, report_path, report = update_ledger(
                month=self.month_var.get(),
                base_ledger=base,
                power_path=power,
                output_dir=self._output_dir(),
                raw_detail=raw,
                reference_ledger=ref,
                copy_reference_existing=self.copy_reference_existing_var.get(),
            )
            self.queue.put(("set_completed_ledger", str(out)))
            self.queue.put(("log", json.dumps(report, ensure_ascii=False, indent=2)))

            result = {
                "stage": "stage1",
                "month": self.month_var.get(),
                "power": str(power),
                "raw_detail": str(raw) if raw else None,
                "ledger": str(out),
                "clean_report": clean_report,
                "ledger_report": str(report_path),
            }
            stage_report = self._output_dir() / f"{self.month_var.get()}月阶段1报告.json"
            stage_report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
            return f"阶段1完成，已生成待整理台账：\n{out}\n\n请打开该台账人工补齐负责人、代理/居间、价格、支付方等字段；补齐后再执行阶段2。\n\n报告：\n{stage_report}"

        self._run_background("阶段1 清洗并导入台账", task)

    def run_settlement(self) -> None:
        def task() -> str:
            ledger = self._required_file(self.completed_ledger_var, "人工整理后的台账")
            proxy_template_dir = self._required_dir(self.proxy_template_dir_var, "上月代理分表文件夹（2026年代理 - 海南）")
            inter_template_dir = self._required_dir(self.inter_template_dir_var, "上月居间分表文件夹（2026年居间 - 海南）")
            summary_template = self._required_file(self.summary_template_var, "上月/修正版汇总表")
            self._progress("阶段2：正在读取人工整理后的台账并生成代理/居间分表")
            summary, report_path, report = run_settlement(
                month=self.month_var.get(),
                ledger_path=ledger,
                proxy_template_dir=proxy_template_dir,
                inter_template_dir=inter_template_dir,
                summary_template=summary_template,
                output_dir=self._output_dir(),
                allow_missing_owner=self.allow_missing_owner_var.get(),
                recalc_excel=not self.no_recalc_var.get(),
            )
            self.queue.put(("log", json.dumps(report, ensure_ascii=False, indent=2)))
            return f"已生成汇总表：\n{summary}\n\n报告：\n{report_path}"

        self._run_background("阶段2 生成结算成果", task)

    def open_output_dir(self) -> None:
        output = self._output_dir()
        output.mkdir(parents=True, exist_ok=True)
        import os

        os.startfile(output)


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
