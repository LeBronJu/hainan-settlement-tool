# C# 架构说明

## 分层

### HainanSettlementTool.WinForms

负责桌面界面：

- 文件选择
- 参数输入
- 日志展示
- 调用 `Stage1Service`

禁止在 UI 层写 Excel 读写、客户匹配、金额计算等业务逻辑。

### HainanSettlementTool.Core

负责业务模型和业务服务：

- `Stage1Options`
- `Stage1Report`
- `PowerRow`
- `Stage1Service`
- `IStage1ExcelGateway`

Core 层只表达“要做什么”，不关心 Excel 文件具体怎么读写。

### HainanSettlementTool.Excel

负责 Excel 文件实现：

- `PowerWorkbookReader`：读取/写入电量处理表
- `RawDetailReader`：清洗原始零售侧明细
- `CustomerCodeReader`：读取户号
- `LedgerStage1Updater`：复制基础台账并写入当月电量
- `ClosedXmlStage1ExcelGateway`：组合以上组件，对 Core 暴露统一接口

## 阶段1边界

输入：

- 基础台账
- 电量处理表，或原始零售侧明细
- 可选参考台账
- 输出文件夹

输出：

- 待整理台账
- 阶段1 JSON 报告

阶段1只负责：

- 写入当月电量
- 新增客户名称
- 补户号
- 输出待人工补齐项

阶段1不负责：

- 自动补负责人
- 自动判断代理/居间
- 生成分表
- 生成汇总表

## 迁移策略

1. 阶段1在脱敏样例上与 Python 版对齐。
2. 阶段1在真实工作副本上试跑，但不覆盖正式文件。
3. 迁移阶段2代理/居间分表。
4. 迁移阶段2汇总表。
5. 最后再考虑替换 Python 版或保留为 legacy。
