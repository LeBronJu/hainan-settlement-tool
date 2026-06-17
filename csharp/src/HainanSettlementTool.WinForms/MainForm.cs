using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using HainanSettlementTool.Core.Models;
using HainanSettlementTool.Core.Services;
using HainanSettlementTool.Excel;

namespace HainanSettlementTool.WinForms
{
    public sealed class MainForm : Form
    {
        private readonly NumericUpDown _month = new NumericUpDown();
        private readonly TextBox _baseLedger = new TextBox();
        private readonly TextBox _power = new TextBox();
        private readonly TextBox _rawDetail = new TextBox();
        private readonly TextBox _referenceLedger = new TextBox();
        private readonly TextBox _outputDir = new TextBox();
        private readonly CheckBox _copyReferenceExisting = new CheckBox();
        private readonly Button _runStage1 = new Button();
        private readonly TextBox _log = new TextBox();

        public MainForm()
        {
            Text = "海南售电结算自动化工具 - C# 重构版";
            Width = 980;
            Height = 720;
            MinimumSize = new System.Drawing.Size(900, 640);
            StartPosition = FormStartPosition.CenterScreen;

            BuildLayout();
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(12)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            var intro = new Label
            {
                AutoSize = true,
                Text = "C# 重构版当前先实现阶段1：清洗/读取电量表，导入基础台账，补新增客户名称和户号。阶段2会在阶段1稳定后迁移。",
                MaximumSize = new System.Drawing.Size(920, 0)
            };
            root.Controls.Add(intro, 0, 0);

            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 3,
                AutoSize = true,
                Padding = new Padding(0, 12, 0, 6)
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 170));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 90));
            root.Controls.Add(form, 0, 1);

            _month.Minimum = 2;
            _month.Maximum = 12;
            _month.Value = 5;
            AddControlRow(form, "月份", _month, null);
            AddPathRow(form, "基础台账(必填)", _baseLedger, "Excel 文件|*.xlsx");
            AddPathRow(form, "电量处理表", _power, "Excel 文件|*.xlsx");
            AddPathRow(form, "原始零售侧明细", _rawDetail, "Excel/CSV|*.xlsx;*.csv");
            AddPathRow(form, "参考台账(可选)", _referenceLedger, "Excel 文件|*.xlsx");
            AddFolderRow(form, "输出文件夹(必填)", _outputDir);

            _copyReferenceExisting.Text = "用参考台账覆盖已有客户基础资料";
            _copyReferenceExisting.AutoSize = true;
            root.Controls.Add(_copyReferenceExisting, 0, 2);

            var actions = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            _runStage1.Text = "阶段1 清洗并导入台账";
            _runStage1.Width = 180;
            _runStage1.Height = 32;
            _runStage1.Click += async (sender, args) => await RunStage1Async();
            actions.Controls.Add(_runStage1);
            root.Controls.Add(actions, 0, 3);

            _log.Multiline = true;
            _log.ScrollBars = ScrollBars.Vertical;
            _log.Dock = DockStyle.Fill;
            _log.ReadOnly = true;
            root.Controls.Add(_log, 0, 4);
        }

        private static void AddControlRow(TableLayoutPanel form, string label, Control control, Button button)
        {
            var row = form.RowCount++;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            form.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
            control.Dock = DockStyle.Fill;
            form.Controls.Add(control, 1, row);
            if (button != null)
            {
                form.Controls.Add(button, 2, row);
            }
        }

        private static void AddPathRow(TableLayoutPanel form, string label, TextBox textBox, string filter)
        {
            var button = new Button { Text = "浏览", Dock = DockStyle.Fill };
            button.Click += (sender, args) =>
            {
                using (var dialog = new OpenFileDialog { Filter = filter })
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.FileName;
                    }
                }
            };
            AddControlRow(form, label, textBox, button);
        }

        private static void AddFolderRow(TableLayoutPanel form, string label, TextBox textBox)
        {
            var button = new Button { Text = "浏览", Dock = DockStyle.Fill };
            button.Click += (sender, args) =>
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        textBox.Text = dialog.SelectedPath;
                    }
                }
            };
            AddControlRow(form, label, textBox, button);
        }

        private async Task RunStage1Async()
        {
            _runStage1.Enabled = false;
            try
            {
                var options = CreateOptions();
                await Task.Run(() =>
                {
                    var service = new Stage1Service(new ClosedXmlStage1ExcelGateway());
                    var report = service.Run(options, LogThreadSafe);
                    LogThreadSafe("阶段1完成。");
                    LogThreadSafe("输出台账：" + report.Output);
                    LogThreadSafe("报告：" + report.ReportPath);
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "出错了", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Log(ex.ToString());
            }
            finally
            {
                _runStage1.Enabled = true;
            }
        }

        private Stage1Options CreateOptions()
        {
            var powerPath = _power.Text.Trim();
            if (string.IsNullOrWhiteSpace(powerPath) && !string.IsNullOrWhiteSpace(_rawDetail.Text))
            {
                powerPath = Path.Combine(Path.GetDirectoryName(_rawDetail.Text) ?? string.Empty, "零售侧用户电量数据处理表.xlsx");
                _power.Text = powerPath;
            }

            return new Stage1Options
            {
                Month = (int)_month.Value,
                BaseLedgerPath = _baseLedger.Text.Trim(),
                PowerPath = powerPath,
                RawDetailPath = _rawDetail.Text.Trim(),
                ReferenceLedgerPath = _referenceLedger.Text.Trim(),
                OutputDirectory = _outputDir.Text.Trim(),
                CopyReferenceExisting = _copyReferenceExisting.Checked
            };
        }

        private void LogThreadSafe(string message)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(Log), message);
                return;
            }
            Log(message);
        }

        private void Log(string message)
        {
            _log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine);
        }
    }
}
