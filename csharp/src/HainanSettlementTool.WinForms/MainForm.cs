using System;
using System.Drawing;
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
        private readonly ComboBox _month = new ComboBox();
        private readonly TextBox _baseLedger = new TextBox();
        private readonly TextBox _power = new TextBox();
        private readonly TextBox _rawDetail = new TextBox();
        private readonly TextBox _referenceLedger = new TextBox();
        private readonly TextBox _outputDir = new TextBox();
        private readonly CheckBox _copyReferenceExisting = new CheckBox();
        private readonly Button _runStage1 = new Button();
        private readonly TextBox _log = new TextBox();
        private readonly Label _status = new Label();

        private static readonly Color WindowBackground = Color.FromArgb(246, 247, 249);
        private static readonly Color PanelBackground = Color.White;
        private static readonly Color BorderColor = Color.FromArgb(222, 226, 230);
        private static readonly Color MutedText = Color.FromArgb(92, 99, 112);
        private static readonly Color MainText = Color.FromArgb(34, 40, 49);
        private static readonly Color Accent = Color.FromArgb(0, 118, 125);
        private static readonly Color AccentHover = Color.FromArgb(0, 96, 102);

        public MainForm()
        {
            Text = "海南售电结算自动化工具";
            Width = 1040;
            Height = 760;
            MinimumSize = new Size(940, 680);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = WindowBackground;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            AutoScaleMode = AutoScaleMode.Dpi;

            BuildLayout();
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(24, 22, 24, 24),
                BackColor = WindowBackground
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            root.Controls.Add(BuildHeader(), 0, 0);
            root.Controls.Add(BuildBody(), 0, 1);
            root.Controls.Add(BuildLogPanel(), 0, 2);
        }

        private Control BuildHeader()
        {
            var header = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 18)
            };
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var titleBlock = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                AutoSize = true
            };
            titleBlock.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            titleBlock.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            titleBlock.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "海南售电结算自动化工具",
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 18F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 4)
            }, 0, 0);
            titleBlock.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "C# 重构版 · 当前实现阶段一：电量清洗、导入台账、补新增客户名称和户号",
                ForeColor = MutedText,
                Font = new Font(Font.FontFamily, 9.5F),
                Margin = new Padding(1, 0, 0, 0)
            }, 0, 1);

            _status.AutoSize = true;
            _status.Text = "就绪";
            _status.ForeColor = Accent;
            _status.BackColor = Color.FromArgb(230, 246, 247);
            _status.Font = new Font(Font.FontFamily, 9F, FontStyle.Bold);
            _status.Padding = new Padding(14, 7, 14, 7);
            _status.Margin = new Padding(12, 8, 0, 0);
            _status.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            header.Controls.Add(titleBlock, 0, 0);
            header.Controls.Add(_status, 1, 0);
            return header;
        }

        private Control BuildBody()
        {
            var body = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 16)
            };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));

            var inputPanel = CreatePanel();
            inputPanel.Margin = new Padding(0, 0, 10, 0);
            inputPanel.Controls.Add(BuildInputSection());

            var workflowPanel = CreatePanel();
            workflowPanel.Margin = new Padding(10, 0, 0, 0);
            workflowPanel.Controls.Add(BuildWorkflowSection());

            body.Controls.Add(inputPanel, 0, 0);
            body.Controls.Add(workflowPanel, 1, 0);
            return body;
        }

        private Control BuildInputSection()
        {
            var form = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = true,
                Padding = new Padding(20, 18, 20, 20),
                BackColor = PanelBackground
            };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 92));
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            form.Controls.Add(SectionTitle("阶段一输入"), 0, 0);
            form.SetColumnSpan(form.GetControlFromPosition(0, 0), 3);

            _month.DropDownStyle = ComboBoxStyle.DropDownList;
            for (var month = 2; month <= 12; month++)
            {
                _month.Items.Add("2026年" + month + "月");
            }
            _month.SelectedIndex = 3;
            _month.Height = 30;
            AddControlRow(form, "月份", _month, null);
            AddPathRow(form, "基础台账(必填)", _baseLedger, "Excel 文件|*.xlsx");
            AddPathRow(form, "电量处理表", _power, "Excel 文件|*.xlsx");
            AddPathRow(form, "原始零售侧明细", _rawDetail, "Excel/CSV|*.xlsx;*.csv");
            AddPathRow(form, "参考台账(可选)", _referenceLedger, "Excel 文件|*.xlsx");
            AddFolderRow(form, "输出文件夹(必填)", _outputDir);

            _copyReferenceExisting.Text = "用参考台账覆盖已有客户基础资料";
            _copyReferenceExisting.ForeColor = MainText;
            _copyReferenceExisting.AutoSize = true;
            _copyReferenceExisting.Margin = new Padding(153, 12, 0, 12);
            var checkboxRow = AddRow(form);
            form.Controls.Add(_copyReferenceExisting, 0, checkboxRow);
            form.SetColumnSpan(_copyReferenceExisting, 3);

            var actions = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(153, 6, 0, 0)
            };
            _runStage1.Text = "阶段1 清洗并导入台账";
            _runStage1.Width = 190;
            _runStage1.Height = 38;
            StylePrimaryButton(_runStage1);
            _runStage1.Click += async (sender, args) => await RunStage1Async();
            actions.Controls.Add(_runStage1);
            var actionRow = AddRow(form);
            form.Controls.Add(actions, 0, actionRow);
            form.SetColumnSpan(actions, 3);

            return form;
        }

        private Control BuildWorkflowSection()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                AutoSize = true,
                Padding = new Padding(20, 18, 20, 20),
                BackColor = PanelBackground
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            panel.Controls.Add(SectionTitle("操作流程"), 0, 0);
            panel.Controls.Add(WorkflowItem("1", "选择基础台账", "选择已经整理好的上月或当前基础台账。"), 0, 1);
            panel.Controls.Add(WorkflowItem("2", "提供电量来源", "可直接选择电量处理表；没有时选择原始零售侧明细。"), 0, 2);
            panel.Controls.Add(WorkflowItem("3", "确认输出位置", "程序会另存新文件，不覆盖原始表格。"), 0, 3);
            panel.Controls.Add(WorkflowItem("4", "运行阶段一", "完成后查看日志和 JSON 校验报告。"), 0, 4);

            return panel;
        }

        private Control BuildLogPanel()
        {
            var panel = CreatePanel();
            panel.Dock = DockStyle.Fill;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(20, 16, 20, 20),
                BackColor = PanelBackground
            };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            panel.Controls.Add(layout);

            layout.Controls.Add(SectionTitle("运行日志"), 0, 0);

            _log.Multiline = true;
            _log.ScrollBars = ScrollBars.Vertical;
            _log.Dock = DockStyle.Fill;
            _log.ReadOnly = true;
            _log.BorderStyle = BorderStyle.None;
            _log.BackColor = Color.FromArgb(31, 35, 42);
            _log.ForeColor = Color.FromArgb(234, 238, 243);
            _log.Font = new Font("Consolas", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            _log.Margin = new Padding(0, 8, 0, 0);
            layout.Controls.Add(_log, 0, 1);

            return panel;
        }

        private static Panel CreatePanel()
        {
            return new Panel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                BackColor = PanelBackground,
                Padding = new Padding(1),
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        private Label SectionTitle(string text)
        {
            return new Label
            {
                AutoSize = true,
                Text = text,
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 11F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 14)
            };
        }

        private Control WorkflowItem(string number, string title, string detail)
        {
            var row = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 2,
                Margin = new Padding(0, 0, 0, 15),
                BackColor = PanelBackground
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 34));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            var badge = new Label
            {
                Text = number,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.White,
                BackColor = Accent,
                Width = 26,
                Height = 26,
                Font = new Font(Font.FontFamily, 9F, FontStyle.Bold),
                Margin = new Padding(0, 1, 8, 0)
            };

            var text = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                RowCount = 2,
                ColumnCount = 1,
                BackColor = PanelBackground
            };
            text.Controls.Add(new Label
            {
                AutoSize = true,
                Text = title,
                ForeColor = MainText,
                Font = new Font(Font.FontFamily, 9.5F, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, 3)
            }, 0, 0);
            text.Controls.Add(new Label
            {
                AutoSize = true,
                Text = detail,
                ForeColor = MutedText,
                MaximumSize = new Size(260, 0),
                Margin = new Padding(0)
            }, 0, 1);

            row.Controls.Add(badge, 0, 0);
            row.Controls.Add(text, 1, 0);
            return row;
        }

        private static void AddControlRow(TableLayoutPanel form, string label, Control control, Button button)
        {
            var row = AddRow(form);
            form.Controls.Add(new Label
            {
                Text = label,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                ForeColor = MutedText,
                Margin = new Padding(0, 8, 14, 8)
            }, 0, row);
            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(0, 5, 10, 5);
            StyleInput(control);
            form.Controls.Add(control, 1, row);
            if (button != null)
            {
                button.Margin = new Padding(0, 5, 0, 5);
                form.Controls.Add(button, 2, row);
            }
            else
            {
                form.Controls.Add(new Panel { Dock = DockStyle.Fill, Height = 1 }, 2, row);
            }
        }

        private static int AddRow(TableLayoutPanel form)
        {
            var row = form.RowCount;
            form.RowCount = row + 1;
            form.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            return row;
        }

        private static void AddPathRow(TableLayoutPanel form, string label, TextBox textBox, string filter)
        {
            var button = CreateBrowseButton();
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
            var button = CreateBrowseButton();
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

        private static Button CreateBrowseButton()
        {
            var button = new Button
            {
                Text = "浏览",
                Dock = DockStyle.Fill,
                Height = 30
            };
            StyleSecondaryButton(button);
            return button;
        }

        private static void StyleInput(Control control)
        {
            var textBox = control as TextBox;
            if (textBox != null)
            {
                textBox.BorderStyle = BorderStyle.FixedSingle;
                textBox.BackColor = Color.White;
                textBox.ForeColor = MainText;
                textBox.Height = 30;
                return;
            }

            var comboBox = control as ComboBox;
            if (comboBox != null)
            {
                comboBox.FlatStyle = FlatStyle.Flat;
                comboBox.BackColor = Color.White;
                comboBox.ForeColor = MainText;
            }
        }

        private static void StylePrimaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Accent;
            button.ForeColor = Color.White;
            button.Font = new Font(button.Font.FontFamily, 9F, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            button.MouseEnter += (sender, args) => button.BackColor = AccentHover;
            button.MouseLeave += (sender, args) => button.BackColor = Accent;
        }

        private static void StyleSecondaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = BorderColor;
            button.FlatAppearance.BorderSize = 1;
            button.BackColor = Color.FromArgb(250, 251, 252);
            button.ForeColor = MainText;
            button.Cursor = Cursors.Hand;
            button.MouseEnter += (sender, args) => button.BackColor = Color.FromArgb(241, 244, 247);
            button.MouseLeave += (sender, args) => button.BackColor = Color.FromArgb(250, 251, 252);
        }

        private async Task RunStage1Async()
        {
            _runStage1.Enabled = false;
            SetBusy(true);
            try
            {
                var options = CreateOptions();
                Log("开始阶段1，请不要关闭应用窗口。");
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
                SetBusy(false);
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
                Month = _month.SelectedIndex + 2,
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

        private void SetBusy(bool busy)
        {
            _status.Text = busy ? "运行中" : "就绪";
            _status.ForeColor = busy ? Color.FromArgb(131, 82, 0) : Accent;
            _status.BackColor = busy ? Color.FromArgb(255, 243, 217) : Color.FromArgb(230, 246, 247);
        }
    }
}
