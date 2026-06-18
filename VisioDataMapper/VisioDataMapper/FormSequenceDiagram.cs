using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormSequenceDiagram : Form
    {
        private Button btnImportTable;
        private Label lblTitle;
        private TextBox txtTitle;
        private Panel pnlTitleBorder;
        
        private TabControl tabControlMain;
        private TabPage tabGrid;
        private TabPage tabCode;
        private DataGridView dgvMessages;
        private TextBox txtMermaidInput;
        
        private CheckedListBox chkActors;
        private Label lblActorsTip;
        
        private TextBox txtStatus;
        private Panel pnlStatusBorder;
        
        private TextBox txtAltNote;
        private Panel pnlAltNoteBorder;
        
        private Panel pnlOptions;
        private Label lblHorSpacing;
        private TextBox txtHorizontalSpacing;
        private Label lblObjSize;
        private TextBox txtObjectWidth;
        private TextBox txtObjectHeight;
        private CheckBox chkDrawActiveRect;
        private CheckBox chkDrawBottomShape;
        private CheckBox chkActorStickman;
        private CheckBox chkRightToLeftDashed;
        private ComboBox cmbActiveRectStyle;
        private Label lblActiveRectStyle;
        private CheckBox chkDrawStageNotes;
        private Label lblStageSpacing;
        private TextBox txtStageSpacing;
        
        private Button btnGenerate;
        private Button btnClose;

        private string currentFontName = "微软雅黑";
        private double currentFontSize = 9.0;
        private double currentLineWidth = 0.75;
        private bool isInternalTextChange = false;

        public FormSequenceDiagram()
        {
            InitializeComponent();
            LoadSampleData();
        }

        private void InitializeComponent()
        {
            this.Text = "智能画图-UML时序图";
            this.Size = new Size(950, 850);
            this.MinimumSize = new Size(950, 780);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;
            this.BackColor = Color.FromArgb(240, 244, 248);
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);

            // Import Button
            btnImportTable = new Button
            {
                Text = "导入表格",
                Location = new Point(15, 15),
                Size = new Size(130, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(102, 45, 145),
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnImportTable.FlatAppearance.BorderColor = Color.FromArgb(102, 45, 145);
            btnImportTable.FlatAppearance.BorderSize = 1;
            btnImportTable.Click += btnImportTable_Click;

            // Title Label
            lblTitle = new Label
            {
                Text = "大标题:",
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104),
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Regular)
            };

            // Title TextBox border panel
            pnlTitleBorder = new Panel
            {
                BackColor = Color.FromArgb(218, 224, 233),
                Padding = new Padding(1),
                Size = new Size(340, 28)
            };
            txtTitle = new TextBox
            {
                Text = "用户充值话费时序图",
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 10F, FontStyle.Regular),
                BackColor = Color.White
            };
            pnlTitleBorder.Controls.Add(txtTitle);

            // TabControl for Data and Code
            tabControlMain = new TabControl
            {
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular),
                DrawMode = TabDrawMode.OwnerDrawFixed,
                Padding = new Point(15, 6)
            };
            tabControlMain.DrawItem += TabControlMain_DrawItem;

            tabGrid = new TabPage("时序图数据");
            tabGrid.BackColor = Color.White;
            dgvMessages = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 28 },
                BackgroundColor = Color.White
            };
            dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { Name = "Index", HeaderText = "序号", FillWeight = 40 });
            dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { Name = "Text", HeaderText = "消息内容", FillWeight = 200 });
            dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { Name = "Source", HeaderText = "起点", FillWeight = 90 });
            dgvMessages.Columns.Add(new DataGridViewTextBoxColumn { Name = "Target", HeaderText = "终点", FillWeight = 90 });

            var cmbTypeCol = new DataGridViewComboBoxColumn { Name = "Type", HeaderText = "消息类型", FillWeight = 90 };
            cmbTypeCol.Items.AddRange("正常", "异步", "返回", "自关联");
            dgvMessages.Columns.Add(cmbTypeCol);
            
            ApplyGridStyle(dgvMessages);
            tabGrid.Controls.Add(dgvMessages);

            tabCode = new TabPage("Mermaid代码");
            tabCode.BackColor = Color.White;
            txtMermaidInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F, FontStyle.Regular),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            txtMermaidInput.TextChanged += TxtMermaidInput_TextChanged;
            tabCode.Controls.Add(txtMermaidInput);

            tabControlMain.TabPages.Add(tabGrid);
            tabControlMain.TabPages.Add(tabCode);

            // CheckedListBox on the right for actors
            chkActors = new CheckedListBox
            {
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular),
                CheckOnClick = true
            };
            lblActorsTip = new Label
            {
                Text = "请勾选是角色的项目",
                ForeColor = Color.FromArgb(100, 116, 139),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular),
                AutoSize = true
            };

            // Log Textbox
            pnlStatusBorder = new Panel
            {
                BackColor = Color.FromArgb(218, 224, 233),
                Padding = new Padding(1)
            };
            txtStatus = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.FromArgb(248, 250, 252),
                ForeColor = Color.FromArgb(100, 116, 139),
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 9F, FontStyle.Regular)
            };
            pnlStatusBorder.Controls.Add(txtStatus);

            // Alt Rules / Phase Notes Textbox
            pnlAltNoteBorder = new Panel
            {
                BackColor = Color.FromArgb(218, 224, 233),
                Padding = new Padding(1)
            };
            txtAltNote = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(45, 55, 72),
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Regular)
            };
            pnlAltNoteBorder.Controls.Add(txtAltNote);

            // Options Panel
            pnlOptions = new Panel
            {
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Height = 150
            };
            pnlOptions.Paint += (s, pe) =>
            {
                using (var pen = new Pen(Color.FromArgb(218, 224, 233), 1))
                {
                    pe.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    pe.Graphics.DrawRectangle(pen, 0, 0, pnlOptions.Width - 1, pnlOptions.Height - 1);
                }
            };

            lblHorSpacing = new Label { Text = "生命线横向间距:", AutoSize = true, ForeColor = Color.FromArgb(74, 85, 104) };
            txtHorizontalSpacing = new TextBox { Text = "50", Size = new Size(50, 23), TextAlign = HorizontalAlignment.Center };
            
            lblObjSize = new Label { Text = "对象名称宽高:", AutoSize = true, ForeColor = Color.FromArgb(74, 85, 104) };
            txtObjectWidth = new TextBox { Text = "20", Size = new Size(35, 23), TextAlign = HorizontalAlignment.Center };
            txtObjectHeight = new TextBox { Text = "10", Size = new Size(35, 23), TextAlign = HorizontalAlignment.Center };

            chkDrawActiveRect = new CheckBox { Text = "生命线上是否画细长矩形", AutoSize = true, Checked = true, ForeColor = Color.FromArgb(74, 85, 104) };
            chkDrawBottomShape = new CheckBox { Text = "底部是否也画上对象形状", AutoSize = true, Checked = false, ForeColor = Color.FromArgb(74, 85, 104) };
            chkActorStickman = new CheckBox { Text = "角色使用线条小人", AutoSize = true, Checked = true, ForeColor = Color.FromArgb(74, 85, 104) };
            chkRightToLeftDashed = new CheckBox { Text = "从右往左线条一律虚线", AutoSize = true, Checked = false, ForeColor = Color.FromArgb(74, 85, 104) };

            lblActiveRectStyle = new Label { Text = "激活矩形头尾：", AutoSize = true, ForeColor = Color.FromArgb(74, 85, 104) };
            cmbActiveRectStyle = new ComboBox { Size = new Size(110, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbActiveRectStyle.Items.AddRange(new string[] { "默认", "多出一格空距" });
            cmbActiveRectStyle.SelectedIndex = 1; // 多出一格空距

            chkDrawStageNotes = new CheckBox { Text = "是否生成阶段图形", AutoSize = true, Checked = true, ForeColor = Color.FromArgb(74, 85, 104) };
            lblStageSpacing = new Label { Text = "阶段额外间距(mm):", AutoSize = true, ForeColor = Color.FromArgb(74, 85, 104) };
            txtStageSpacing = new TextBox { Text = "5", Size = new Size(40, 23), TextAlign = HorizontalAlignment.Center };

            btnGenerate = new Button
            {
                Text = "生成绘图",
                Size = new Size(130, 38),
                BackColor = Color.FromArgb(102, 45, 145), // Purple theme color
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 10F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnGenerate.FlatAppearance.BorderSize = 0;
            btnGenerate.Click += btnGenerate_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Size = new Size(100, 38),
                BackColor = Color.FromArgb(233, 236, 239),
                ForeColor = Color.FromArgb(73, 80, 87),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += btnClose_Click;

            // Place option controls in panel
            FlowLayoutPanel optLayout = new FlowLayoutPanel
            {
                Location = new Point(10, 10),
                Size = new Size(650, 130),
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(0)
            };

            FlowLayoutPanel row1 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 5) };
            row1.Controls.Add(lblHorSpacing);
            row1.Controls.Add(txtHorizontalSpacing);
            Label spacer1 = new Label { Width = 20, AutoSize = false };
            row1.Controls.Add(spacer1);
            row1.Controls.Add(lblObjSize);
            row1.Controls.Add(txtObjectWidth);
            row1.Controls.Add(txtObjectHeight);

            FlowLayoutPanel row2 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 5) };
            row2.Controls.Add(chkDrawActiveRect);
            Label spacer2 = new Label { Width = 20, AutoSize = false };
            row2.Controls.Add(spacer2);
            row2.Controls.Add(lblActiveRectStyle);
            row2.Controls.Add(cmbActiveRectStyle);

            FlowLayoutPanel row3 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 5) };
            row3.Controls.Add(chkDrawStageNotes);
            Label spacer3 = new Label { Width = 20, AutoSize = false };
            row3.Controls.Add(spacer3);
            row3.Controls.Add(lblStageSpacing);
            row3.Controls.Add(txtStageSpacing);

            FlowLayoutPanel row4 = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0, 0, 0, 5) };
            row4.Controls.Add(chkDrawBottomShape);
            Label spacer4 = new Label { Width = 20, AutoSize = false };
            row4.Controls.Add(spacer4);
            row4.Controls.Add(chkActorStickman);
            Label spacer5 = new Label { Width = 20, AutoSize = false };
            row4.Controls.Add(spacer5);
            row4.Controls.Add(chkRightToLeftDashed);

            optLayout.Controls.Add(row1);
            optLayout.SetFlowBreak(row1, true);
            optLayout.Controls.Add(row2);
            optLayout.SetFlowBreak(row2, true);
            optLayout.Controls.Add(row3);
            optLayout.SetFlowBreak(row3, true);
            optLayout.Controls.Add(row4);

            pnlOptions.Controls.Add(optLayout);
            pnlOptions.Controls.Add(btnGenerate);
            pnlOptions.Controls.Add(btnClose);

            // Add all controls to Form
            this.Controls.Add(btnImportTable);
            this.Controls.Add(lblTitle);
            this.Controls.Add(pnlTitleBorder);
            this.Controls.Add(tabControlMain);
            this.Controls.Add(chkActors);
            this.Controls.Add(lblActorsTip);
            this.Controls.Add(pnlStatusBorder);
            this.Controls.Add(pnlAltNoteBorder);
            this.Controls.Add(pnlOptions);

            this.Resize += FormSequenceDiagram_Resize;
            LayoutControls();
            
            AppendLog("已启动智能画图-UML时序图。请点击“导入表格”读取剪贴板中的Mermaid时序图代码。");
        }

        private void TabControlMain_DrawItem(object sender, DrawItemEventArgs e)
        {
            var tabCtrl = (TabControl)sender;
            var tabPage = tabCtrl.TabPages[e.Index];
            var tabRect = tabCtrl.GetTabRect(e.Index);
            tabRect.Inflate(2, 2);

            bool isSelected = tabCtrl.SelectedIndex == e.Index;
            using (var bgBrush = new SolidBrush(isSelected ? Color.White : Color.FromArgb(226, 232, 240)))
            {
                e.Graphics.FillRectangle(bgBrush, tabRect);
            }

            using (var textBrush = new SolidBrush(isSelected ? Color.FromArgb(102, 45, 145) : Color.FromArgb(100, 116, 139)))
            {
                string text = tabPage.Text;
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                using (var font = new Font(tabCtrl.Font, isSelected ? FontStyle.Bold : FontStyle.Regular))
                {
                    e.Graphics.DrawString(text, font, textBrush, tabRect, sf);
                }
            }
        }

        private void ApplyGridStyle(DataGridView grid)
        {
            grid.BorderStyle = BorderStyle.None;
            grid.BackgroundColor = Color.White;
            grid.GridColor = Color.FromArgb(230, 235, 240);
            grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 248);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(73, 80, 87);
            grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(240, 244, 248);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold);
            grid.EnableHeadersVisualStyles = false;

            grid.DefaultCellStyle.BackColor = Color.White;
            grid.DefaultCellStyle.ForeColor = Color.FromArgb(45, 55, 72);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(230, 242, 255);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(0, 122, 255);
            grid.DefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
        }

        private void LayoutControls()
        {
            int margin = 15;
            int width = Math.Max(900, this.ClientSize.Width - margin * 2);
            int rightActorsWidth = 160;

            // Title alignment
            pnlTitleBorder.Location = new Point(Math.Max(500, margin + width - pnlTitleBorder.Width - rightActorsWidth - margin), 19);
            lblTitle.Location = new Point(pnlTitleBorder.Left - lblTitle.Width - 5, 23);

            // TabControl placement
            tabControlMain.Location = new Point(margin, 65);
            tabControlMain.Width = width - rightActorsWidth - margin;
            tabControlMain.Height = Math.Max(200, this.ClientSize.Height - 65 - 150 - 80 - 15 - margin);

            // CheckedListBox on the right side of tabControl
            chkActors.Location = new Point(tabControlMain.Right + margin, 85);
            chkActors.Width = rightActorsWidth - margin;
            chkActors.Height = tabControlMain.Height - 30;
            lblActorsTip.Location = new Point(chkActors.Left, chkActors.Bottom + 5);

            // Alt / note block text box
            pnlAltNoteBorder.Location = new Point(margin, tabControlMain.Bottom + 10);
            pnlAltNoteBorder.Width = (int)(tabControlMain.Width * 0.55);
            pnlAltNoteBorder.Height = 70;

            // Status logs block
            pnlStatusBorder.Location = new Point(pnlAltNoteBorder.Right + 10, tabControlMain.Bottom + 10);
            pnlStatusBorder.Width = tabControlMain.Right - pnlStatusBorder.Left;
            pnlStatusBorder.Height = 70;

            // Options panel placement
            pnlOptions.Location = new Point(margin, pnlAltNoteBorder.Bottom + 10);
            pnlOptions.Width = width;

            // Close and Generate buttons in panel
            btnClose.Location = new Point(pnlOptions.Width - btnClose.Width - 15, 100);
            btnGenerate.Location = new Point(btnClose.Left - btnGenerate.Width - 15, 100);
        }

        private void FormSequenceDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void LoadSampleData()
        {
            txtMermaidInput.Text = @"sequenceDiagram
    actor 用户
    participant 充值前端
    participant 充值后端
    participant 运营商接口
    participant 用户账户

    Note over 用户, 运营商接口: 阶段一：充值请求发起
    用户->>充值前端: 选择充值金额并提交
    充值前端->>充值前端: 校验金额及登录状态
    充值前端->>充值后端: 发起充值请求（金额、手机号）

    Note over 充值后端, 用户账户: 阶段二：订单创建与支付
    充值后端->>充值后端: 生成充值订单
    充值后端->>用户账户: 扣减用户余额
    用户账户-->>充值后端: 返回扣款结果

    alt 扣款成功
        充值后端->>充值后端: 更新订单状态为“已支付”
    else 扣款失败
        充值后端-->>充值前端: 返回余额不足
        充值前端-->>用户: 提示余额不足，流程终止
    end

    Note over 充值后端, 运营商接口: 阶段三：调用运营商充值
    充值后端->>运营商接口: 调用充值接口（手机号、金额）
    运营商接口-->>充值后端: 返回充值结果（成功/失败）

    Note over 充值后端, 用户: 阶段四：结果处理与通知
    alt 充值成功
        充值后端->>充值后端: 更新订单状态为“成功”
        充值后端->>用户账户: 记录充值流水
        充值后端-->>充值前端: 返回成功结果
        充值前端-->>用户: 显示充值成功
    else 充值失败
        充值后端->>充值后端: 更新订单状态为“失败”
        充值后端->>用户账户: 执行退款（补回余额）
        充值后端-->>充值前端: 返回失败结果
        充值前端-->>用户: 显示充值失败及退款信息
    end

    Note over 用户, 充值后端: 阶段五：流程结束";
        }

        private void btnImportTable_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    MessageBox.Show("剪贴板中没有文本内容！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string clipboardText = Clipboard.GetText().Trim();
                if (clipboardText.Contains("sequenceDiagram") || clipboardText.Contains("->") || clipboardText.Contains("participant"))
                {
                    isInternalTextChange = true;
                    try
                    {
                        txtMermaidInput.Text = clipboardText;
                    }
                    finally
                    {
                        isInternalTextChange = false;
                    }
                    ParseMermaid(clipboardText);
                }
                else
                {
                    // Attempt to parse tabular format
                    ParseTabular(clipboardText);
                }
            }
            catch (Exception ex)
            {
                AppendLog($"导入失败: {ex.Message}");
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TxtMermaidInput_TextChanged(object sender, EventArgs e)
        {
            if (isInternalTextChange) return;

            string text = txtMermaidInput.Text.Trim();
            if (text.Contains("sequenceDiagram"))
            {
                try
                {
                    ParseMermaid(text);
                }
                catch
                {
                    // Ignore transient parsing errors while typing
                }
            }
        }

        private void ParseMermaid(string mermaid)
        {
            dgvMessages.Rows.Clear();
            chkActors.Items.Clear();
            txtAltNote.Clear();

            var lines = mermaid.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var participants = new List<string>();
            var actors = new HashSet<string>();
            var messages = new List<MessageItem>();
            var notes = new List<NoteItem>();
            var altBlocks = new List<AltBlock>();
            var altStack = new Stack<AltBlock>();

            int msgCount = 0;

            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line) || line.StartsWith("%%")) continue;

                if (line.StartsWith("sequenceDiagram", StringComparison.OrdinalIgnoreCase)) continue;

                // 1. Participant/Actor declaration
                var declMatch = Regex.Match(line, @"^\s*(actor|participant)\s+""?([^""\s]+)""?(?:\s+as\s+(\S+))?", RegexOptions.IgnoreCase);
                if (declMatch.Success)
                {
                    string type = declMatch.Groups[1].Value.ToLower();
                    string name = declMatch.Groups[2].Value;
                    if (!participants.Contains(name))
                    {
                        participants.Add(name);
                    }
                    if (type == "actor")
                    {
                        actors.Add(name);
                    }
                    continue;
                }

                // 2. Note declaration
                if (line.StartsWith("Note over", StringComparison.OrdinalIgnoreCase))
                {
                    int colonIdx = line.IndexOf(':');
                    if (colonIdx > 0)
                    {
                        string noteText = line.Substring(colonIdx + 1).Trim();
                        string header = line.Substring(0, colonIdx).Trim();
                        string actorsStr = header.Substring(9).Trim();
                        var noteActors = actorsStr.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                                  .Select(a => a.Trim())
                                                  .ToList();
                        
                        notes.Add(new NoteItem
                        {
                            Actors = noteActors,
                            Text = noteText,
                            NextMessageIndex = msgCount + 1
                        });
                    }
                    continue;
                }

                // 3. Alt / Else / End
                if (line.StartsWith("alt ", StringComparison.OrdinalIgnoreCase))
                {
                    string cond = line.Substring(4).Trim();
                    var block = new AltBlock
                    {
                        AltCondition = cond,
                        AltStartIdx = msgCount + 1
                    };
                    altStack.Push(block);
                    continue;
                }
                if (line.StartsWith("else", StringComparison.OrdinalIgnoreCase))
                {
                    string cond = line.Length > 4 ? line.Substring(4).Trim() : "";
                    if (altStack.Count > 0)
                    {
                        altStack.Peek().ElseBranches.Add(new ElseBranch
                        {
                            Condition = cond,
                            StartIdx = msgCount + 1
                        });
                    }
                    continue;
                }
                if (line.Equals("end", StringComparison.OrdinalIgnoreCase))
                {
                    if (altStack.Count > 0)
                    {
                        var block = altStack.Pop();
                        block.EndIdx = msgCount;
                        altBlocks.Add(block);
                    }
                    continue;
                }

                // 4. Message parsing
                int colon = line.IndexOf(':');
                if (colon > 0)
                {
                    string left = line.Substring(0, colon).Trim();
                    string msgText = line.Substring(colon + 1).Trim();

                    string arrow = "";
                    if (left.Contains("-->>")) arrow = "-->>";
                    else if (left.Contains("-->")) arrow = "-->";
                    else if (left.Contains("->>")) arrow = "->>";
                    else if (left.Contains("->")) arrow = "->";

                    if (!string.IsNullOrEmpty(arrow))
                    {
                        int arrowIdx = left.IndexOf(arrow);
                        string source = left.Substring(0, arrowIdx).Trim();
                        string target = left.Substring(arrowIdx + arrow.Length).Trim();

                        source = source.Trim('"');
                        target = target.Trim('"');

                        if (!string.IsNullOrEmpty(source) && !string.IsNullOrEmpty(target))
                        {
                            msgCount++;

                            if (!participants.Contains(source)) participants.Add(source);
                            if (!participants.Contains(target)) participants.Add(target);

                            string msgType = "正常";
                            if (source == target)
                            {
                                msgType = "自关联";
                            }
                            else if (arrow == "-->>" || arrow == "-->")
                            {
                                msgType = "返回";
                            }
                            else if (arrow == "->")
                            {
                                msgType = "异步";
                            }

                            messages.Add(new MessageItem
                            {
                                Index = msgCount,
                                Source = source,
                                Target = target,
                                Text = msgText,
                                Type = msgType
                            });
                        }
                    }
                }
            }

            // Bind to CheckedListBox
            foreach (var p in participants)
            {
                chkActors.Items.Add(p, actors.Contains(p));
            }

            // Bind to DataGridView
            foreach (var msg in messages)
            {
                dgvMessages.Rows.Add(msg.Index, msg.Text, msg.Source, msg.Target, msg.Type);
            }

            // Output Rules
            var ruleLines = new List<string>();
            foreach (var block in altBlocks.OrderBy(b => b.AltStartIdx))
            {
                string rule = $"alt {block.AltCondition}, {block.AltStartIdx}";
                foreach (var branch in block.ElseBranches)
                {
                    rule += $" else {branch.Condition}, {branch.StartIdx}";
                }
                rule += $" end, {block.EndIdx}";
                ruleLines.Add(rule);
            }

            if (ruleLines.Count > 0)
            {
                ruleLines.Add("");
            }

            foreach (var note in notes.OrderBy(n => n.NextMessageIndex))
            {
                string actorsJoint = string.Join(", ", note.Actors);
                string rule = $"阶段:{actorsJoint}:{note.Text}, {note.NextMessageIndex}";
                ruleLines.Add(rule);
            }

            txtAltNote.Text = string.Join(Environment.NewLine, ruleLines);

            AppendLog($"已成功添加时序图数据 {msgCount}步");
        }

        private void ParseTabular(string text)
        {
            dgvMessages.Rows.Clear();
            chkActors.Items.Clear();
            txtAltNote.Clear();

            string[] rows = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int idx = 1;
            var participants = new HashSet<string>();

            foreach (var row in rows)
            {
                string[] cells = row.Split('\t');
                if (cells.Length >= 3)
                {
                    string content = cells[0].Trim();
                    string src = cells[1].Trim();
                    string dest = cells[2].Trim();
                    string type = cells.Length >= 4 ? cells[3].Trim() : "正常";

                    if (!string.IsNullOrEmpty(src)) participants.Add(src);
                    if (!string.IsNullOrEmpty(dest)) participants.Add(dest);

                    dgvMessages.Rows.Add(idx++, content, src, dest, type);
                }
            }

            foreach (var p in participants)
            {
                chkActors.Items.Add(p, p.Contains("用户") || p.Contains("User") || p.Contains("Client"));
            }

            AppendLog($"已从表格格式导入数据 {idx - 1}步");
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                AppendLog("开始绘制时序图...");
                RenderDiagram();
                AppendLog("时序图生成完成！");
                MessageBox.Show("时序图生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog($"生成失败: {ex.Message}");
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderDiagram()
        {
            // 1. Gather all participants in horizontal order
            var participants = new List<string>();
            var checkedActors = new HashSet<string>();
            foreach (var item in chkActors.Items)
            {
                string pName = item.ToString();
                participants.Add(pName);
                if (chkActors.CheckedItems.Contains(item))
                {
                    checkedActors.Add(pName);
                }
            }

            if (participants.Count == 0)
            {
                throw new InvalidOperationException("未找到参与者对象，请先导入数据。");
            }

            // 2. Read messages from DataGridView
            var messages = new List<MessageItem>();
            foreach (DataGridViewRow row in dgvMessages.Rows)
            {
                if (row.IsNewRow) continue;
                
                int index = Convert.ToInt32(row.Cells["Index"].Value);
                string text = Convert.ToString(row.Cells["Text"].Value);
                string source = Convert.ToString(row.Cells["Source"].Value);
                string target = Convert.ToString(row.Cells["Target"].Value);
                string type = Convert.ToString(row.Cells["Type"].Value);

                messages.Add(new MessageItem
                {
                    Index = index,
                    Text = text,
                    Source = source,
                    Target = target,
                    Type = type
                });
            }

            // 3. Parse Alt Rules and Stage Notes
            var altBlocks = new List<AltBlock>();
            var stageNotes = new List<NoteItem>();
            ParseAltNoteTextBox(txtAltNote.Text, altBlocks, stageNotes);

            // 4. Initialize Visio Application
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            Visio.Page page = GetOrCreateActivePage(visioApp);
            ClearPageShapes(page);

            // 5. Layout math
            double horizSpacing = ParsePositiveNumber(txtHorizontalSpacing.Text, "生命线横向间距") / 25.4;
            double objWidth = ParsePositiveNumber(txtObjectWidth.Text, "对象宽度") / 25.4;
            double objHeight = ParsePositiveNumber(txtObjectHeight.Text, "对象高度") / 25.4;

            double pageWidth = page.PageSheet.CellsU["PageWidth"].Result["in"];
            double pageHeight = page.PageSheet.CellsU["PageHeight"].Result["in"];

            // Calculate auto width/height if needed
            double totalActorsWidth = (participants.Count - 1) * horizSpacing;
            double neededPageWidth = totalActorsWidth + 2.0;
            if (neededPageWidth > pageWidth)
            {
                page.PageSheet.CellsU["PageWidth"].FormulaU = $"{neededPageWidth.ToString(CultureInfo.InvariantCulture)} in";
                pageWidth = neededPageWidth;
            }

            // We estimate page height based on number of messages, alts, and notes.
            double verticalOffsetPerMsg = 0.5; // space for message
            double currentY = pageHeight - 1.2; // Start Y for top shapes
            double topY = currentY;

            // Calculate message offsets in advance
            var messageYCoordinates = new Dictionary<int, double>();
            double tempY = topY - 0.6; // initial margin below header

            // Group notes and alts by index to calculate correct vertical spacing
            var notesByIndex = stageNotes.GroupBy(n => n.NextMessageIndex).ToDictionary(g => g.Key, g => g.ToList());
            var altsByStartIdx = altBlocks.ToDictionary(a => a.AltStartIdx, a => a);
            var altsByEndIdx = altBlocks.GroupBy(a => a.EndIdx).ToDictionary(g => g.Key, g => g.ToList());

            // Parse stage spacing parameter
            double stageSpacingVal = 0;
            if (stageNotes.Count > 0)
            {
                double spacingNum;
                string spacingStr = (txtStageSpacing.Text ?? "5").Trim().Replace('，', '.');
                if (double.TryParse(spacingStr, out spacingNum) && spacingNum >= 0)
                {
                    stageSpacingVal = spacingNum / 25.4; // convert mm to inches
                }
            }

            for (int i = 1; i <= messages.Count + 1; i++)
            {
                // Note check
                if (notesByIndex.ContainsKey(i))
                {
                    tempY -= stageSpacingVal;
                    if (chkDrawStageNotes.Checked)
                    {
                        tempY -= 0.5 * notesByIndex[i].Count;
                    }
                }

                // Alt box start
                if (altsByStartIdx.ContainsKey(i))
                {
                    tempY -= 0.25;
                }

                if (i <= messages.Count)
                {
                    var msg = messages[i - 1];
                    if (msg.Type == "自关联")
                    {
                        tempY -= 0.25;
                    }
                    messageYCoordinates[i] = tempY;
                    tempY -= verticalOffsetPerMsg;
                }

                // Alt box end
                if (altsByEndIdx.ContainsKey(i - 1))
                {
                    tempY -= 0.25;
                }
            }

            double bottomY = tempY - 0.4;
            if (bottomY < 1.0)
            {
                // Auto resize page height to avoid cutting off bottom
                double heightDiff = 1.0 - bottomY;
                double newPageHeight = pageHeight + heightDiff;
                page.PageSheet.CellsU["PageHeight"].FormulaU = $"{newPageHeight.ToString(CultureInfo.InvariantCulture)} in";
                
                // Shift all calculations up by heightDiff
                topY += heightDiff;
                for (int i = 1; i <= messages.Count; i++)
                {
                    messageYCoordinates[i] += heightDiff;
                }
                bottomY = 1.0;
            }

            // Draw 大标题
            string titleText = string.IsNullOrWhiteSpace(txtTitle.Text) ? "时序图" : txtTitle.Text.Trim();
            Visio.Shape titleShape = page.DrawRectangle(pageWidth / 2.0 - 1.5, topY + 0.6, pageWidth / 2.0 + 1.5, topY + 1.0);
            titleShape.Text = titleText;
            ApplyTextShapeStyle(titleShape, currentFontName, currentFontSize + 2, true);
            ApplyBoxStyle(titleShape);

            // 6. Draw Lifelines
            double startX = (pageWidth - totalActorsWidth) / 2.0;
            if (startX < 0.8) startX = 0.8;

            var lifelineXCoords = new Dictionary<string, double>();
            for (int i = 0; i < participants.Count; i++)
            {
                string pName = participants[i];
                double x = startX + i * horizSpacing;
                lifelineXCoords[pName] = x;

                bool isStickman = checkedActors.Contains(pName) && chkActorStickman.Checked;

                // Draw Top Shape
                Visio.Shape topShape;
                if (isStickman)
                {
                    topShape = DrawStickActor(page, pName, x, topY + 0.25);
                }
                else
                {
                    topShape = page.DrawRectangle(x - objWidth / 2.0, topY - objHeight / 2.0, x + objWidth / 2.0, topY + objHeight / 2.0);
                    topShape.Text = pName;
                    ApplyTextShapeStyle(topShape, currentFontName, currentFontSize, false);
                    ApplyBoxStyle(topShape);
                }

                // Draw Bottom Shape if checked
                if (chkDrawBottomShape.Checked)
                {
                    if (isStickman)
                    {
                        DrawStickActor(page, pName, x, bottomY - 0.5);
                    }
                    else
                    {
                        Visio.Shape botShape = page.DrawRectangle(x - objWidth / 2.0, bottomY - objHeight / 2.0, x + objWidth / 2.0, bottomY + objHeight / 2.0);
                        botShape.Text = pName;
                        ApplyTextShapeStyle(botShape, currentFontName, currentFontSize, false);
                        ApplyBoxStyle(botShape);
                    }
                }

                // Draw vertical lifeline line
                double lifeTopY = topY - (isStickman ? 0.3 : objHeight / 2.0);
                double lifeBotY = bottomY + (chkDrawBottomShape.Checked ? (isStickman ? 0.3 : objHeight / 2.0) : 0);
                Visio.Shape lifeline = page.DrawLine(x, lifeTopY, x, lifeBotY);
                TrySetFormula(lifeline, "LineColor", "RGB(120, 120, 120)");
                TrySetFormula(lifeline, "LinePattern", "2"); // Dashed
                TrySetFormula(lifeline, "LineWeight", "0.75pt");
            }

            // 7. Calculate Activation Rects
            if (chkDrawActiveRect.Checked)
            {
                double rectW = 0.12; // 3mm
                bool extraGap = cmbActiveRectStyle.SelectedItem != null && cmbActiveRectStyle.SelectedItem.ToString() == "多出一格空距";

                foreach (var p in participants)
                {
                    // Find messages involving this participant
                    var activeMsgs = messages.Where(m => m.Source == p || m.Target == p).ToList();
                    if (activeMsgs.Count > 0)
                    {
                        int minIdx = activeMsgs.Min(m => m.Index);
                        int maxIdx = activeMsgs.Max(m => m.Index);

                        double activeStartY = messageYCoordinates[minIdx];
                        double activeEndY = messageYCoordinates[maxIdx];

                        if (extraGap)
                        {
                            activeStartY += 0.12;
                            activeEndY -= 0.12;
                        }

                        double x = lifelineXCoords[p];
                        Visio.Shape activeRect = page.DrawRectangle(x - rectW / 2.0, activeEndY, x + rectW / 2.0, activeStartY);
                        TrySetFormula(activeRect, "FillPattern", "1");
                        TrySetFormula(activeRect, "FillForegnd", "RGB(255, 255, 255)");
                        TrySetFormula(activeRect, "LineColor", "RGB(0, 0, 0)");
                        TrySetFormula(activeRect, "LineWeight", "0.75pt");
                    }
                }
            }

            // 8. Draw Stage Notes
            if (chkDrawStageNotes.Checked)
            {
                foreach (var note in stageNotes)
                {
                    int nextIdx = note.NextMessageIndex;
                    double noteY;
                    if (nextIdx <= messages.Count)
                    {
                        noteY = messageYCoordinates[nextIdx] + 0.3;
                    }
                    else
                    {
                        noteY = bottomY + 0.2;
                    }

                    // Calculate note width and X center
                    double noteXStart = startX;
                    double noteXEnd = startX + totalActorsWidth;

                    if (note.Actors.Count > 0)
                    {
                        var validActors = note.Actors.Where(lifelineXCoords.ContainsKey).ToList();
                        if (validActors.Count > 0)
                        {
                            double minX = validActors.Min(a => lifelineXCoords[a]);
                            double maxX = validActors.Max(a => lifelineXCoords[a]);
                            
                            if (validActors.Count == 1)
                            {
                                noteXStart = minX - 0.75;
                                noteXEnd = minX + 0.75;
                            }
                            else
                            {
                                noteXStart = minX - 0.25;
                                noteXEnd = maxX + 0.25;
                            }
                        }
                    }

                    double noteWidth = noteXEnd - noteXStart;
                    double noteX = (noteXStart + noteXEnd) / 2.0;

                    Visio.Shape noteBox = page.DrawRectangle(noteX - noteWidth / 2.0, noteY - 0.15, noteX + noteWidth / 2.0, noteY + 0.15);
                    noteBox.Text = note.Text;
                    ApplyTextShapeStyle(noteBox, currentFontName, currentFontSize, false);
                    
                    // Color Note: Light Yellow theme
                    TrySetFormula(noteBox, "FillForegnd", "RGB(255, 253, 230)");
                    TrySetFormula(noteBox, "LineColor", "RGB(180, 180, 100)");
                    TrySetFormula(noteBox, "LineWeight", "0.75pt");
                }
            }

            // 9. Draw Alt Blocks (Framing)
            foreach (var block in altBlocks)
            {
                double blockTop = messageYCoordinates[block.AltStartIdx] + 0.2;
                double blockBottom = messageYCoordinates[block.EndIdx] - 0.25;
                
                double frameLeft = startX - 0.35;
                double frameRight = startX + totalActorsWidth + 0.35;

                Visio.Shape altBox = page.DrawRectangle(frameLeft, blockBottom, frameRight, blockTop);
                TrySetFormula(altBox, "FillPattern", "0"); // No fill
                TrySetFormula(altBox, "LineColor", "RGB(100, 100, 100)");
                TrySetFormula(altBox, "LineWeight", "1pt");

                // Draw condition tag box in top-left
                double tagW = 1.2;
                double tagH = 0.24;
                Visio.Shape tag = page.DrawRectangle(frameLeft, blockTop - tagH, frameLeft + tagW, blockTop);
                tag.Text = "alt [" + block.AltCondition + "]";
                ApplyTextShapeStyle(tag, currentFontName, currentFontSize - 1, true);
                TrySetFormula(tag, "FillForegnd", "RGB(245, 245, 245)");
                TrySetFormula(tag, "LineColor", "RGB(100, 100, 100)");
                
                // Draw Else dashed line dividers
                foreach (var branch in block.ElseBranches)
                {
                    double elseY = messageYCoordinates[branch.StartIdx] + 0.2;
                    Visio.Shape divider = page.DrawLine(frameLeft, elseY, frameRight, elseY);
                    TrySetFormula(divider, "LineColor", "RGB(100, 100, 100)");
                    TrySetFormula(divider, "LinePattern", "2"); // Dashed
                    TrySetFormula(divider, "LineWeight", "0.75pt");

                    // Label just below separator
                    Visio.Shape label = page.DrawRectangle(frameLeft + 0.1, elseY - 0.2, frameLeft + 1.8, elseY);
                    label.Text = "[" + branch.Condition + "]";
                    ApplyTextShapeStyle(label, currentFontName, currentFontSize - 1, false);
                    TrySetFormula(label, "FillPattern", "0");
                    TrySetFormula(label, "LinePattern", "0");
                    TrySetFormula(label, "Para.HorzAlign", "0"); // Left aligned
                }
            }

            // 10. Draw Message arrows
            foreach (var msg in messages)
            {
                double y = messageYCoordinates[msg.Index];
                if (!lifelineXCoords.ContainsKey(msg.Source) || !lifelineXCoords.ContainsKey(msg.Target))
                {
                    continue;
                }

                double x1 = lifelineXCoords[msg.Source];
                double x2 = lifelineXCoords[msg.Target];

                if (msg.Type == "自关联")
                {
                    // Draw loop-back 3-point connector
                    var loopShapes = new List<Visio.Shape>();
                    Visio.Shape seg1 = page.DrawLine(x1, y, x1 + 0.45, y);
                    Visio.Shape seg2 = page.DrawLine(x1 + 0.45, y, x1 + 0.45, y - 0.22);
                    Visio.Shape seg3 = page.DrawLine(x1 + 0.45, y - 0.22, x1, y - 0.22);

                    ApplyLineStyle(seg1, false, "0");
                    ApplyLineStyle(seg2, false, "0");
                    ApplyLineStyle(seg3, false, "4"); // Filled arrowhead at end

                    // Add text label
                    Visio.Shape label = page.DrawRectangle(x1 + 0.05, y - 0.22, x1 + 1.8, y);
                    label.Text = msg.Text;
                    ApplyTextShapeStyle(label, currentFontName, currentFontSize, false);
                    TrySetFormula(label, "FillPattern", "0");
                    TrySetFormula(label, "LinePattern", "0");
                    TrySetFormula(label, "Para.HorzAlign", "0"); // Left align

                    loopShapes.Add(seg1);
                    loopShapes.Add(seg2);
                    loopShapes.Add(seg3);
                    loopShapes.Add(label);
                    GroupShapes(page, loopShapes);
                }
                else
                {
                    // Draw horizontal line always from left to right to ensure text orientation is correct
                    double startXVal = Math.Min(x1, x2);
                    double endXVal = Math.Max(x1, x2);
                    Visio.Shape line = page.DrawLine(startXVal, y, endXVal, y);
                    line.Text = msg.Text;
                    
                    bool isReturn = msg.Type == "返回";
                    bool isRightToLeft = x1 > x2;

                    bool dashed = isReturn || (isRightToLeft && chkRightToLeftDashed.Checked);
                    
                    // Determine arrow direction
                    string beginArrow = "0";
                    string endArrow = "0";
                    if (isRightToLeft)
                    {
                        beginArrow = "4"; // Pointing left (towards the start of the line)
                    }
                    else
                    {
                        endArrow = "4"; // Pointing right (towards the end of the line)
                    }
                    
                    TrySetFormula(line, "LineColor", "RGB(0, 0, 0)");
                    TrySetFormula(line, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
                    TrySetFormula(line, "LinePattern", dashed ? "2" : "1");
                    TrySetFormula(line, "BeginArrow", beginArrow);
                    TrySetFormula(line, "EndArrow", endArrow);
                    
                    ApplyTextShapeStyle(line, currentFontName, currentFontSize, false);
                    
                    // Make text background white/opaque
                    TrySetFormula(line, "TextBkgnd", "1");
                }
            }
        }

        private void ParseAltNoteTextBox(string rulesText, List<AltBlock> altBlocks, List<NoteItem> stageNotes)
        {
            if (string.IsNullOrWhiteSpace(rulesText)) return;

            string[] lines = rulesText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                if (line.StartsWith("alt ", StringComparison.OrdinalIgnoreCase))
                {
                    // Format: alt 扣款成功, 7 else 扣款失败, 8 end, 9
                    // Or nested
                    ParseAltRuleLine(line, altBlocks);
                }
                else if (line.StartsWith("阶段:", StringComparison.OrdinalIgnoreCase))
                {
                    // Format: 阶段:用户, 运营商接口:阶段一：充值请求发起, 1
                    ParseStageNoteLine(line, stageNotes);
                }
            }
        }

        private void ParseAltRuleLine(string line, List<AltBlock> altBlocks)
        {
            // Parse: alt <cond1>, <idx1> else <cond2>, <idx2> ... end, <endIdx>
            try
            {
                // Regex to extract initial alt block: alt (.*?), (\d+)
                var altMatch = Regex.Match(line, @"^alt\s+(.*?),\s*(\d+)");
                if (!altMatch.Success) return;

                string altCond = altMatch.Groups[1].Value.Trim();
                int altStart = int.Parse(altMatch.Groups[2].Value);

                var block = new AltBlock
                {
                    AltCondition = altCond,
                    AltStartIdx = altStart
                };

                // Find all else clauses
                var elseMatches = Regex.Matches(line, @"else\s+(.*?),\s*(\d+)");
                foreach (Match m in elseMatches)
                {
                    string cond = m.Groups[1].Value.Trim();
                    int start = int.Parse(m.Groups[2].Value);
                    block.ElseBranches.Add(new ElseBranch { Condition = cond, StartIdx = start });
                }

                // Find end
                var endMatch = Regex.Match(line, @"end,\s*(\d+)$");
                if (endMatch.Success)
                {
                    block.EndIdx = int.Parse(endMatch.Groups[1].Value);
                }
                else
                {
                    // Fallback to start index if end not found
                    block.EndIdx = altStart;
                }

                altBlocks.Add(block);
            }
            catch
            {
                // Skip malformed alt rules
            }
        }

        private void ParseStageNoteLine(string line, List<NoteItem> stageNotes)
        {
            // Parse: 阶段:<actorsStr>:<noteText>, <msgIdx>
            try
            {
                int firstColon = line.IndexOf(':');
                int lastColon = line.LastIndexOf(':');
                if (firstColon >= 0 && lastColon > firstColon)
                {
                    string actorsPart = line.Substring(firstColon + 1, lastColon - firstColon - 1).Trim();
                    string rest = line.Substring(lastColon + 1).Trim();
                    
                    int lastComma = rest.LastIndexOf(',');
                    string noteText = rest;
                    int msgIdx = 1;
                    if (lastComma >= 0)
                    {
                        noteText = rest.Substring(0, lastComma).Trim();
                        int.TryParse(rest.Substring(lastComma + 1).Trim(), out msgIdx);
                    }
                    
                    var actorsList = actorsPart.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                                .Select(a => a.Trim())
                                                .ToList();
                    
                    stageNotes.Add(new NoteItem
                    {
                        Actors = actorsList,
                        Text = noteText,
                        NextMessageIndex = msgIdx
                    });
                }
            }
            catch
            {
                // Skip malformed note lines
            }
        }

        private void ApplyLineStyle(Visio.Shape line, bool dashed, string endArrow)
        {
            TrySetFormula(line, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(line, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(line, "LinePattern", dashed ? "2" : "1");
            TrySetFormula(line, "BeginArrow", "0");
            TrySetFormula(line, "EndArrow", endArrow);
        }

        private void ApplyBoxStyle(Visio.Shape shape)
        {
            TrySetFormula(shape, "FillPattern", "1");
            TrySetFormula(shape, "FillForegnd", "RGB(255, 255, 255)");
            TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(shape, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
        }

        private void ApplyTextShapeStyle(Visio.Shape shape, string fontName, double fontSize, bool bold)
        {
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSize.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(0, 0, 0)");
            TrySetFormula(shape, "Char.Style", bold ? "1" : "0");
            TrySetFormula(shape, "Para.HorzAlign", "1"); // Centered
            TrySetFormula(shape, "VerticalAlign", "1"); // Middle
        }

        private Visio.Shape DrawStickActor(Visio.Page page, string name, double x, double y)
        {
            var shapes = new List<Visio.Shape>();
            double headR = 0.12;

            // Head (circle)
            Visio.Shape head = page.DrawOval(x - headR, y + 0.24, x + headR, y + 0.48);
            ApplyBoxStyle(head);
            shapes.Add(head);

            // Body
            shapes.Add(page.DrawLine(x, y + 0.24, x, y - 0.18));
            // Arms
            shapes.Add(page.DrawLine(x - 0.25, y + 0.08, x + 0.25, y + 0.08));
            // Legs
            shapes.Add(page.DrawLine(x, y - 0.18, x - 0.24, y - 0.48));
            shapes.Add(page.DrawLine(x, y - 0.18, x + 0.24, y - 0.48));

            // Name Label Box below feet
            Visio.Shape label = page.DrawRectangle(x - 0.6, y - 0.82, x + 0.6, y - 0.52);
            label.Text = name;
            ApplyTextShapeStyle(label, currentFontName, currentFontSize, false);
            TrySetFormula(label, "FillPattern", "0");
            TrySetFormula(label, "LinePattern", "0");
            shapes.Add(label);

            // Style lines
            foreach (var shape in shapes.Skip(1).Take(4))
            {
                TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
                TrySetFormula(shape, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            }

            return GroupShapes(page, shapes);
        }

        private Visio.Shape GroupShapes(Visio.Page page, List<Visio.Shape> shapes)
        {
            Visio.Selection selection = page.Application.ActiveWindow.Selection;
            selection.DeselectAll();
            foreach (Visio.Shape shape in shapes)
            {
                selection.Select(shape, (short)Visio.VisSelectArgs.visSelect);
            }
            return selection.Group();
        }

        private void ClearPageShapes(Visio.Page page)
        {
            if (page == null) return;
            for (int i = page.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    page.Shapes[i].Delete();
                }
                catch { }
            }
        }

        private Visio.Page GetOrCreateActivePage(Visio.Application visioApp)
        {
            if (visioApp == null)
            {
                throw new InvalidOperationException("未能获取 Visio 应用实例，请确认插件已在 Visio 中加载。");
            }

            try
            {
                if (visioApp.ActiveDocument != null && visioApp.ActivePage != null)
                {
                    return visioApp.ActivePage;
                }
            }
            catch { }

            Visio.Document newDoc = visioApp.Documents.Add("");
            return newDoc.Pages[1];
        }

        private void TrySetFormula(Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                shape.CellsU[cellName].FormulaU = formula;
            }
            catch { }
        }

        private double ParsePositiveNumber(string value, string fieldName)
        {
            double parsedValue;
            string normalizedValue = (value ?? string.Empty).Trim().Replace('，', '.');
            if (!double.TryParse(normalizedValue, NumberStyles.Float, CultureInfo.CurrentCulture, out parsedValue) &&
                !double.TryParse(normalizedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out parsedValue))
            {
                throw new ArgumentException($"{fieldName}必须是数字。");
            }

            if (parsedValue <= 0)
            {
                throw new ArgumentException($"{fieldName}必须大于0。");
            }

            return parsedValue;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AppendLog(string message)
        {
            isInternalTextChange = true;
            try
            {
                txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now:HH:mm:ss} {message}");
            }
            finally
            {
                isInternalTextChange = false;
            }
        }

        private class MessageItem
        {
            public int Index { get; set; }
            public string Source { get; set; }
            public string Target { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
        }

        private class NoteItem
        {
            public List<string> Actors { get; set; }
            public string Text { get; set; }
            public int NextMessageIndex { get; set; }
        }

        private class AltBlock
        {
            public string AltCondition { get; set; }
            public int AltStartIdx { get; set; }
            public List<ElseBranch> ElseBranches { get; set; } = new List<ElseBranch>();
            public int EndIdx { get; set; }
        }

        private class ElseBranch
        {
            public string Condition { get; set; }
            public int StartIdx { get; set; }
        }
    }
}
