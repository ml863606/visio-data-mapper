using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormBasicFlowchart : Form
    {
        private Button btnAiPrompt;
        private Button btnImportTable;
        private Button btnParseMermaid;
        private Label lblTitle;
        private TextBox txtTitle;
        private TabControl tabControl;
        private TabPage tabTable;
        private TabPage tabMermaid;
        private DataGridView dgvFlow;
        private TextBox txtMermaidInput;
        private TextBox txtStatus;
        private Panel pnlOptions;
        private ComboBox cmbDirection;
        private TextBox txtNodeSpacing;
        private ComboBox cmbFontName;
        private ComboBox cmbFontSize;
        private CheckBox chkCompactLayout;
        private Button btnGenerate;
        private Button btnClose;
        private bool isInternalTextChange;

        private const double DefaultShapeWidthMm = 32.0;
        private const double DefaultShapeHeightMm = 16.0;
        // 每个图形之间的纵行间距默认值，单位 mm。
        private const double DefaultNodeSpacingMm = 10.0;
        // 紧凑排版时每个图形之间的纵行间距上限，单位 mm。
        private const double CompactNodeSpacingMm = 12.0;
        // 基本流程图弹窗打开时“紧凑排版”是否默认勾选。
        private const bool DefaultCompactLayoutChecked = true;
        private const string DefaultFontName = "Microsoft YaHei";
        private const string DefaultDrawingFontName = "宋体";
        private const double DefaultDrawingFontSizePt = 10.5;

        private static readonly string[] ShapeTypes = new[]
        {
            "开始",
            "流程",
            "判定",
            "文档",
            "结束",
            "数据",
            "子流程"
        };

        public FormBasicFlowchart()
        {
            InitializeComponent();
            LoadSampleData();
        }

        private void InitializeComponent()
        {
            Text = "智能画图-基本流程图";
            Size = new Size(980, 820);
            MinimumSize = new Size(900, 720);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            BackColor = Color.FromArgb(240, 244, 248);
            Font = new Font(DefaultFontName, 9F, FontStyle.Regular);

            Label lblTip = new Label
            {
                Text = "请将流程信息交给 AI 生成 Mermaid flowchart 代码，再导入表格后生成流程图。",
                Location = new Point(15, 14),
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104),
                Font = new Font(DefaultFontName, 9.5F, FontStyle.Regular)
            };

            btnAiPrompt = CreatePrimaryButton("AI生成提示词", new Point(15, 42), new Size(120, 32), Color.FromArgb(91, 76, 196));
            btnAiPrompt.Click += btnAiPrompt_Click;

            btnImportTable = CreatePrimaryButton("导入表格", new Point(145, 42), new Size(105, 32), Color.FromArgb(0, 122, 255));
            btnImportTable.Click += btnImportTable_Click;

            btnParseMermaid = CreatePrimaryButton("解析Mermaid", new Point(260, 42), new Size(120, 32), Color.FromArgb(21, 128, 61));
            btnParseMermaid.Click += btnParseMermaid_Click;

            lblTitle = new Label
            {
                Text = "流程图标题:",
                Location = new Point(410, 48),
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104)
            };

            txtTitle = new TextBox
            {
                Location = new Point(495, 44),
                Size = new Size(440, 25),
                Font = new Font(DefaultFontName, 10F, FontStyle.Regular)
            };

            tabControl = new TabControl
            {
                Location = new Point(15, 88),
                Size = new Size(930, 480),
                Font = new Font(DefaultFontName, 9F, FontStyle.Regular)
            };

            tabTable = new TabPage("流程内容") { BackColor = Color.White };
            tabMermaid = new TabPage("Mermaid代码") { BackColor = Color.White };

            dgvFlow = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                GridColor = Color.FromArgb(220, 226, 235),
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 28 }
            };
            dgvFlow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Index", HeaderText = "序号", FillWeight = 55 });
            dgvFlow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Text", HeaderText = "工作内容", FillWeight = 180 });
            var symbolColumn = new DataGridViewComboBoxColumn { Name = "Symbol", HeaderText = "符号", FillWeight = 90 };
            symbolColumn.Items.AddRange(ShapeTypes);
            dgvFlow.Columns.Add(symbolColumn);
            dgvFlow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Next", HeaderText = "下一步", FillWeight = 100 });
            dgvFlow.Columns.Add(new DataGridViewTextBoxColumn { Name = "LineText", HeaderText = "连接线文字", FillWeight = 140 });
            ApplyGridStyle(dgvFlow);
            tabTable.Controls.Add(dgvFlow);

            txtMermaidInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F, FontStyle.Regular),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            txtMermaidInput.TextChanged += txtMermaidInput_TextChanged;
            tabMermaid.Controls.Add(txtMermaidInput);

            tabControl.TabPages.Add(tabTable);
            tabControl.TabPages.Add(tabMermaid);

            txtStatus = new TextBox
            {
                Location = new Point(15, 580),
                Size = new Size(930, 70),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(248, 250, 252),
                ForeColor = Color.FromArgb(71, 85, 105),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9F, FontStyle.Regular),
                Text = $"{DateTime.Now:HH:mm:ss} 已启动基本流程图。可粘贴 Mermaid flowchart 后点击“解析Mermaid”。"
            };

            pnlOptions = new Panel
            {
                Location = new Point(15, 665),
                Size = new Size(930, 112),
                BackColor = Color.White
            };
            pnlOptions.Paint += (s, pe) =>
            {
                using (var pen = new Pen(Color.FromArgb(218, 224, 233), 1))
                {
                    pe.Graphics.DrawRectangle(pen, 0, 0, pnlOptions.Width - 1, pnlOptions.Height - 1);
                }
            };

            BuildOptionsPanel();

            Controls.Add(lblTip);
            Controls.Add(btnAiPrompt);
            Controls.Add(btnImportTable);
            Controls.Add(btnParseMermaid);
            Controls.Add(lblTitle);
            Controls.Add(txtTitle);
            Controls.Add(tabControl);
            Controls.Add(txtStatus);
            Controls.Add(pnlOptions);

            Resize += FormBasicFlowchart_Resize;
            LayoutControls();
        }

        private Button CreatePrimaryButton(string text, Point location, Size size, Color color)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = size,
                FlatStyle = FlatStyle.Flat,
                BackColor = color,
                ForeColor = Color.White,
                Font = new Font(DefaultFontName, 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void BuildOptionsPanel()
        {
            var firstRow = CreateOptionsRow(15, 12);
            var secondRow = CreateOptionsRow(15, 56);

            cmbDirection = new ComboBox { Size = new Size(105, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDirection.Items.AddRange(new[] { "从上到下", "从左到右" });
            cmbDirection.SelectedIndex = 0;

            txtNodeSpacing = new TextBox { Text = DefaultNodeSpacingMm.ToString("0.#", CultureInfo.InvariantCulture), Size = new Size(55, 25), TextAlign = HorizontalAlignment.Center };

            chkCompactLayout = new CheckBox { Text = "紧凑排版", Size = new Size(95, 24), AutoSize = false, Checked = DefaultCompactLayoutChecked, Margin = new Padding(18, 2, 0, 0) };

            cmbFontName = new ComboBox { Size = new Size(120, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontName.Items.AddRange(new[] { "宋体", "微软雅黑", "黑体", "楷体", "仿宋" });
            cmbFontName.SelectedItem = DefaultDrawingFontName;

            cmbFontSize = new ComboBox { Size = new Size(85, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontSize.Items.AddRange(new[] { "三号", "小三", "四号", "小四", "五号", "小五" });
            cmbFontSize.SelectedItem = "五号";

            btnGenerate = CreatePrimaryButton("生成流程图", new Point(680, 56), new Size(120, 36), Color.FromArgb(40, 167, 69));
            btnGenerate.Click += btnGenerate_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(815, 56),
                Size = new Size(90, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(233, 236, 239),
                ForeColor = Color.FromArgb(73, 80, 87),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => Close();

            firstRow.Controls.Add(CreateLabeledOption("方向:", cmbDirection));
            firstRow.Controls.Add(CreateLabeledOption("图形纵行间距(mm):", txtNodeSpacing));
            firstRow.Controls.Add(chkCompactLayout);
            secondRow.Controls.Add(CreateLabeledOption("字体:", cmbFontName));
            secondRow.Controls.Add(CreateLabeledOption("字号:", cmbFontSize));
            pnlOptions.Controls.Add(firstRow);
            pnlOptions.Controls.Add(secondRow);
            pnlOptions.Controls.Add(btnGenerate);
            pnlOptions.Controls.Add(btnClose);
        }

        private FlowLayoutPanel CreateOptionsRow(int left, int top)
        {
            return new FlowLayoutPanel
            {
                Location = new Point(left, top),
                Size = new Size(620, 32),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoScroll = false,
                BackColor = Color.Transparent
            };
        }

        private Control CreateLabeledOption(string labelText, Control editor)
        {
            var panel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 0, 22, 0),
                BackColor = Color.Transparent
            };

            var label = new Label
            {
                Text = labelText,
                AutoSize = true,
                Margin = new Padding(0, 6, 6, 0)
            };
            editor.Margin = new Padding(0, 2, 0, 0);
            panel.Controls.Add(label);
            panel.Controls.Add(editor);
            return panel;
        }

        private void ApplyGridStyle(DataGridView grid)
        {
            grid.EnableHeadersVisualStyles = false;
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 248);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(30, 41, 59);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(DefaultFontName, 9F, FontStyle.Bold);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(15, 23, 42);
        }

        private void LayoutControls()
        {
            int margin = 15;
            int width = Math.Max(850, ClientSize.Width - margin * 2);
            tabControl.Width = width;
            pnlOptions.Width = width;
            pnlOptions.Top = ClientSize.Height - pnlOptions.Height - margin;
            txtStatus.Width = width;
            txtStatus.Top = pnlOptions.Top - txtStatus.Height - 15;
            tabControl.Height = Math.Max(260, txtStatus.Top - tabControl.Top - 12);
            btnClose.Left = pnlOptions.Width - btnClose.Width - 20;
            btnGenerate.Left = btnClose.Left - btnGenerate.Width - 15;
            btnClose.Top = 56;
            btnGenerate.Top = 56;
            txtTitle.Width = Math.Max(260, width - txtTitle.Left + margin);
        }

        private void FormBasicFlowchart_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void LoadSampleData()
        {
            txtTitle.Text = "请假审批流程";
            txtMermaidInput.Text =
@"flowchart TD
    A([开始])
    B[提交请假申请]
    C{直属领导审批}
    D[人事备案]
    E[驳回申请]
    F([结束])
    A --> B
    B --> C
    C -->|通过| D
    C -->|不通过| E
    D --> F
    E --> F";
            ParseMermaidToGrid(txtMermaidInput.Text, false);
        }

        private void AppendLog(string message)
        {
            txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now:HH:mm:ss} {message}");
        }

        private void btnAiPrompt_Click(object sender, EventArgs e)
        {
            string prompt =
@"请根据下面的流程信息，生成 Mermaid flowchart 代码，只输出代码，不要解释。
要求：
1. 使用 flowchart TD 或 flowchart LR。
2. 节点 ID 使用 A、B、C...。
3. 开始/结束使用 A([文本])。
4. 普通流程使用 A[文本]。
5. 判断使用 A{文本}。
6. 判断分支连接线使用 A -->|是| B 这种格式。

流程信息：
";
            Clipboard.SetText(prompt);
            AppendLog("已复制 AI 生成 Mermaid 的提示词到剪贴板。");
        }

        private void btnImportTable_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    MessageBox.Show("剪贴板中没有文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string text = Clipboard.GetText().Trim();
                if (IsMermaidFlowchart(text))
                {
                    isInternalTextChange = true;
                    try
                    {
                        txtMermaidInput.Text = text;
                    }
                    finally
                    {
                        isInternalTextChange = false;
                    }
                    ParseMermaidToGrid(text, true);
                    tabControl.SelectedTab = tabTable;
                    return;
                }

                ParseTabularToGrid(text);
                tabControl.SelectedTab = tabTable;
            }
            catch (Exception ex)
            {
                AppendLog("导入失败: " + ex.Message);
                MessageBox.Show("导入失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnParseMermaid_Click(object sender, EventArgs e)
        {
            try
            {
                ParseMermaidToGrid(txtMermaidInput.Text, true);
                tabControl.SelectedTab = tabTable;
            }
            catch (Exception ex)
            {
                AppendLog("解析失败: " + ex.Message);
                MessageBox.Show("解析失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtMermaidInput_TextChanged(object sender, EventArgs e)
        {
            if (isInternalTextChange) return;

            string text = txtMermaidInput.Text;
            if (!IsMermaidFlowchart(text)) return;

            try
            {
                ParseMermaidToGrid(text, false);
            }
            catch
            {
                // 输入过程中可能暂时不完整，忽略。
            }
        }

        private bool IsMermaidFlowchart(string text)
        {
            return !string.IsNullOrWhiteSpace(text) &&
                Regex.IsMatch(text, @"\b(flowchart|graph)\s+(TD|TB|BT|LR|RL)\b", RegexOptions.IgnoreCase);
        }

        private void ParseMermaidToGrid(string mermaid, bool showLog)
        {
            FlowModel model = ParseMermaid(mermaid);
            dgvFlow.Rows.Clear();
            cmbDirection.SelectedIndex = model.Direction == "LR" ? 1 : 0;

            foreach (FlowNode node in model.Nodes.Values.OrderBy(n => n.Order))
            {
                List<FlowEdge> edges = model.Edges.Where(edge => edge.FromId == node.Id).ToList();
                string next = string.Join(",", edges.Select(edge => edge.ToId));
                string labels = string.Join(",", edges.Select(edge => edge.Label ?? string.Empty));
                dgvFlow.Rows.Add(node.Id, node.Text, node.Symbol, next, labels);
            }

            if (showLog)
            {
                AppendLog($"已解析 Mermaid：{model.Nodes.Count} 个节点，{model.Edges.Count} 条连接。");
            }
        }

        private FlowModel ParseMermaid(string mermaid)
        {
            if (!IsMermaidFlowchart(mermaid))
            {
                throw new ArgumentException("请输入 Mermaid flowchart 代码，例如 flowchart TD。");
            }

            var model = new FlowModel();
            var lines = mermaid.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int order = 0;

            foreach (string rawLine in lines)
            {
                string line = CleanMermaidLine(rawLine);
                if (string.IsNullOrWhiteSpace(line)) continue;

                Match header = Regex.Match(line, @"^(flowchart|graph)\s+(TD|TB|BT|LR|RL)\b", RegexOptions.IgnoreCase);
                if (header.Success)
                {
                    string direction = header.Groups[2].Value.ToUpperInvariant();
                    model.Direction = direction == "LR" || direction == "RL" ? "LR" : "TD";
                    continue;
                }

                ParseMermaidLine(line, model, ref order);
            }

            foreach (FlowEdge edge in model.Edges)
            {
                EnsureNode(model, edge.FromId, edge.FromId, "流程", ref order);
                EnsureNode(model, edge.ToId, edge.ToId, "流程", ref order);
            }

            if (model.Nodes.Count == 0)
            {
                throw new ArgumentException("没有解析到流程节点。");
            }

            return model;
        }

        private string CleanMermaidLine(string rawLine)
        {
            string line = rawLine.Trim();
            if (line.StartsWith("%%")) return string.Empty;
            int commentIndex = line.IndexOf("%%", StringComparison.Ordinal);
            if (commentIndex >= 0)
            {
                line = line.Substring(0, commentIndex).Trim();
            }
            return line.TrimEnd(';').Trim();
        }

        private void ParseMermaidLine(string line, FlowModel model, ref int order)
        {
            string[] edgeTokens = new[] { "-->", "---", "-.->", "==>" };
            string token = edgeTokens.FirstOrDefault(line.Contains);
            if (!string.IsNullOrEmpty(token))
            {
                int tokenIndex = line.IndexOf(token, StringComparison.Ordinal);
                string left = line.Substring(0, tokenIndex).Trim();
                string right = line.Substring(tokenIndex + token.Length).Trim();
                string label = ExtractEdgeLabel(ref left, ref right);

                ParsedNode fromNode = ParseNodeExpression(left);
                ParsedNode toNode = ParseNodeExpression(right);
                EnsureNode(model, fromNode.Id, fromNode.Text, fromNode.Symbol, ref order);
                EnsureNode(model, toNode.Id, toNode.Text, toNode.Symbol, ref order);
                model.Edges.Add(new FlowEdge { FromId = fromNode.Id, ToId = toNode.Id, Label = label });
                return;
            }

            ParsedNode node = ParseNodeExpression(line);
            EnsureNode(model, node.Id, node.Text, node.Symbol, ref order);
        }

        private string ExtractEdgeLabel(ref string left, ref string right)
        {
            string label = string.Empty;
            Match pipe = Regex.Match(right, @"^\|([^|]*)\|\s*(.+)$");
            if (pipe.Success)
            {
                label = pipe.Groups[1].Value.Trim();
                right = pipe.Groups[2].Value.Trim();
                return label;
            }

            Match middle = Regex.Match(left, @"^(.+?)\s+--\s*(.+)$");
            if (middle.Success)
            {
                left = middle.Groups[1].Value.Trim();
                label = middle.Groups[2].Value.Trim();
            }

            return label;
        }

        private ParsedNode ParseNodeExpression(string expression)
        {
            string value = expression.Trim();
            value = Regex.Replace(value, @"\s+:::.*$", string.Empty).Trim();

            Match match = Regex.Match(value, @"^([A-Za-z0-9_\-\u4e00-\u9fa5]+)\s*(.+)?$");
            if (!match.Success)
            {
                throw new ArgumentException("无法解析 Mermaid 节点: " + expression);
            }

            string id = match.Groups[1].Value.Trim();
            string shapeExpr = match.Groups[2].Value.Trim();
            if (string.IsNullOrWhiteSpace(shapeExpr))
            {
                return new ParsedNode { Id = id, Text = id, Symbol = "流程" };
            }

            string text = ExtractNodeText(shapeExpr);
            string symbol = DetectSymbol(shapeExpr, text);
            return new ParsedNode { Id = id, Text = string.IsNullOrWhiteSpace(text) ? id : text, Symbol = symbol };
        }

        private string ExtractNodeText(string shapeExpr)
        {
            string text = shapeExpr.Trim();
            bool changed = true;
            while (changed && text.Length >= 2)
            {
                changed = false;
                string before = text;
                if ((text.StartsWith("[") && text.EndsWith("]")) ||
                    (text.StartsWith("(") && text.EndsWith(")")) ||
                    (text.StartsWith("{") && text.EndsWith("}")))
                {
                    text = text.Substring(1, text.Length - 2).Trim();
                }
                if ((text.StartsWith("/") && text.EndsWith("/")) ||
                    (text.StartsWith("\\") && text.EndsWith("\\")))
                {
                    text = text.Substring(1, text.Length - 2).Trim();
                }
                changed = before != text;
            }

            return text.Trim('"');
        }

        private string DetectSymbol(string shapeExpr, string text)
        {
            string expr = shapeExpr.Trim();
            if (expr.StartsWith("{")) return "判定";
            if (expr.StartsWith("[[") || expr.StartsWith("[|")) return "子流程";
            if (expr.StartsWith("[(")) return "数据";
            if (expr.StartsWith("[/") || expr.StartsWith("[\\")) return "文档";
            if (expr.StartsWith("((")) return "开始";
            if (expr.StartsWith("([")) return LooksLikeEnd(text) ? "结束" : "开始";
            return "流程";
        }

        private bool LooksLikeEnd(string text)
        {
            string normalized = (text ?? string.Empty).Trim();
            return normalized.Contains("结束") ||
                normalized.Equals("End", StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("Finish", StringComparison.OrdinalIgnoreCase);
        }

        private void EnsureNode(FlowModel model, string id, string text, string symbol, ref int order)
        {
            if (string.IsNullOrWhiteSpace(id)) return;

            if (!model.Nodes.TryGetValue(id, out FlowNode node))
            {
                model.Nodes[id] = new FlowNode
                {
                    Id = id,
                    Text = string.IsNullOrWhiteSpace(text) ? id : text,
                    Symbol = NormalizeSymbol(symbol),
                    Order = ++order
                };
                return;
            }

            if (!string.IsNullOrWhiteSpace(text) && node.Text == node.Id)
            {
                node.Text = text;
            }
            if (!string.IsNullOrWhiteSpace(symbol) && node.Symbol == "流程")
            {
                node.Symbol = NormalizeSymbol(symbol);
            }
        }

        private string NormalizeSymbol(string symbol)
        {
            if (ShapeTypes.Contains(symbol)) return symbol;
            if (symbol == "开始/结束") return "开始";
            return "流程";
        }

        private void ParseTabularToGrid(string text)
        {
            dgvFlow.Rows.Clear();
            string[] rows = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string row in rows)
            {
                string[] cells = row.Split('\t');
                if (cells.Length < 2) continue;
                if (cells[0].Trim() == "序号") continue;

                string id = cells[0].Trim();
                string content = cells.Length > 1 ? cells[1].Trim() : string.Empty;
                string symbol = cells.Length > 2 ? NormalizeSymbol(cells[2].Trim()) : "流程";
                string next = cells.Length > 3 ? cells[3].Trim() : string.Empty;
                string lineText = cells.Length > 4 ? cells[4].Trim() : string.Empty;
                dgvFlow.Rows.Add(id, content, symbol, next, lineText);
            }
            AppendLog($"已从表格文本导入 {dgvFlow.Rows.Cast<DataGridViewRow>().Count(r => !r.IsNewRow)} 行。");
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                RenderFlowchart();
                AppendLog("基本流程图生成成功。");
                Close();
            }
            catch (Exception ex)
            {
                AppendLog("生成失败: " + ex.Message);
                MessageBox.Show("生成失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderFlowchart()
        {
            FlowModel model = ReadGridModel();
            if (model.Nodes.Count == 0)
            {
                throw new InvalidOperationException("表格中没有可绘制的流程节点。");
            }

            double shapeWidth = DefaultShapeWidthMm / 25.4;
            double shapeHeight = DefaultShapeHeightMm / 25.4;
            double nodeSpacing = ParsePositiveNumber(txtNodeSpacing.Text, "图形纵行间距") / 25.4;
            bool leftToRight = cmbDirection.SelectedIndex == 1;
            string drawingFontName = GetSelectedFontName();
            double drawingFontSizePt = GetSelectedFontSizePt();
            bool compactLayout = chkCompactLayout.Checked;
            if (compactLayout)
            {
                nodeSpacing = Math.Min(nodeSpacing, CompactNodeSpacingMm / 25.4);
            }

            Dictionary<string, SizeF> nodeSizes = CalculateNodeSizes(model, shapeWidth, shapeHeight, compactLayout, drawingFontSizePt);
            double layoutShapeWidth = nodeSizes.Values.Max(size => size.Width);
            double layoutShapeHeight = nodeSizes.Values.Max(size => size.Height);

            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Page page = GetOrCreateActivePage(app);
            ClearPageShapes(page);

            Dictionary<string, PointF> coordinates = CalculateLayeredLayout(model, leftToRight, layoutShapeWidth, layoutShapeHeight, nodeSpacing, compactLayout);
            FitPage(page, coordinates, layoutShapeWidth, layoutShapeHeight);

            double pageHeight = page.PageSheet.CellsU["PageHeight"].Result["in"];
            if (!string.IsNullOrWhiteSpace(txtTitle.Text))
            {
                DrawTitle(page, txtTitle.Text.Trim(), pageHeight, drawingFontName);
            }

            foreach (FlowNode node in model.Nodes.Values.OrderBy(n => n.Order))
            {
                PointF p = coordinates[node.Id];
                SizeF nodeSize = nodeSizes[node.Id];
                node.VisioShape = DrawFlowShape(page, node, p.X, p.Y, nodeSize.Width, nodeSize.Height, drawingFontName, drawingFontSizePt);
            }

            DrawConnectors(page, model, drawingFontName, drawingFontSizePt);
            ApplyPageRouting(page, leftToRight);
        }

        private Dictionary<string, SizeF> CalculateNodeSizes(FlowModel model, double defaultWidth, double defaultHeight, bool compactLayout, double fontSizePt)
        {
            var result = new Dictionary<string, SizeF>();
            foreach (FlowNode node in model.Nodes.Values)
            {
                if (!compactLayout)
                {
                    result[node.Id] = new SizeF((float)defaultWidth, (float)defaultHeight);
                    continue;
                }

                int charCount = Math.Max(2, GetDisplayCharCount(node.Text));
                double width = Math.Max(0.55, charCount * fontSizePt / 72.0 * 0.58 + 0.38);
                double height = Math.Max(0.28, fontSizePt / 72.0 * 1.9);
                if (node.Symbol == "判定")
                {
                    width += 0.35;
                    height += 0.12;
                }
                else if (node.Symbol == "开始" || node.Symbol == "结束")
                {
                    width += 0.18;
                }

                result[node.Id] = new SizeF((float)width, (float)height);
            }

            return result;
        }

        private int GetDisplayCharCount(string text)
        {
            int count = 0;
            foreach (char ch in text ?? string.Empty)
            {
                count += ch <= 127 ? 1 : 2;
            }
            return Math.Max(1, count);
        }

        private string GetSelectedFontName()
        {
            return cmbFontName == null || cmbFontName.SelectedItem == null
                ? DefaultDrawingFontName
                : cmbFontName.SelectedItem.ToString();
        }

        private double GetSelectedFontSizePt()
        {
            string selectedSize = cmbFontSize == null || cmbFontSize.SelectedItem == null ? "五号" : cmbFontSize.SelectedItem.ToString();
            switch (selectedSize)
            {
                case "三号":
                    return 16.0;
                case "小三":
                    return 15.0;
                case "四号":
                    return 14.0;
                case "小四":
                    return 12.0;
                case "小五":
                    return 9.0;
                case "五号":
                default:
                    return DefaultDrawingFontSizePt;
            }
        }

        private FlowModel ReadGridModel()
        {
            var model = new FlowModel { Direction = cmbDirection.SelectedIndex == 1 ? "LR" : "TD" };
            int order = 0;
            foreach (DataGridViewRow row in dgvFlow.Rows)
            {
                if (row.IsNewRow) continue;
                string id = Convert.ToString(row.Cells["Index"].Value)?.Trim();
                if (string.IsNullOrWhiteSpace(id)) continue;

                string text = Convert.ToString(row.Cells["Text"].Value)?.Trim();
                string symbol = NormalizeSymbol(Convert.ToString(row.Cells["Symbol"].Value)?.Trim());
                model.Nodes[id] = new FlowNode
                {
                    Id = id,
                    Text = string.IsNullOrWhiteSpace(text) ? id : text,
                    Symbol = symbol,
                    Order = ++order
                };
            }

            foreach (DataGridViewRow row in dgvFlow.Rows)
            {
                if (row.IsNewRow) continue;
                string id = Convert.ToString(row.Cells["Index"].Value)?.Trim();
                if (string.IsNullOrWhiteSpace(id) || !model.Nodes.ContainsKey(id)) continue;

                string next = Convert.ToString(row.Cells["Next"].Value) ?? string.Empty;
                string lineText = Convert.ToString(row.Cells["LineText"].Value) ?? string.Empty;
                string[] nextIds = next.Split(new[] { ',', '，', ' ', ';', '；' }, StringSplitOptions.RemoveEmptyEntries);
                string[] labels = lineText.Split(new[] { ',', '，' }, StringSplitOptions.None);
                for (int i = 0; i < nextIds.Length; i++)
                {
                    string targetId = nextIds[i].Trim();
                    if (!model.Nodes.ContainsKey(targetId)) continue;
                    model.Edges.Add(new FlowEdge
                    {
                        FromId = id,
                        ToId = targetId,
                        Label = i < labels.Length ? labels[i].Trim() : string.Empty
                    });
                }
            }

            return model;
        }

        private Dictionary<string, PointF> CalculateLayeredLayout(FlowModel model, bool leftToRight, double shapeWidth, double shapeHeight, double nodeSpacing, bool compactLayout)
        {
            var indegree = model.Nodes.Keys.ToDictionary(id => id, id => 0);
            foreach (FlowEdge edge in model.Edges)
            {
                if (indegree.ContainsKey(edge.ToId)) indegree[edge.ToId]++;
            }

            var layerByNode = model.Nodes.Keys.ToDictionary(id => id, id => 0);
            var queue = new Queue<string>(model.Nodes.Values.Where(n => indegree[n.Id] == 0).OrderBy(n => n.Order).Select(n => n.Id));
            if (queue.Count == 0)
            {
                queue.Enqueue(model.Nodes.Values.OrderBy(n => n.Order).First().Id);
            }

            var visited = new HashSet<string>();
            while (queue.Count > 0)
            {
                string id = queue.Dequeue();
                if (!visited.Add(id)) continue;

                foreach (FlowEdge edge in model.Edges.Where(e => e.FromId == id))
                {
                    layerByNode[edge.ToId] = Math.Max(layerByNode[edge.ToId], layerByNode[id] + 1);
                    indegree[edge.ToId] = Math.Max(0, indegree[edge.ToId] - 1);
                    if (indegree[edge.ToId] == 0)
                    {
                        queue.Enqueue(edge.ToId);
                    }
                }
            }

            int fallbackLayer = layerByNode.Values.DefaultIfEmpty(0).Max() + 1;
            foreach (FlowNode node in model.Nodes.Values.OrderBy(n => n.Order))
            {
                if (!visited.Contains(node.Id))
                {
                    layerByNode[node.Id] = fallbackLayer++;
                }
            }

            var groups = model.Nodes.Values
                .GroupBy(n => layerByNode[n.Id])
                .OrderBy(g => g.Key)
                .ToList();
            var result = new Dictionary<string, PointF>();

            double startX = compactLayout ? 0.9 : 1.4;
            double minStartY = compactLayout ? 2.2 : 4.5;
            double startY = Math.Max(minStartY, groups.Max(g => g.Count()) * (shapeHeight + nodeSpacing) + (compactLayout ? 0.8 : 1.5));
            for (int layerIndex = 0; layerIndex < groups.Count; layerIndex++)
            {
                List<FlowNode> nodes = groups[layerIndex].OrderBy(n => n.Order).ToList();
                double groupSize = nodes.Count;
                for (int i = 0; i < nodes.Count; i++)
                {
                    double x;
                    double y;
                    if (leftToRight)
                    {
                        x = startX + layerIndex * (shapeWidth + nodeSpacing);
                        y = startY - (i - (groupSize - 1) / 2.0) * (shapeHeight + nodeSpacing);
                    }
                    else
                    {
                        x = startX + (i - (groupSize - 1) / 2.0) * (shapeWidth + nodeSpacing);
                        y = startY - layerIndex * (shapeHeight + nodeSpacing);
                    }
                    result[nodes[i].Id] = new PointF((float)x, (float)y);
                }
            }

            NormalizeCoordinates(result);
            return result;
        }

        private void NormalizeCoordinates(Dictionary<string, PointF> coordinates)
        {
            double minX = coordinates.Values.Min(p => p.X);
            double minY = coordinates.Values.Min(p => p.Y);
            double shiftX = minX < 1.2 ? 1.2 - minX : 0;
            double shiftY = minY < 1.2 ? 1.2 - minY : 0;
            if (shiftX <= 0 && shiftY <= 0) return;

            foreach (string id in coordinates.Keys.ToList())
            {
                PointF p = coordinates[id];
                coordinates[id] = new PointF((float)(p.X + shiftX), (float)(p.Y + shiftY));
            }
        }

        private void FitPage(Visio.Page page, Dictionary<string, PointF> coordinates, double shapeWidth, double shapeHeight)
        {
            double maxX = coordinates.Values.Max(p => p.X) + shapeWidth + 1.0;
            double maxY = coordinates.Values.Max(p => p.Y) + shapeHeight + 1.0;
            if (maxX > page.PageSheet.CellsU["PageWidth"].Result["in"])
            {
                page.PageSheet.CellsU["PageWidth"].FormulaU = $"{maxX.ToString(CultureInfo.InvariantCulture)} in";
            }
            if (maxY > page.PageSheet.CellsU["PageHeight"].Result["in"])
            {
                page.PageSheet.CellsU["PageHeight"].FormulaU = $"{maxY.ToString(CultureInfo.InvariantCulture)} in";
            }
        }

        private void DrawTitle(Visio.Page page, string title, double pageHeight, string fontName)
        {
            double pageWidth = page.PageSheet.CellsU["PageWidth"].Result["in"];
            Visio.Shape titleShape = page.DrawRectangle(pageWidth / 2.0 - 1.8, pageHeight - 0.55, pageWidth / 2.0 + 1.8, pageHeight - 0.15);
            titleShape.Text = title;
            TrySetFormula(titleShape, "FillPattern", "0");
            TrySetFormula(titleShape, "LinePattern", "0");
            TrySetFormula(titleShape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(titleShape, "Char.Size", "14pt");
            TrySetFormula(titleShape, "Para.HorzAlign", "1");
            TrySetFormula(titleShape, "VerticalAlign", "1");
        }

        private Visio.Shape DrawFlowShape(Visio.Page page, FlowNode node, double cx, double cy, double w, double h, string fontName, double fontSizePt)
        {
            Visio.Shape shape;
            switch (node.Symbol)
            {
                case "开始":
                case "结束":
                    shape = DrawTerminatorShape(page, cx, cy, w, h);
                    break;
                case "判定":
                    shape = page.DrawPolyline(new[]
                    {
                        cx, cy + h / 2.0,
                        cx + w / 2.0, cy,
                        cx, cy - h / 2.0,
                        cx - w / 2.0, cy,
                        cx, cy + h / 2.0
                    }, 0);
                    break;
                case "数据":
                    shape = page.DrawRectangle(cx - w / 2.0, cy - h / 2.0, cx + w / 2.0, cy + h / 2.0);
                    TrySetFormula(shape, "Rounding", "0.08 in");
                    break;
                case "子流程":
                    shape = page.DrawRectangle(cx - w / 2.0, cy - h / 2.0, cx + w / 2.0, cy + h / 2.0);
                    DrawSubprocessLines(page, cx, cy, w, h);
                    break;
                case "文档":
                case "流程":
                default:
                    shape = page.DrawRectangle(cx - w / 2.0, cy - h / 2.0, cx + w / 2.0, cy + h / 2.0);
                    break;
            }

            shape.Text = node.Text;
            ApplyShapeStyle(shape, fontName, fontSizePt);
            return shape;
        }

        private Visio.Shape DrawTerminatorShape(Visio.Page page, double cx, double cy, double w, double h)
        {
            Visio.Shape shape = page.DrawRectangle(cx - w / 2.0, cy - h / 2.0, cx + w / 2.0, cy + h / 2.0);
            TrySetFormula(shape, "Rounding", $"{(h / 2.0).ToString(CultureInfo.InvariantCulture)} in");
            return shape;
        }

        private void DrawSubprocessLines(Visio.Page page, double cx, double cy, double w, double h)
        {
            double leftX = cx - w / 2.0 + 0.12;
            double rightX = cx + w / 2.0 - 0.12;
            Visio.Shape leftLine = page.DrawLine(leftX, cy - h / 2.0, leftX, cy + h / 2.0);
            Visio.Shape rightLine = page.DrawLine(rightX, cy - h / 2.0, rightX, cy + h / 2.0);
            ApplyLineStyle(leftLine, false);
            ApplyLineStyle(rightLine, false);
        }

        private void ApplyShapeStyle(Visio.Shape shape, string fontName, double fontSizePt)
        {
            TrySetFormula(shape, "FillPattern", "0");
            TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(shape, "LineWeight", "0.75pt");
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSizePt.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(0, 0, 0)");
            TrySetFormula(shape, "Para.HorzAlign", "1");
            TrySetFormula(shape, "VerticalAlign", "1");
            TrySetFormula(shape, "TxtPinX", "Width*0.5");
            TrySetFormula(shape, "TxtPinY", "Height*0.5");
            TrySetFormula(shape, "TxtWidth", "Width*0.9");
            TrySetFormula(shape, "TxtHeight", "Height*0.9");
        }

        private void DrawConnectors(Visio.Page page, FlowModel model, string fontName, double fontSizePt)
        {
            object connectorTool = page.Application.ConnectorToolDataObject;
            foreach (FlowEdge edge in model.Edges)
            {
                if (!model.Nodes.TryGetValue(edge.FromId, out FlowNode source)) continue;
                if (!model.Nodes.TryGetValue(edge.ToId, out FlowNode target)) continue;
                if (source.VisioShape == null || target.VisioShape == null) continue;

                Visio.Shape connector = page.Drop(connectorTool, 0, 0);
                connector.Text = edge.Label ?? string.Empty;
                ApplyLineStyle(connector, true);
                TrySetFormula(connector, "Char.Font", $"\"{fontName}\"");
                TrySetFormula(connector, "Char.Size", $"{fontSizePt.ToString(CultureInfo.InvariantCulture)}pt");
                connector.CellsU["BeginX"].GlueTo(source.VisioShape.CellsU["PinX"]);
                connector.CellsU["EndX"].GlueTo(target.VisioShape.CellsU["PinX"]);
            }
        }

        private void ApplyLineStyle(Visio.Shape line, bool arrow)
        {
            TrySetFormula(line, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(line, "LineWeight", "0.75pt");
            TrySetFormula(line, "BeginArrow", "0");
            TrySetFormula(line, "EndArrow", arrow ? "4" : "0");
        }

        private void ApplyPageRouting(Visio.Page page, bool leftToRight)
        {
            TrySetFormula(page.PageSheet, "RouteStyle", leftToRight ? "6" : "5");
            TrySetFormula(page.PageSheet, "PlaceStyle", leftToRight ? "2" : "1");
            TrySetFormula(page.PageSheet, "LineToNodeGap", "0.16 in");
            TrySetFormula(page.PageSheet, "LineToLineGap", "0.12 in");
        }

        private Visio.Page GetOrCreateActivePage(Visio.Application visioApp)
        {
            if (visioApp.Documents.Count == 0)
            {
                visioApp.Documents.Add("");
            }
            if (visioApp.ActivePage == null)
            {
                visioApp.ActiveDocument.Pages.Add();
            }
            return visioApp.ActivePage;
        }

        private void ClearPageShapes(Visio.Page page)
        {
            for (int i = page.Shapes.Count; i >= 1; i--)
            {
                page.Shapes[i].Delete();
            }
        }

        private double ParsePositiveNumber(string value, string fieldName)
        {
            string normalized = (value ?? string.Empty).Trim().Replace('，', '.');
            double parsed;
            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed) &&
                !double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                throw new ArgumentException($"{fieldName}必须是数字。");
            }
            if (parsed <= 0)
            {
                throw new ArgumentException($"{fieldName}必须大于 0。");
            }
            return parsed;
        }

        private void TrySetFormula(Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                shape.CellsU[cellName].FormulaU = formula;
            }
            catch { }
        }

        private class FlowModel
        {
            public string Direction { get; set; } = "TD";
            public Dictionary<string, FlowNode> Nodes { get; } = new Dictionary<string, FlowNode>();
            public List<FlowEdge> Edges { get; } = new List<FlowEdge>();
        }

        private class FlowNode
        {
            public string Id { get; set; }
            public string Text { get; set; }
            public string Symbol { get; set; }
            public int Order { get; set; }
            public Visio.Shape VisioShape { get; set; }
        }

        private class FlowEdge
        {
            public string FromId { get; set; }
            public string ToId { get; set; }
            public string Label { get; set; }
        }

        private class ParsedNode
        {
            public string Id { get; set; }
            public string Text { get; set; }
            public string Symbol { get; set; }
        }
    }
}
