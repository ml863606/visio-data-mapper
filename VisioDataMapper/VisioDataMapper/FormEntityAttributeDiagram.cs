using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormEntityAttributeDiagram : Form
    {
        private const string SqlPrompt = "请根据我提供的SQL建表语句，提取实体和属性，整理为适合绘制ER实体属性图的结构。要求实体名和属性名使用中文，保留主键关系，去掉字段名中括号及括号内容，只输出结构化结果。";

        private TextBox txtPrompt;
        private Button btnCopyPrompt;
        private RichTextBox txtSql;
        private Button btnPasteSql;
        private Button btnParse;
        private Button btnGenerate;
        private Button btnClose;
        private TextBox txtStatus;
        private Label lblSummary;
        private DataGridView dgvPreview;
        private ComboBox cmbEntityFont;
        private TextBox txtEntityFontSize;
        private ComboBox cmbAttributeFont;
        private TextBox txtAttributeFontSize;
        private TextBox txtEntityWidth;
        private TextBox txtRowHeight;
        private TextBox txtLineWidth;
        private ComboBox cmbLayoutStyle;
        private Label lblEntityFont;
        private Label lblEntityFontSize;
        private Label lblAttributeFont;
        private Label lblAttributeFontSize;
        private Label lblLayoutStyle;
        private Label lblEntityWidth;
        private Label lblRowHeight;
        private Label lblLineWidth;
        private CheckBox chkPkUnderline;
        private CheckBox chkShowType;
        private CheckBox chkShowComment;
        private GroupBox grpTextStyle;
        private GroupBox grpDrawOptions;
        private Timer autoParseTimer;
        private bool isUpdatingSql;

        private List<EntityDefinition> entities = new List<EntityDefinition>();
        private string currentEntityFont = "宋体";
        private string currentAttributeFont = "宋体";
        private double currentEntityFontSize = 10.5;
        private double currentAttributeFontSize = 10.5;
        private double currentLineWidth = 0.75;

        public FormEntityAttributeDiagram()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "智能画图-实体属性图";
            Size = new Size(1120, 860);
            MinimumSize = new Size(1040, 760);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;

            Font defaultFont = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            Font = defaultFont;

            Label lblPrompt = new Label { Text = "AI整理提示词:", Location = new Point(15, 15), Size = new Size(90, 22) };
            txtPrompt = new TextBox
            {
                Text = SqlPrompt,
                Location = new Point(110, 12),
                Size = new Size(840, 62),
                Multiline = true,
                ReadOnly = true,
                BackColor = Color.White
            };
            btnCopyPrompt = new Button { Text = "快速复制", Location = new Point(970, 25), Size = new Size(110, 32) };
            btnCopyPrompt.Click += btnCopyPrompt_Click;

            Label lblSql = new Label { Text = "SQL建表语句:", Location = new Point(15, 90), Size = new Size(90, 22) };
            btnPasteSql = new Button { Text = "粘贴SQL", Location = new Point(110, 86), Size = new Size(90, 28) };
            btnPasteSql.Click += btnPasteSql_Click;
            btnParse = new Button { Text = "解析预览", Location = new Point(215, 86), Size = new Size(90, 28) };
            btnParse.Click += btnParse_Click;

            txtSql = new RichTextBox
            {
                Location = new Point(15, 120),
                Size = new Size(1060, 260),
                ScrollBars = RichTextBoxScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 10F, FontStyle.Regular)
            };

            txtStatus = new TextBox
            {
                Location = new Point(15, 390),
                Size = new Size(1060, 42),
                Multiline = true,
                ReadOnly = true,
                BackColor = SystemColors.Control,
                ForeColor = Color.DarkGray,
                Text = $"{DateTime.Now:HH:mm:ss} 请粘贴CREATE TABLE语句后解析。"
            };

            lblSummary = new Label { Text = "共 0 个实体", Location = new Point(15, 445), Size = new Size(1060, 24) };

            dgvPreview = new DataGridView
            {
                Location = new Point(15, 475),
                Size = new Size(1060, 205),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 26 }
            };
            InitializePreviewGrid();

            grpTextStyle = new GroupBox { Text = "字体设置", Location = new Point(15, 695), Size = new Size(520, 95) };
            lblEntityFont = new Label { Text = "实体字体", Location = new Point(18, 28), Size = new Size(65, 22) };
            cmbEntityFont = CreateFontCombo(new Point(90, 24));
            lblEntityFontSize = new Label { Text = "实体字号", Location = new Point(230, 28), Size = new Size(65, 22) };
            txtEntityFontSize = new TextBox { Text = "10.5", Location = new Point(300, 24), Size = new Size(60, 25) };
            lblAttributeFont = new Label { Text = "属性字体", Location = new Point(18, 64), Size = new Size(65, 22) };
            cmbAttributeFont = CreateFontCombo(new Point(90, 60));
            lblAttributeFontSize = new Label { Text = "属性字号", Location = new Point(230, 64), Size = new Size(65, 22) };
            txtAttributeFontSize = new TextBox { Text = "10.5", Location = new Point(300, 60), Size = new Size(60, 25) };

            grpDrawOptions = new GroupBox { Text = "绘图设置", Location = new Point(550, 665), Size = new Size(525, 125) };
            lblLayoutStyle = new Label { Text = "方向样式", Location = new Point(18, 28), Size = new Size(65, 22) };
            cmbLayoutStyle = new ComboBox { Location = new Point(90, 24), Size = new Size(250, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLayoutStyle.Items.AddRange(new string[]
            {
                "属性在下，围成半圆",
                "属性在下，一条直线",
                "属性在上，围成半圆",
                "属性在上，一条直线",
                "上下两侧",
                "左右两侧",
                "一圈围成圆环"
            });
            cmbLayoutStyle.SelectedIndex = 4;
            lblEntityWidth = new Label { Text = "实体宽(mm)", Location = new Point(18, 63), Size = new Size(85, 22) };
            txtEntityWidth = new TextBox { Text = "50", Location = new Point(110, 59), Size = new Size(60, 25) };
            lblRowHeight = new Label { Text = "属性高(mm)", Location = new Point(200, 63), Size = new Size(85, 22) };
            txtRowHeight = new TextBox { Text = "8", Location = new Point(290, 59), Size = new Size(60, 25) };
            lblLineWidth = new Label { Text = "线宽(pt)", Location = new Point(380, 63), Size = new Size(65, 22) };
            txtLineWidth = new TextBox { Text = "0.75", Location = new Point(450, 59), Size = new Size(60, 25) };
            chkPkUnderline = new CheckBox { Text = "主键下划线", Location = new Point(18, 96), Size = new Size(115, 22), Checked = true };
            chkShowType = new CheckBox { Text = "显示字段类型", Location = new Point(160, 96), Size = new Size(120, 22), Checked = false };
            chkShowComment = new CheckBox { Text = "优先显示注释", Location = new Point(305, 96), Size = new Size(120, 22), Checked = true };

            grpTextStyle.Controls.Add(lblEntityFont);
            grpTextStyle.Controls.Add(cmbEntityFont);
            grpTextStyle.Controls.Add(lblEntityFontSize);
            grpTextStyle.Controls.Add(txtEntityFontSize);
            grpTextStyle.Controls.Add(lblAttributeFont);
            grpTextStyle.Controls.Add(cmbAttributeFont);
            grpTextStyle.Controls.Add(lblAttributeFontSize);
            grpTextStyle.Controls.Add(txtAttributeFontSize);
            grpDrawOptions.Controls.Add(lblLayoutStyle);
            grpDrawOptions.Controls.Add(cmbLayoutStyle);
            grpDrawOptions.Controls.Add(lblEntityWidth);
            grpDrawOptions.Controls.Add(txtEntityWidth);
            grpDrawOptions.Controls.Add(lblRowHeight);
            grpDrawOptions.Controls.Add(txtRowHeight);
            grpDrawOptions.Controls.Add(lblLineWidth);
            grpDrawOptions.Controls.Add(txtLineWidth);
            grpDrawOptions.Controls.Add(chkPkUnderline);
            grpDrawOptions.Controls.Add(chkShowType);
            grpDrawOptions.Controls.Add(chkShowComment);

            btnGenerate = new Button { Text = "生成绘图", Location = new Point(820, 775), Size = new Size(120, 35), BackColor = Color.LightSkyBlue, Font = new Font(defaultFont, FontStyle.Bold) };
            btnGenerate.Click += btnGenerate_Click;
            btnClose = new Button { Text = "关闭", Location = new Point(970, 775), Size = new Size(100, 35) };
            btnClose.Click += btnClose_Click;

            Controls.Add(lblPrompt);
            Controls.Add(txtPrompt);
            Controls.Add(btnCopyPrompt);
            Controls.Add(lblSql);
            Controls.Add(btnPasteSql);
            Controls.Add(btnParse);
            Controls.Add(txtSql);
            Controls.Add(txtStatus);
            Controls.Add(lblSummary);
            Controls.Add(dgvPreview);
            Controls.Add(grpTextStyle);
            Controls.Add(grpDrawOptions);
            Controls.Add(btnGenerate);
            Controls.Add(btnClose);

            autoParseTimer = new Timer { Interval = 500 };
            autoParseTimer.Tick += autoParseTimer_Tick;
            txtSql.TextChanged += txtSql_TextChanged;
            Resize += FormEntityAttributeDiagram_Resize;
            Shown += FormEntityAttributeDiagram_Shown;
            LayoutControls();
        }

        private ComboBox CreateFontCombo(Point location)
        {
            var combo = new ComboBox { Location = location, Size = new Size(105, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(new string[] { "宋体", "黑体", "仿宋", "楷体", "微软雅黑" });
            combo.SelectedIndex = 0;
            return combo;
        }

        private void InitializePreviewGrid()
        {
            dgvPreview.Columns.Clear();
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Entity", HeaderText = "实体", FillWeight = 90 });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "EntityCount", HeaderText = "属性数", FillWeight = 55 });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Attribute", HeaderText = "属性", FillWeight = 110 });
            dgvPreview.Columns.Add(new DataGridViewCheckBoxColumn { Name = "PK", HeaderText = "主键", FillWeight = 45 });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "类型", FillWeight = 90 });
            dgvPreview.Columns.Add(new DataGridViewTextBoxColumn { Name = "Comment", HeaderText = "备注", FillWeight = 160 });
        }

        private void LayoutControls()
        {
            if (txtSql == null || dgvPreview == null)
            {
                return;
            }

            int margin = 15;
            int width = Math.Max(900, ClientSize.Width - margin * 2);
            int bottomTop = Math.Max(730, ClientSize.Height - 55);
            int optionsTop = bottomTop - 140;
            int previewTop = 475;
            int previewBottom = optionsTop - 15;

            txtPrompt.Width = Math.Max(500, width - 220);
            btnCopyPrompt.Left = margin + width - btnCopyPrompt.Width;
            txtSql.Width = width;
            txtStatus.Width = width;
            lblSummary.Width = width;
            dgvPreview.Width = width;
            dgvPreview.Height = Math.Max(180, previewBottom - previewTop);

            grpTextStyle.Top = optionsTop;
            grpTextStyle.Left = margin;
            grpTextStyle.Width = Math.Max(420, (width - 15) / 2);
            grpTextStyle.Height = 95;
            grpDrawOptions.Top = optionsTop;
            grpDrawOptions.Left = grpTextStyle.Right + 15;
            grpDrawOptions.Width = Math.Max(420, width - grpTextStyle.Width - 15);
            grpDrawOptions.Height = 125;
            btnClose.Top = bottomTop;
            btnGenerate.Top = bottomTop;
            btnClose.Left = margin + width - btnClose.Width;
            btnGenerate.Left = btnClose.Left - btnGenerate.Width - 20;
        }

        private void FormEntityAttributeDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void FormEntityAttributeDiagram_Shown(object sender, EventArgs e)
        {
            txtPrompt.SelectionStart = 0;
            txtPrompt.SelectionLength = 0;
            txtSql.Focus();
        }

        private void btnCopyPrompt_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtPrompt.Text);
            AppendLog("已复制实体属性图提示词。");
        }

        private void btnPasteSql_Click(object sender, EventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                isUpdatingSql = true;
                txtSql.Text = Clipboard.GetText();
                isUpdatingSql = false;
                ClearPreview();
                ParseAndPreview();
            }
            else
            {
                MessageBox.Show("剪贴板中没有文本内容！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnParse_Click(object sender, EventArgs e)
        {
            ParseAndPreview();
        }

        private void txtSql_TextChanged(object sender, EventArgs e)
        {
            if (isUpdatingSql)
            {
                return;
            }

            ClearPreview();
            autoParseTimer.Stop();
            if (!string.IsNullOrWhiteSpace(txtSql.Text))
            {
                autoParseTimer.Start();
            }
        }

        private void autoParseTimer_Tick(object sender, EventArgs e)
        {
            autoParseTimer.Stop();
            ParseAndPreview();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                if (entities.Count == 0)
                {
                    ParseAndPreview();
                }

                if (entities.Count == 0)
                {
                    MessageBox.Show("没有解析到实体，请确认SQL中包含CREATE TABLE语句。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                currentEntityFont = cmbEntityFont.SelectedItem == null ? "宋体" : cmbEntityFont.SelectedItem.ToString();
                currentAttributeFont = cmbAttributeFont.SelectedItem == null ? "宋体" : cmbAttributeFont.SelectedItem.ToString();
                currentEntityFontSize = ParsePositiveNumber(txtEntityFontSize.Text, "实体字号");
                currentAttributeFontSize = ParsePositiveNumber(txtAttributeFontSize.Text, "属性字号");
                currentLineWidth = ParsePositiveNumber(txtLineWidth.Text, "连接线宽");
                double entityWidth = ParsePositiveNumber(txtEntityWidth.Text, "实体宽") / 25.4;
                double rowHeight = ParsePositiveNumber(txtRowHeight.Text, "行高") / 25.4;
                string layoutStyle = cmbLayoutStyle.SelectedItem == null ? "上下两侧" : cmbLayoutStyle.SelectedItem.ToString();

                Visio.Application visioApp = Globals.ThisAddIn.Application;
                Visio.Page page = GetOrCreateActivePage(visioApp);
                ClearPageShapes(page);
                DrawEntities(page, entityWidth, rowHeight, layoutStyle);

                AppendLog("实体属性图生成完成。");
                MessageBox.Show("实体属性图生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog($"发生错误: {ex.Message}");
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ParseAndPreview()
        {
            ClearPreview();
            entities = ParseSql(txtSql.Text);
            RefreshPreview();
            AppendLog($"解析完成，共 {entities.Count} 个实体。");
        }

        private void ClearPreview()
        {
            autoParseTimer?.Stop();
            entities = new List<EntityDefinition>();
            dgvPreview.Rows.Clear();
            lblSummary.Text = "共 0 个实体";
        }

        private List<EntityDefinition> ParseSql(string sql)
        {
            var result = new List<EntityDefinition>();
            if (string.IsNullOrWhiteSpace(sql))
            {
                return result;
            }

            string text = Regex.Replace(sql, @"/\*.*?\*/", string.Empty, RegexOptions.Singleline);
            text = Regex.Replace(text, @"--.*?$", string.Empty, RegexOptions.Multiline);
            foreach (TableSqlBlock block in ExtractCreateTableBlocks(text))
            {
                string rawName = block.Name;
                string body = block.Body;
                string tail = block.Tail;
                var entity = new EntityDefinition
                {
                    RawName = rawName,
                    Name = CleanEntityName(ExtractReadableName(rawName, ExtractTableComment(tail)))
                };

                List<string> parts = SplitSqlParts(body);
                foreach (string part in parts)
                {
                    ParseTablePart(part, entity);
                }

                if (entity.Fields.Count > 0)
                {
                    result.Add(entity);
                }
            }

            return result;
        }

        private void ParseTablePart(string part, EntityDefinition entity)
        {
            string item = part.Trim();
            if (item.Length == 0)
            {
                return;
            }

            Match pkMatch = Regex.Match(item, @"^PRIMARY\s+KEY\s*\((?<cols>[^)]+)\)", RegexOptions.IgnoreCase);
            if (pkMatch.Success)
            {
                foreach (string col in SplitColumns(pkMatch.Groups["cols"].Value))
                {
                    EntityField pkField = entity.Fields.Find(f => string.Equals(f.RawName, col, StringComparison.OrdinalIgnoreCase));
                    if (pkField != null)
                    {
                        pkField.IsPrimaryKey = true;
                    }
                }
                return;
            }

            Match fkMatch = Regex.Match(item, @"FOREIGN\s+KEY\s*\((?<cols>[^)]+)\)\s+REFERENCES\s+(?<table>`[^`]+`|\[[^\]]+\]|""[^""]+""|[^\s(]+)\s*\((?<refs>[^)]+)\)", RegexOptions.IgnoreCase);
            if (fkMatch.Success)
            {
                List<string> cols = SplitColumns(fkMatch.Groups["cols"].Value);
                List<string> refs = SplitColumns(fkMatch.Groups["refs"].Value);
                string refTable = TrimSqlName(fkMatch.Groups["table"].Value);
                for (int i = 0; i < cols.Count; i++)
                {
                    EntityField fkField = entity.Fields.Find(f => string.Equals(f.RawName, cols[i], StringComparison.OrdinalIgnoreCase));
                    if (fkField != null)
                    {
                        fkField.IsForeignKey = true;
                        fkField.ReferenceTable = refTable;
                        fkField.ReferenceField = i < refs.Count ? refs[i] : string.Empty;
                    }
                }
                return;
            }

            if (Regex.IsMatch(item, @"^(KEY|INDEX|UNIQUE|CONSTRAINT|CHECK)\b", RegexOptions.IgnoreCase))
            {
                return;
            }

            Match fieldMatch = Regex.Match(item, @"^(?<name>`[^`]+`|\[[^\]]+\]|""[^""]+""|\w+)\s+(?<type>[^\s,]+(?:\s*\([^)]*\))?)(?<rest>.*)$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!fieldMatch.Success)
            {
                return;
            }

            string rawFieldName = TrimSqlName(fieldMatch.Groups["name"].Value);
            string rest = fieldMatch.Groups["rest"].Value;
            string comment = ExtractColumnComment(rest);
            var field = new EntityField
            {
                RawName = rawFieldName,
                Name = CleanAttributeName(ExtractReadableName(rawFieldName, comment)),
                Type = fieldMatch.Groups["type"].Value.Trim(),
                Comment = comment,
                IsPrimaryKey = Regex.IsMatch(rest, @"\bPRIMARY\s+KEY\b", RegexOptions.IgnoreCase)
            };
            entity.Fields.Add(field);
        }

        private List<TableSqlBlock> ExtractCreateTableBlocks(string sql)
        {
            var blocks = new List<TableSqlBlock>();
            Regex createRegex = new Regex(@"CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?(?<name>`[^`]+`|\[[^\]]+\]|""[^""]+""|[^\s(]+)", RegexOptions.IgnoreCase);

            foreach (Match match in createRegex.Matches(sql))
            {
                int openParen = sql.IndexOf('(', match.Index + match.Length);
                if (openParen < 0)
                {
                    continue;
                }

                int closeParen = FindMatchingParen(sql, openParen);
                if (closeParen < 0)
                {
                    continue;
                }

                int statementEnd = sql.IndexOf(';', closeParen + 1);
                if (statementEnd < 0)
                {
                    statementEnd = sql.Length;
                }

                blocks.Add(new TableSqlBlock
                {
                    Name = TrimSqlName(match.Groups["name"].Value),
                    Body = sql.Substring(openParen + 1, closeParen - openParen - 1),
                    Tail = sql.Substring(closeParen + 1, statementEnd - closeParen - 1)
                });
            }

            return blocks;
        }

        private int FindMatchingParen(string text, int openParen)
        {
            int depth = 0;
            char quote = '\0';

            for (int i = openParen; i < text.Length; i++)
            {
                char c = text[i];
                if (quote != '\0')
                {
                    if (c == quote && (i == 0 || text[i - 1] != '\\'))
                    {
                        quote = '\0';
                    }
                    continue;
                }

                if (c == '\'' || c == '"' || c == '`')
                {
                    quote = c;
                    continue;
                }

                if (c == '(')
                {
                    depth++;
                }
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }

        private string ExtractReadableName(string rawName, string comment)
        {
            return string.IsNullOrWhiteSpace(comment) ? rawName : comment.Trim();
        }

        private string ExtractTableComment(string tail)
        {
            Match match = Regex.Match(tail ?? string.Empty, @"COMMENT\s*=\s*['""](?<comment>.*?)['""]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups["comment"].Value : string.Empty;
        }

        private string ExtractColumnComment(string rest)
        {
            Match match = Regex.Match(rest ?? string.Empty, @"COMMENT\s+['""](?<comment>.*?)['""]", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return match.Success ? match.Groups["comment"].Value : string.Empty;
        }

        private List<string> SplitSqlParts(string body)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            int depth = 0;
            char quote = '\0';

            foreach (char c in body)
            {
                if (quote != '\0')
                {
                    current.Append(c);
                    if (c == quote)
                    {
                        quote = '\0';
                    }
                    continue;
                }

                if (c == '\'' || c == '"' || c == '`')
                {
                    quote = c;
                    current.Append(c);
                    continue;
                }

                if (c == '(') depth++;
                if (c == ')') depth--;

                if (c == ',' && depth == 0)
                {
                    result.Add(current.ToString());
                    current.Clear();
                    continue;
                }

                current.Append(c);
            }

            if (current.Length > 0)
            {
                result.Add(current.ToString());
            }

            return result;
        }

        private List<string> SplitColumns(string value)
        {
            var cols = new List<string>();
            foreach (string part in value.Split(','))
            {
                string col = TrimSqlName(part.Trim());
                if (col.Length > 0)
                {
                    cols.Add(col);
                }
            }
            return cols;
        }

        private string TrimSqlName(string name)
        {
            string value = (name ?? string.Empty).Trim();
            if ((value.StartsWith("`") && value.EndsWith("`")) ||
                (value.StartsWith("[") && value.EndsWith("]")) ||
                (value.StartsWith("\"") && value.EndsWith("\"")))
            {
                value = value.Substring(1, value.Length - 2);
            }
            return value;
        }

        private string CleanEntityName(string name)
        {
            string value = PreferChineseText(CleanBracketContent(name)).Trim();
            if (value.EndsWith("表", StringComparison.Ordinal))
            {
                value = value.Substring(0, value.Length - 1);
            }
            return value;
        }

        private string CleanAttributeName(string name)
        {
            return PreferChineseText(CleanBracketContent(name)).Trim();
        }

        private string CleanBracketContent(string value)
        {
            string result = value ?? string.Empty;
            result = Regex.Replace(result, @"（[^）]*）", string.Empty);
            result = Regex.Replace(result, @"\([^)]*\)", string.Empty);
            return result.Trim();
        }

        private string PreferChineseText(string value)
        {
            string text = value ?? string.Empty;
            MatchCollection matches = Regex.Matches(text, @"[\u4e00-\u9fa5]+");
            if (matches.Count == 0)
            {
                return text.Trim();
            }

            var builder = new StringBuilder();
            foreach (Match match in matches)
            {
                builder.Append(match.Value);
            }

            return builder.ToString();
        }

        private void RefreshPreview()
        {
            dgvPreview.Rows.Clear();
            var summary = new StringBuilder();
            summary.Append($"共 {entities.Count} 个实体");

            foreach (EntityDefinition entity in entities)
            {
                summary.Append($"；{entity.Name}：{entity.Fields.Count} 个属性");
                for (int i = 0; i < entity.Fields.Count; i++)
                {
                    EntityField field = entity.Fields[i];
                    int rowIndex = dgvPreview.Rows.Add(
                        i == 0 ? entity.Name : string.Empty,
                        i == 0 ? entity.Fields.Count.ToString(CultureInfo.InvariantCulture) : string.Empty,
                        field.Name,
                        field.IsPrimaryKey,
                        field.Type,
                        field.Comment);
                    if (i == 0)
                    {
                        dgvPreview.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252);
                    }
                }
            }

            lblSummary.Text = summary.ToString();
        }

        private void DrawEntities(Visio.Page page, double entityWidth, double rowHeight, string layoutStyle)
        {
            double pageWidth = page.PageSheet.CellsU["PageWidth"].Result["in"];
            double pageHeight = page.PageSheet.CellsU["PageHeight"].Result["in"];
            double clusterWidth = Math.Max(GetMaxClusterWidth(entityWidth, layoutStyle), 3.3);
            double clusterHeight = Math.Max(GetClusterHeight(rowHeight, layoutStyle), 2.7);
            int columns = Math.Max(1, (int)Math.Floor((pageWidth - 0.8) / clusterWidth));
            double cellWidth = pageWidth / columns;
            double startY = pageHeight - 1.1;

            for (int i = 0; i < entities.Count; i++)
            {
                EntityDefinition entity = entities[i];
                int col = i % columns;
                int row = i / columns;
                double x = cellWidth * col + cellWidth / 2.0;
                double y = startY - row * clusterHeight;
                DrawEntity(page, entity, x, y, entityWidth, rowHeight, layoutStyle);
            }
        }

        private double GetMaxClusterWidth(double configuredEntityWidth, string layoutStyle)
        {
            double maxWidth = configuredEntityWidth + 1.0;
            foreach (EntityDefinition entity in entities)
            {
                int count = Math.Max(1, entity.Fields.Count);
                var widths = new List<double>();
                foreach (EntityField field in entity.Fields)
                {
                    widths.Add(EstimateTextWidthInches(BuildFieldText(field), currentAttributeFontSize, 0.48));
                }

                double attrWidth;
                if (layoutStyle == "属性在下，一条直线" || layoutStyle == "属性在上，一条直线")
                {
                    attrWidth = GetRowWidth(widths, 0, count, 0.22);
                }
                else if (layoutStyle == "左右两侧")
                {
                    attrWidth = configuredEntityWidth + GetMaxWidth(widths, 0, (count + 1) / 2) + GetMaxWidth(widths, (count + 1) / 2, count / 2) + 1.0;
                }
                else if (layoutStyle == "属性在下，围成半圆" || layoutStyle == "属性在上，围成半圆" || layoutStyle == "一圈围成圆环")
                {
                    double maxAttrWidth = GetMaxWidth(widths, 0, count);
                    attrWidth = Math.Max(configuredEntityWidth + maxAttrWidth * 2.0 + 1.1, Math.Min(6.8, count * 0.72 + maxAttrWidth));
                }
                else
                {
                    int topCount = (count + 1) / 2;
                    int bottomCount = count - topCount;
                    double topWidth = GetRowWidth(widths, 0, topCount, 0.22);
                    double bottomWidth = bottomCount > 0 ? GetRowWidth(widths, topCount, bottomCount, 0.22) : 0;
                    attrWidth = Math.Max(topWidth, bottomWidth);
                }

                double entityNameWidth = EstimateTextWidthInches(entity.Name, currentEntityFontSize, 0.55);
                maxWidth = Math.Max(maxWidth, Math.Max(attrWidth, entityNameWidth) + 0.5);
            }

            return maxWidth;
        }

        private double GetClusterHeight(double rowHeight, string layoutStyle)
        {
            if (layoutStyle == "左右两侧")
            {
                return Math.Max(rowHeight * 9.0, 3.2);
            }

            if (layoutStyle == "属性在下，围成半圆" || layoutStyle == "属性在上，围成半圆" || layoutStyle == "一圈围成圆环")
            {
                return Math.Max(rowHeight * 9.0, 3.4);
            }

            return Math.Max(rowHeight * 6.8, 2.8);
        }

        private void DrawEntity(Visio.Page page, EntityDefinition entity, double centerX, double centerY, double width, double rowHeight, string layoutStyle)
        {
            int count = Math.Max(1, entity.Fields.Count);
            width = Math.Max(width, EstimateTextWidthInches(entity.Name, currentEntityFontSize, 0.55));
            double entityHeight = Math.Max(rowHeight * 1.25, 0.34);
            Visio.Shape entityShape = page.DrawRectangle(centerX - width / 2.0, centerY - entityHeight / 2.0, centerX + width / 2.0, centerY + entityHeight / 2.0);
            entityShape.Text = entity.Name;
            ApplyTextShapeStyle(entityShape, currentEntityFont, currentEntityFontSize, true, false);
            ApplyBoxStyle(entityShape);

            double ellipseHeight = Math.Max(0.34, rowHeight * 0.95);
            var attrWidths = new List<double>();
            var attrTexts = new List<string>();
            foreach (EntityField field in entity.Fields)
            {
                string text = BuildFieldText(field);
                attrTexts.Add(text);
                attrWidths.Add(EstimateTextWidthInches(text, currentAttributeFontSize, 0.48));
            }

            if (layoutStyle == "属性在下，一条直线")
            {
                DrawAttributeRow(page, entity, attrTexts, attrWidths, 0, entity.Fields.Count, centerX, centerY - entityHeight / 2.0 - rowHeight * 1.35, centerX, centerY - entityHeight / 2.0, ellipseHeight, false);
                return;
            }

            if (layoutStyle == "属性在上，一条直线")
            {
                DrawAttributeRow(page, entity, attrTexts, attrWidths, 0, entity.Fields.Count, centerX, centerY + entityHeight / 2.0 + rowHeight * 1.35, centerX, centerY + entityHeight / 2.0, ellipseHeight, true);
                return;
            }

            if (layoutStyle == "左右两侧")
            {
                DrawAttributeSides(page, entity, attrTexts, attrWidths, centerX, centerY, width, entityHeight, ellipseHeight, rowHeight);
                return;
            }

            if (layoutStyle == "属性在下，围成半圆")
            {
                DrawAttributeArc(page, entity, attrTexts, attrWidths, centerX, centerY, width, entityHeight, ellipseHeight, false, false);
                return;
            }

            if (layoutStyle == "属性在上，围成半圆")
            {
                DrawAttributeArc(page, entity, attrTexts, attrWidths, centerX, centerY, width, entityHeight, ellipseHeight, true, false);
                return;
            }

            if (layoutStyle == "一圈围成圆环")
            {
                DrawAttributeArc(page, entity, attrTexts, attrWidths, centerX, centerY, width, entityHeight, ellipseHeight, true, true);
                return;
            }

            int topCount = (count + 1) / 2;
            int bottomCount = count - topCount;
            double topY = centerY + entityHeight / 2.0 + rowHeight * 1.35;
            double bottomY = centerY - entityHeight / 2.0 - rowHeight * 1.35;
            DrawAttributeRow(page, entity, attrTexts, attrWidths, 0, topCount, centerX, topY, centerX, centerY + entityHeight / 2.0, ellipseHeight, true);
            if (bottomCount > 0)
            {
                DrawAttributeRow(page, entity, attrTexts, attrWidths, topCount, bottomCount, centerX, bottomY, centerX, centerY - entityHeight / 2.0, ellipseHeight, false);
            }
        }

        private void DrawAttributeRow(Visio.Page page, EntityDefinition entity, List<string> attrTexts, List<double> attrWidths, int start, int rowCount, double rowCenterX, double attrY, double lineEndX, double lineEndY, double ellipseHeight, bool isTop)
        {
            double sideGap = 0.22;
            double rowWidth = GetRowWidth(attrWidths, start, rowCount, sideGap);
            for (int i = 0; i < rowCount; i++)
            {
                int index = start + i;
                double attrWidth = attrWidths[index];
                double attrX = rowCenterX - rowWidth / 2.0 + GetWidthBefore(attrWidths, start, i, sideGap) + attrWidth / 2.0;
                double attrEdgeY = isTop ? attrY - ellipseHeight / 2.0 : attrY + ellipseHeight / 2.0;
                DrawAttribute(page, entity.Fields[index], attrTexts[index], attrX, attrY, attrWidth, ellipseHeight);
                DrawLine(page, attrX, attrEdgeY, lineEndX, lineEndY);
            }
        }

        private void DrawAttributeSides(Visio.Page page, EntityDefinition entity, List<string> attrTexts, List<double> attrWidths, double centerX, double centerY, double entityWidth, double entityHeight, double ellipseHeight, double rowHeight)
        {
            int leftCount = (entity.Fields.Count + 1) / 2;
            int rightCount = entity.Fields.Count - leftCount;
            double gap = Math.Max(rowHeight * 1.25, ellipseHeight + 0.12);
            double leftMax = GetMaxWidth(attrWidths, 0, leftCount);
            double rightMax = GetMaxWidth(attrWidths, leftCount, rightCount);
            double leftX = centerX - entityWidth / 2.0 - leftMax / 2.0 - 0.55;
            double rightX = centerX + entityWidth / 2.0 + rightMax / 2.0 + 0.55;

            DrawAttributeColumn(page, entity, attrTexts, attrWidths, 0, leftCount, leftX, centerY, gap, ellipseHeight, centerX, centerY, entityWidth, entityHeight, false);
            if (rightCount > 0)
            {
                DrawAttributeColumn(page, entity, attrTexts, attrWidths, leftCount, rightCount, rightX, centerY, gap, ellipseHeight, centerX, centerY, entityWidth, entityHeight, true);
            }
        }

        private void DrawAttributeColumn(Visio.Page page, EntityDefinition entity, List<string> attrTexts, List<double> attrWidths, int start, int count, double attrX, double centerY, double gap, double ellipseHeight, double entityCenterX, double entityCenterY, double entityWidth, double entityHeight, bool isRight)
        {
            if (count <= 0)
            {
                return;
            }

            double startY = centerY + (count - 1) * gap / 2.0;
            for (int i = 0; i < count; i++)
            {
                int index = start + i;
                double attrY = startY - i * gap;
                double attrWidth = attrWidths[index];
                double attrEdgeX = isRight ? attrX - attrWidth / 2.0 : attrX + attrWidth / 2.0;
                double entityEdgeX;
                double entityEdgeY;
                GetRectangleEdgePoint(entityCenterX, entityCenterY, entityWidth, entityHeight, attrX, attrY, out entityEdgeX, out entityEdgeY);
                DrawAttribute(page, entity.Fields[index], attrTexts[index], attrX, attrY, attrWidth, ellipseHeight);
                DrawLine(page, attrEdgeX, attrY, entityEdgeX, entityEdgeY);
            }
        }

        private void DrawAttributeArc(Visio.Page page, EntityDefinition entity, List<string> attrTexts, List<double> attrWidths, double centerX, double centerY, double entityWidth, double entityHeight, double ellipseHeight, bool upper, bool fullRing)
        {
            int count = entity.Fields.Count;
            double maxAttrWidth = GetMaxWidth(attrWidths, 0, count);
            double radiusX = Math.Max(entityWidth / 2.0 + maxAttrWidth / 2.0 + 0.55, Math.Min(3.0, count * 0.32 + 1.0));
            double radiusY = Math.Max(entityHeight / 2.0 + ellipseHeight + 0.55, fullRing ? 1.15 : 0.95);

            for (int i = 0; i < entity.Fields.Count; i++)
            {
                double angle;
                if (fullRing)
                {
                    angle = -90 + i * 360.0 / Math.Max(1, count);
                }
                else if (count == 1)
                {
                    angle = upper ? 90 : 270;
                }
                else
                {
                    double startAngle = upper ? 25 : 205;
                    double endAngle = upper ? 155 : 335;
                    angle = startAngle + (endAngle - startAngle) * i / (count - 1);
                }

                double radians = angle * Math.PI / 180.0;
                double attrWidth = attrWidths[i];
                double attrX = centerX + Math.Cos(radians) * radiusX;
                double attrY = centerY + Math.Sin(radians) * radiusY;
                double entityEdgeX;
                double entityEdgeY;
                double attrEdgeX;
                double attrEdgeY;
                GetRectangleEdgePoint(centerX, centerY, entityWidth, entityHeight, attrX, attrY, out entityEdgeX, out entityEdgeY);
                GetOvalEdgePoint(attrX, attrY, attrWidth, ellipseHeight, centerX, centerY, out attrEdgeX, out attrEdgeY);
                DrawAttribute(page, entity.Fields[i], attrTexts[i], attrX, attrY, attrWidth, ellipseHeight);
                DrawLine(page, attrEdgeX, attrEdgeY, entityEdgeX, entityEdgeY);
            }
        }

        private void DrawAttribute(Visio.Page page, EntityField field, string text, double centerX, double centerY, double width, double height)
        {
            Visio.Shape attrShape = page.DrawOval(centerX - width / 2.0, centerY - height / 2.0, centerX + width / 2.0, centerY + height / 2.0);
            attrShape.Text = text;
            ApplyTextShapeStyle(attrShape, currentAttributeFont, currentAttributeFontSize, false, chkPkUnderline.Checked && field.IsPrimaryKey);
            ApplyBoxStyle(attrShape);
        }

        private double GetMaxWidth(List<double> widths, int start, int count)
        {
            double maxWidth = 0;
            for (int i = 0; i < count && start + i < widths.Count; i++)
            {
                maxWidth = Math.Max(maxWidth, widths[start + i]);
            }

            return maxWidth;
        }

        private void GetRectangleEdgePoint(double centerX, double centerY, double width, double height, double targetX, double targetY, out double edgeX, out double edgeY)
        {
            double dx = targetX - centerX;
            double dy = targetY - centerY;
            if (Math.Abs(dx) < 0.0001 && Math.Abs(dy) < 0.0001)
            {
                edgeX = centerX;
                edgeY = centerY;
                return;
            }

            double scaleX = Math.Abs(dx) < 0.0001 ? double.MaxValue : width / 2.0 / Math.Abs(dx);
            double scaleY = Math.Abs(dy) < 0.0001 ? double.MaxValue : height / 2.0 / Math.Abs(dy);
            double scale = Math.Min(scaleX, scaleY);
            edgeX = centerX + dx * scale;
            edgeY = centerY + dy * scale;
        }

        private void GetOvalEdgePoint(double centerX, double centerY, double width, double height, double targetX, double targetY, out double edgeX, out double edgeY)
        {
            double dx = targetX - centerX;
            double dy = targetY - centerY;
            if (Math.Abs(dx) < 0.0001 && Math.Abs(dy) < 0.0001)
            {
                edgeX = centerX;
                edgeY = centerY;
                return;
            }

            double a = width / 2.0;
            double b = height / 2.0;
            double scale = 1.0 / Math.Sqrt(dx * dx / (a * a) + dy * dy / (b * b));
            edgeX = centerX + dx * scale;
            edgeY = centerY + dy * scale;
        }

        private double EstimateTextWidthInches(string text, double fontSize, double minWidth)
        {
            int count = Math.Max(1, (text ?? string.Empty).Length);
            double width = count * fontSize * 0.0105 + 0.28;
            return Math.Max(minWidth, width);
        }

        private double GetRowWidth(List<double> widths, int start, int count, double gap)
        {
            if (count <= 0 || start >= widths.Count)
            {
                return 0;
            }

            double total = 0;
            for (int i = 0; i < count; i++)
            {
                total += widths[start + i];
            }

            return total + Math.Max(0, count - 1) * gap;
        }

        private double GetWidthBefore(List<double> widths, int start, int count, double gap)
        {
            double total = 0;
            for (int i = 0; i < count; i++)
            {
                total += widths[start + i] + gap;
            }

            return total;
        }

        private string BuildFieldText(EntityField field)
        {
            string text = chkShowComment.Checked ? field.Name : CleanAttributeName(field.RawName);
            if (chkShowType.Checked && !string.IsNullOrWhiteSpace(field.Type))
            {
                text += "：" + field.Type;
            }
            return text;
        }

        private void ApplyTextShapeStyle(Visio.Shape shape, string fontName, double fontSize, bool bold, bool underline)
        {
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSize.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(0, 0, 0)");
            TrySetFormula(shape, "Char.Style", bold ? "1" : "0");
            TrySetFormula(shape, "Char.Underline", underline ? "1" : "0");
            TrySetFormula(shape, "Para.HorzAlign", "1");
            TrySetFormula(shape, "VerticalAlign", "1");
            TrySetFormula(shape, "TxtPinX", "Width*0.5");
            TrySetFormula(shape, "TxtPinY", "Height*0.5");
            TrySetFormula(shape, "TxtWidth", "Width");
            TrySetFormula(shape, "TxtHeight", "Height");
            TrySetFormula(shape, "LeftMargin", "0");
            TrySetFormula(shape, "RightMargin", "0");
            TrySetFormula(shape, "TopMargin", "0");
            TrySetFormula(shape, "BottomMargin", "0");
        }

        private void ApplyBoxStyle(Visio.Shape shape)
        {
            TrySetFormula(shape, "FillPattern", "0");
            TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(shape, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
        }

        private void DrawLine(Visio.Page page, double x1, double y1, double x2, double y2)
        {
            Visio.Shape line = page.DrawLine(x1, y1, x2, y2);
            TrySetFormula(line, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(line, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(line, "BeginArrow", "0");
            TrySetFormula(line, "EndArrow", "0");
        }

        private void TrySetFormula(Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                shape.CellsU[cellName].FormulaU = formula;
            }
            catch
            {
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
            catch
            {
            }

            Visio.Document newDoc = visioApp.Documents.Add("");
            return newDoc.Pages[1];
        }

        private void ClearPageShapes(Visio.Page page)
        {
            for (int i = page.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    page.Shapes[i].Delete();
                }
                catch
                {
                }
            }
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

        private void AppendLog(string message)
        {
            txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now:HH:mm:ss} {message}");
        }

        private class EntityDefinition
        {
            public string RawName { get; set; }
            public string Name { get; set; }
            public List<EntityField> Fields { get; } = new List<EntityField>();
        }

        private class TableSqlBlock
        {
            public string Name { get; set; }
            public string Body { get; set; }
            public string Tail { get; set; }
        }

        private class EntityField
        {
            public string RawName { get; set; }
            public string Name { get; set; }
            public string Type { get; set; }
            public string Comment { get; set; }
            public bool IsPrimaryKey { get; set; }
            public bool IsForeignKey { get; set; }
            public string ReferenceTable { get; set; }
            public string ReferenceField { get; set; }
        }
    }
}
