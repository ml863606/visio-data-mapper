using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormModuleDiagram : Form
    {
        private Button btnPasteOptimize;
        private Label lblPrompt;
        private TextBox txtPrompt;
        private Button btnCopyPrompt;
        private Label lblTitle;
        private TextBox txtTitle;
        private RichTextBox txtStructure;
        private TextBox txtStatus;

        private Label lblDirection;
        private ComboBox cmbDirection;
        private Label lblShapeHeight;
        private TextBox txtShapeHeight;
        private Label lblHorSpacing;
        private TextBox txtHorSpacing;
        private Label lblVerSpacing;
        private TextBox txtVerSpacing;

        private Label lblLevelCount;
        private DataGridView dgvLevels;
        private Label lblFontName;
        private ComboBox cmbFontName;
        private Label lblFontSize;
        private TextBox txtFontSize;
        private Label lblLineWidth;
        private TextBox txtLineWidth;
        private CheckBox chkArrow;

        private Button btnGenerate;
        private Button btnClose;
        private List<LevelOption> levelOptions = new List<LevelOption>();
        private string currentFontName = "宋体";
        private double currentFontSizePt = 10.5;
        private double currentLineWidthPt = 0.75;
        
        private const string StructurePrompt = "上面是我的功能模块资料，根据资料里的内容，帮我输出一个功能模块图 JSON 格式数据。要求格式如下：\n{\n  \"sys_name\": \"系统名称\",\n  \"platform\": [\n    {\n      \"name\": \"平台/角色名称\",\n      \"module\": [\n        {\n          \"name\": \"模块名称\",\n          \"sub\": [\n            \"功能名称1\",\n            \"功能名称2\"\n          ]\n        }\n      ]\n    }\n  ]\n}";

        public FormModuleDiagram()
        {
            InitializeComponent();
            LoadSampleData();
            RefreshLevelOptions();
        }

        private void InitializeComponent()
        {
            this.Text = "智能画图-功能模块图";
            this.Size = new Size(1100, 860);
            this.MinimumSize = new Size(1100, 760);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;

            Color purpleTheme = Color.FromArgb(102, 45, 145);
            Font defaultFont = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            this.Font = defaultFont;

            lblPrompt = new Label { Text = "AI整理提示词:", Location = new Point(15, 15), Size = new Size(90, 20) };
            txtPrompt = new TextBox
            {
                Text = StructurePrompt,
                Location = new Point(110, 12),
                Size = new Size(820, 70),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White
            };
            btnCopyPrompt = new Button { Text = "快速复制", Location = new Point(945, 28), Size = new Size(110, 32) };
            btnCopyPrompt.Click += btnCopyPrompt_Click;

            Label lblTutorialTip = new Label { Text = "请将 JSON 结构内容粘贴到下方", Location = new Point(15, 95), AutoSize = true };

            btnPasteOptimize = new Button { Text = "粘贴并格式化 JSON", Location = new Point(15, 120), Size = new Size(140, 28) };
            btnPasteOptimize.Click += btnPasteOptimize_Click;

            lblTitle = new Label { Text = "图标题:", Location = new Point(180, 125), Size = new Size(55, 20) };
            txtTitle = new TextBox { Text = "功能模块图", Location = new Point(240, 121), Size = new Size(360, 25) };

            txtStructure = new RichTextBox { Location = new Point(15, 160), Size = new Size(1040, 360), ScrollBars = RichTextBoxScrollBars.Both, WordWrap = false, Font = new Font("Consolas", 11F, FontStyle.Regular) };
            txtStructure.TextChanged += txtStructure_TextChanged;

            // Status ReadOnly TextBox
            txtStatus = new TextBox
            {
                Location = new Point(15, 510),
                Size = new Size(1040, 45),
                Multiline = true,
                ReadOnly = true,
                BackColor = SystemColors.Control,
                ForeColor = Color.DarkGray,
                Text = $"{DateTime.Now.ToString("HH:mm:ss")} 初次使用智能画图，建议重启1-2次Visio，确保插件完全加载~"
            };

            lblDirection = new Label { Text = "绘图方向:", Location = new Point(15, 570), Size = new Size(70, 22) };
            cmbDirection = new ComboBox { Location = new Point(85, 566), Size = new Size(120, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDirection.Items.AddRange(new string[] { "从上到下", "从左到右" });
            cmbDirection.SelectedIndex = 0;

            lblShapeHeight = new Label { Text = "横向默认高(mm):", Location = new Point(225, 570), Size = new Size(105, 22) };
            txtShapeHeight = new TextBox { Text = "10", Location = new Point(335, 566), Size = new Size(60, 25) };

            lblHorSpacing = new Label { Text = "横向间距(mm):", Location = new Point(415, 570), Size = new Size(100, 22) };
            txtHorSpacing = new TextBox { Text = "12", Location = new Point(520, 566), Size = new Size(60, 25) };

            lblVerSpacing = new Label { Text = "纵向间距(mm):", Location = new Point(600, 570), Size = new Size(100, 22) };
            txtVerSpacing = new TextBox { Text = "14", Location = new Point(705, 566), Size = new Size(60, 25) };

            lblFontName = new Label { Text = "字体:", Location = new Point(785, 570), Size = new Size(40, 22) };
            cmbFontName = new ComboBox { Location = new Point(830, 566), Size = new Size(110, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontName.Items.AddRange(new string[] { "宋体", "黑体", "仿宋", "楷体", "微软雅黑" });
            cmbFontName.SelectedIndex = 0;

            lblFontSize = new Label { Text = "字号:", Location = new Point(955, 570), Size = new Size(40, 22) };
            txtFontSize = new TextBox { Text = "10.5", Location = new Point(1000, 566), Size = new Size(55, 25) };

            lblLevelCount = new Label { Text = "共 0 个层级", Location = new Point(15, 605), Size = new Size(180, 22) };

            dgvLevels = new DataGridView
            {
                Location = new Point(15, 630),
                Size = new Size(1040, 115),
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 28 }
            };
            InitializeLevelGrid();

            lblLineWidth = new Label { Text = "连接线宽(pt):", Location = new Point(15, 765), Size = new Size(95, 22) };
            txtLineWidth = new TextBox { Text = "0.75", Location = new Point(115, 761), Size = new Size(60, 25) };
            chkArrow = new CheckBox { Text = "连接线带箭头", Location = new Point(200, 764), Size = new Size(130, 22), Checked = false };

            btnGenerate = new Button { Text = "生成绘图", Location = new Point(795, 755), Size = new Size(120, 35), BackColor = Color.LightSkyBlue, Font = new Font(defaultFont, FontStyle.Bold) };
            btnGenerate.Click += btnGenerate_Click;

            btnClose = new Button { Text = "关闭", Location = new Point(955, 755), Size = new Size(100, 35) };
            btnClose.Click += btnClose_Click;

            // Add all controls to form
            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtPrompt);
            this.Controls.Add(btnCopyPrompt);
            this.Controls.Add(lblTutorialTip);
            this.Controls.Add(btnPasteOptimize);
            this.Controls.Add(lblTitle);
            this.Controls.Add(txtTitle);
            this.Controls.Add(txtStructure);
            this.Controls.Add(txtStatus);

            this.Controls.Add(lblDirection);
            this.Controls.Add(cmbDirection);
            this.Controls.Add(lblShapeHeight);
            this.Controls.Add(txtShapeHeight);
            this.Controls.Add(lblHorSpacing);
            this.Controls.Add(txtHorSpacing);
            this.Controls.Add(lblVerSpacing);
            this.Controls.Add(txtVerSpacing);

            this.Controls.Add(lblFontName);
            this.Controls.Add(cmbFontName);
            this.Controls.Add(lblFontSize);
            this.Controls.Add(txtFontSize);
            this.Controls.Add(lblLevelCount);
            this.Controls.Add(dgvLevels);
            this.Controls.Add(lblLineWidth);
            this.Controls.Add(txtLineWidth);
            this.Controls.Add(chkArrow);
            this.Controls.Add(btnGenerate);
            this.Controls.Add(btnClose);

            this.Resize += FormModuleDiagram_Resize;
            this.Shown += FormModuleDiagram_Shown;
            LayoutControls();
        }

        private void LoadSampleData()
        {
            txtStructure.Text = @"{
  ""sys_name"": ""XXXXXXXXX系统"",
  ""platform"": [
    {
      ""name"": ""管理员"",
      ""module"": [
        {
          ""name"": ""管理员模块1"",
          ""sub"": [
            ""管理员功能功能1"",
            ""管理员功能功能2"",
            ""管理员功能功能3"",
            ""管理员功能功能4""
          ]
        },
        {
          ""name"": ""管理员模块2"",
          ""sub"": [
            ""管理员功能功能21"",
            ""管理员功能功能22"",
            ""管理员功能功能23"",
            ""管理员功能功能24""
          ]
        }
      ]
    },
    {
      ""name"": ""用户"",
      ""module"": [
        {
          ""name"": ""用户模块1"",
          ""sub"": [
            ""用户功能功能1"",
            ""用户功能功能2"",
            ""用户功能功能3"",
            ""用户功能功能4""
          ]
        },
        {
          ""name"": ""用户模块2"",
          ""sub"": [
            ""用户功能功能21"",
            ""用户功能功能22"",
            ""用户功能功能23"",
            ""用户功能功能24""
          ]
        }
      ]
    }
  ]
}";
        }

        private void InitializeLevelGrid()
        {
            dgvLevels.Columns.Clear();
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "Level", HeaderText = "层级", ReadOnly = true, FillWeight = 45 });
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "NodeCount", HeaderText = "节点数", ReadOnly = true, FillWeight = 55 });
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "MaxText", HeaderText = "最长文字", ReadOnly = true, FillWeight = 110 });

            var directionColumn = new DataGridViewComboBoxColumn { Name = "Direction", HeaderText = "形状排列", FillWeight = 90 };
            directionColumn.Items.AddRange("横着排列", "竖着排列");
            dgvLevels.Columns.Add(directionColumn);

            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "Width", HeaderText = "宽(mm)", FillWeight = 70 });
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "Height", HeaderText = "高(mm)", FillWeight = 70 });
        }

        private void FormModuleDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void FormModuleDiagram_Shown(object sender, EventArgs e)
        {
            txtPrompt.SelectionStart = 0;
            txtPrompt.SelectionLength = 0;
            txtStructure.Focus();
        }

        private void LayoutControls()
        {
            if (txtStructure == null || dgvLevels == null)
            {
                return;
            }

            int margin = 15;
            int width = Math.Max(760, this.ClientSize.Width - margin * 2);
            int bottomButtonTop = Math.Max(710, this.ClientSize.Height - 55);
            int gridTop = bottomButtonTop - 170;
            int optionsTop = gridTop - 60;
            int statusTop = optionsTop - 55;
            int structureTop = 160;
            int structureHeight = Math.Max(250, statusTop - structureTop - 10);

            txtPrompt.Width = Math.Max(500, width - lblPrompt.Width - btnCopyPrompt.Width - 25);
            txtPrompt.Height = 70;
            btnCopyPrompt.Left = margin + width - btnCopyPrompt.Width;

            txtStructure.Width = width;
            txtStructure.Height = structureHeight;

            txtStatus.Top = txtStructure.Bottom + 10;
            txtStatus.Width = width;

            int rowTop = txtStatus.Bottom + 12;
            lblDirection.Top = rowTop + 4;
            cmbDirection.Top = rowTop;
            lblShapeHeight.Top = rowTop + 4;
            txtShapeHeight.Top = rowTop;
            lblHorSpacing.Top = rowTop + 4;
            txtHorSpacing.Top = rowTop;
            lblVerSpacing.Top = rowTop + 4;
            txtVerSpacing.Top = rowTop;
            lblFontName.Top = rowTop + 4;
            cmbFontName.Top = rowTop;
            lblFontSize.Top = rowTop + 4;
            txtFontSize.Top = rowTop;

            int rightEdge = margin + width;
            txtFontSize.Left = rightEdge - txtFontSize.Width;
            lblFontSize.Left = txtFontSize.Left - lblFontSize.Width - 5;
            cmbFontName.Left = lblFontSize.Left - cmbFontName.Width - 15;
            lblFontName.Left = cmbFontName.Left - lblFontName.Width - 5;

            lblLevelCount.Top = rowTop + 35;
            dgvLevels.Top = lblLevelCount.Bottom + 5;
            dgvLevels.Width = width;
            dgvLevels.Height = Math.Max(125, bottomButtonTop - dgvLevels.Top - 10);

            lblLineWidth.Top = bottomButtonTop + 8;
            txtLineWidth.Top = bottomButtonTop + 4;
            chkArrow.Top = bottomButtonTop + 7;
            btnGenerate.Top = bottomButtonTop;
            btnClose.Top = bottomButtonTop;
            btnClose.Left = rightEdge - btnClose.Width;
            btnGenerate.Left = btnClose.Left - btnGenerate.Width - 20;
        }

        private void btnCopyPrompt_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(txtPrompt.Text);
            AppendLog("已复制功能模块图 JSON 提示词。");
        }

        private void txtStructure_TextChanged(object sender, EventArgs e)
        {
            RefreshLevelOptions();
        }

        private void RefreshLevelOptions()
        {
            if (dgvLevels == null)
            {
                return;
            }

            Dictionary<int, LevelOption> existingOptions = ReadExistingLevelOptionsFromGrid();
            TreeNode root = null;
            try
            {
                root = ParseTree(txtStructure.Text);
            }
            catch
            {
                // Parse fails silently during typing
            }

            if (root == null)
            {
                return;
            }

            var stats = new SortedDictionary<int, LevelStats>();
            CollectLevelStats(root, 1, stats);

            dgvLevels.Rows.Clear();
            levelOptions.Clear();
            lblLevelCount.Text = $"共 {stats.Count} 个层级";

            foreach (var item in stats)
            {
                int level = item.Key;
                LevelStats stat = item.Value;
                bool isLastLevel = level == stats.Count;
                LevelOption existingOption;
                bool hasExistingOption = existingOptions.TryGetValue(level, out existingOption);
                bool isVertical = hasExistingOption ? existingOption.IsVertical : isLastLevel;
                double widthMm = hasExistingOption ? existingOption.WidthMm : (isVertical ? 14 : EstimateHorizontalWidthMm(stat.LongestText));
                double heightMm = hasExistingOption ? existingOption.HeightMm : (isVertical ? Math.Max(35, GetChineseCharCount(stat.LongestText) * 6.0) : 10);

                var option = new LevelOption
                {
                    Level = level,
                    NodeCount = stat.NodeCount,
                    MaxText = stat.LongestText,
                    IsVertical = isVertical,
                    WidthMm = widthMm,
                    HeightMm = heightMm
                };
                levelOptions.Add(option);

                dgvLevels.Rows.Add(
                    level,
                    stat.NodeCount,
                    stat.LongestText,
                    isVertical ? "竖着排列" : "横着排列",
                    widthMm.ToString("0.#", CultureInfo.InvariantCulture),
                    heightMm.ToString("0.#", CultureInfo.InvariantCulture));
            }
        }

        private void CollectLevelStats(TreeNode node, int level, SortedDictionary<int, LevelStats> stats)
        {
            if (node == null)
            {
                return;
            }

            LevelStats stat;
            if (!stats.TryGetValue(level, out stat))
            {
                stat = new LevelStats();
                stats[level] = stat;
            }

            stat.NodeCount++;
            if ((node.Text ?? string.Empty).Length > (stat.LongestText ?? string.Empty).Length)
            {
                stat.LongestText = node.Text;
            }

            foreach (TreeNode child in node.Children)
            {
                CollectLevelStats(child, level + 1, stats);
            }
        }

        private double EstimateHorizontalWidthMm(string text)
        {
            return Math.Max(28, GetChineseCharCount(text) * 5.2 + 12);
        }

        private int GetChineseCharCount(string text)
        {
            return Math.Max(1, (text ?? string.Empty).Trim().Length);
        }

        private Dictionary<int, LevelOption> ReadExistingLevelOptionsFromGrid()
        {
            var result = new Dictionary<int, LevelOption>();
            if (dgvLevels == null)
            {
                return result;
            }

            foreach (DataGridViewRow row in dgvLevels.Rows)
            {
                if (row.IsNewRow || row.Cells["Level"].Value == null)
                {
                    continue;
                }

                int level;
                double widthMm;
                double heightMm;
                if (!int.TryParse(Convert.ToString(row.Cells["Level"].Value), out level) ||
                    !TryParsePositiveNumber(Convert.ToString(row.Cells["Width"].Value), out widthMm) ||
                    !TryParsePositiveNumber(Convert.ToString(row.Cells["Height"].Value), out heightMm))
                {
                    continue;
                }

                result[level] = new LevelOption
                {
                    Level = level,
                    NodeCount = row.Cells["NodeCount"].Value == null ? 0 : Convert.ToInt32(row.Cells["NodeCount"].Value),
                    MaxText = Convert.ToString(row.Cells["MaxText"].Value),
                    IsVertical = Convert.ToString(row.Cells["Direction"].Value) == "竖着排列",
                    WidthMm = widthMm,
                    HeightMm = heightMm
                };
            }

            return result;
        }

        private void btnPasteOptimize_Click(object sender, EventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                txtStructure.Text = FormatJson(Clipboard.GetText());
                AppendLog("成功从剪贴板粘贴文本并优化 JSON 格式。");
            }
            else
            {
                MessageBox.Show("剪贴板中没有文本内容！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private string FormatJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return string.Empty;
            try
            {
                // 尝试反序列化以验证其是合法的 JSON
                var serializer = new JavaScriptSerializer();
                serializer.Deserialize<object>(json);
                return JsonPrettyPrint(json);
            }
            catch
            {
                return json; // 格式错误时原样返回
            }
        }

        private string JsonPrettyPrint(string json)
        {
            if (string.IsNullOrEmpty(json)) return string.Empty;

            int indent = 0;
            bool inQuotes = false;
            var sb = new System.Text.StringBuilder();

            for (int i = 0; i < json.Length; i++)
            {
                char ch = json[i];
                switch (ch)
                {
                    case '"':
                        sb.Append(ch);
                        bool escaped = false;
                        int j = i - 1;
                        while (j >= 0 && json[j] == '\\')
                        {
                            escaped = !escaped;
                            j--;
                        }
                        if (!escaped) inQuotes = !inQuotes;
                        break;
                    case '{':
                    case '[':
                        sb.Append(ch);
                        if (!inQuotes)
                        {
                            sb.AppendLine();
                            indent++;
                            sb.Append(new string(' ', indent * 2));
                        }
                        break;
                    case '}':
                    case ']':
                        if (!inQuotes)
                        {
                            sb.AppendLine();
                            indent--;
                            sb.Append(new string(' ', indent * 2));
                        }
                        sb.Append(ch);
                        break;
                    case ',':
                        sb.Append(ch);
                        if (!inQuotes)
                        {
                            sb.AppendLine();
                            sb.Append(new string(' ', indent * 2));
                        }
                        break;
                    case ':':
                        sb.Append(ch);
                        if (!inQuotes) sb.Append(" ");
                        break;
                    default:
                        if (!char.IsWhiteSpace(ch) || inQuotes)
                        {
                            sb.Append(ch);
                        }
                        break;
                }
            }
            return sb.ToString().Trim();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AppendLog(string message)
        {
            txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now.ToString("HH:mm:ss")} {message}");
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                AppendLog("开始解析结构内容...");
                TreeNode root = ParseTree(txtStructure.Text);
                if (root == null)
                {
                    MessageBox.Show("无法解析结构内容，请确保 JSON 文本非空并且格式符合要求！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                AppendLog("获取 Visio 实例并准备画布...");
                Visio.Application visioApp = Globals.ThisAddIn.Application;
                Visio.Page activePage = GetOrCreateActivePage(visioApp);
                ClearPageShapes(activePage);

                // Read parameters
                double shapeHeightVal = ParsePositiveMillimeters(txtShapeHeight.Text, "形状高度") / 25.4; // convert mm to inches
                double horSpacingVal = ParsePositiveMillimeters(txtHorSpacing.Text, "横向间距") / 25.4;
                double verSpacingVal = ParsePositiveMillimeters(txtVerSpacing.Text, "纵向间距") / 25.4;
                Dictionary<int, LevelOption> optionsByLevel = ReadLevelOptionsFromGrid();
                currentFontName = cmbFontName.SelectedItem == null ? "宋体" : cmbFontName.SelectedItem.ToString();
                currentFontSizePt = ParsePositiveNumber(txtFontSize.Text, "字号");
                currentLineWidthPt = ParsePositiveNumber(txtLineWidth.Text, "连接线宽度");
                bool isLeftToRight = cmbDirection.SelectedItem.ToString() == "从左到右";

                // Layout settings
                double defaultShapeWidth = 35.0 / 25.4; // 35mm default width

                AppendLog("运行层级布局引擎算法...");
                // Step 1: Calculate subtree sizes
                CalculateSubtreeSizes(root, defaultShapeWidth, shapeHeightVal, horSpacingVal, verSpacingVal, optionsByLevel, isLeftToRight, 1);

                // Page dimensions
                double pageWidth = activePage.PageSheet.CellsU["PageWidth"].Result["in"];
                double pageHeight = activePage.PageSheet.CellsU["PageHeight"].Result["in"];

                // Set root coordinates
                if (isLeftToRight)
                {
                    root.X = 1.0; // 1 inch from left
                    root.Y = pageHeight / 2.0; // vertical center
                }
                else
                {
                    root.X = pageWidth / 2.0; // horizontal center
                    root.Y = pageHeight - 1.45; // leave room for framed title
                }

                // Step 2: Compute absolute coordinates
                AssignCoordinates(root, defaultShapeWidth, shapeHeightVal, horSpacingVal, verSpacingVal, optionsByLevel, isLeftToRight, 1);

                AppendLog("在 Visio 中绘制形状和连线...");
                // Draw nodes and connect them
                DrawTree(activePage, root, defaultShapeWidth, shapeHeightVal, optionsByLevel, isLeftToRight, 1);

                // Add main title to the drawing if specified
                if (!string.IsNullOrWhiteSpace(txtTitle.Text))
                {
                    double titleY = pageHeight - 0.4;
                    double titleX = pageWidth / 2.0;
                    double titleHeight = 0.44;
                    Visio.Shape titleShape = activePage.DrawRectangle(titleX - 1.2, titleY - 0.22, titleX + 1.2, titleY + 0.22);
                    titleShape.Text = txtTitle.Text;
                    ApplyAcademicShapeStyle(titleShape);
                    titleShape.CellsU["Char.Size"].FormulaU = $"{(currentFontSizePt + 1).ToString(CultureInfo.InvariantCulture)}pt";
                    titleShape.CellsU["Char.Style"].Formula = "1"; // Bold
                    DrawTitleConnector(activePage, titleX, titleY - titleHeight / 2.0, root.X, root.Y + shapeHeightVal / 2.0);
                }

                AppendLog("生成绘图完成！");
                MessageBox.Show("功能模块图生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog($"发生错误: {ex.Message}");
                MessageBox.Show($"绘图失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                // Visio throws when there is no active document/page. Create a blank page below.
            }

            Visio.Document newDoc = visioApp.Documents.Add("");
            return newDoc.Pages[1];
        }

        private void ClearPageShapes(Visio.Page page)
        {
            if (page == null)
            {
                return;
            }

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

        private double ParsePositiveMillimeters(string value, string fieldName)
        {
            return ParsePositiveNumber(value, fieldName);
        }

        private double ParsePositiveNumber(string value, string fieldName)
        {
            double parsedValue;
            if (!TryParsePositiveNumber(value, out parsedValue))
            {
                throw new ArgumentException($"{fieldName}必须是数字。");
            }

            return parsedValue;
        }

        private bool TryParsePositiveNumber(string value, out double parsedValue)
        {
            string normalizedValue = (value ?? string.Empty).Trim().Replace('，', '.');
            if (!double.TryParse(normalizedValue, NumberStyles.Float, CultureInfo.CurrentCulture, out parsedValue) &&
                !double.TryParse(normalizedValue, NumberStyles.Float, CultureInfo.InvariantCulture, out parsedValue))
            {
                return false;
            }

            return parsedValue > 0;
        }

        private Dictionary<int, LevelOption> ReadLevelOptionsFromGrid()
        {
            var result = new Dictionary<int, LevelOption>();

            foreach (DataGridViewRow row in dgvLevels.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                int level = Convert.ToInt32(row.Cells["Level"].Value);
                string direction = Convert.ToString(row.Cells["Direction"].Value);
                double widthMm = ParsePositiveMillimeters(Convert.ToString(row.Cells["Width"].Value), $"第{level}层宽度");
                double heightMm = ParsePositiveMillimeters(Convert.ToString(row.Cells["Height"].Value), $"第{level}层高度");

                result[level] = new LevelOption
                {
                    Level = level,
                    NodeCount = Convert.ToInt32(row.Cells["NodeCount"].Value),
                    MaxText = Convert.ToString(row.Cells["MaxText"].Value),
                    IsVertical = direction == "竖着排列",
                    WidthMm = widthMm,
                    HeightMm = heightMm
                };
            }

            return result;
        }

        #region Tree Parser & Layout Algorithm

        public class TreeNode
        {
            public string Text { get; set; }
            public List<TreeNode> Children { get; set; } = new List<TreeNode>();
            public TreeNode Parent { get; set; }

            // Layout coordinates in inches
            public double X { get; set; }
            public double Y { get; set; }

            // Width or Height of this node's subtree (depends on orientation)
            public double SubtreeSize { get; set; }
            public Visio.Shape VisioShape { get; set; }
        }

        private class LevelOption
        {
            public int Level { get; set; }
            public int NodeCount { get; set; }
            public string MaxText { get; set; }
            public bool IsVertical { get; set; }
            public double WidthMm { get; set; }
            public double HeightMm { get; set; }
        }

        private class LevelStats
        {
            public int NodeCount { get; set; }
            public string LongestText { get; set; } = string.Empty;
        }

        // Strongly typed JSON structures
        public class ModuleData
        {
            public string sys_name { get; set; }
            public List<PlatformData> platform { get; set; }
        }

        public class PlatformData
        {
            public string name { get; set; }
            public List<ModuleInfo> module { get; set; }
        }

        public class ModuleInfo
        {
            public string name { get; set; }
            public List<string> sub { get; set; }
        }

        private TreeNode ParseTree(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            try
            {
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<ModuleData>(text);
                if (data == null) return null;

                // Set diagram title to sys_name if user hasn't modified the title box
                if (txtTitle.Text == "功能模块图" && !string.IsNullOrWhiteSpace(data.sys_name))
                {
                    txtTitle.Text = data.sys_name;
                }

                TreeNode root = new TreeNode { Text = data.sys_name ?? "系统" };

                if (data.platform != null)
                {
                    foreach (var plat in data.platform)
                    {
                        TreeNode platNode = new TreeNode { Text = plat.name, Parent = root };
                        root.Children.Add(platNode);

                        if (plat.module != null)
                        {
                            foreach (var mod in plat.module)
                            {
                                TreeNode modNode = new TreeNode { Text = mod.name, Parent = platNode };
                                platNode.Children.Add(modNode);

                                if (mod.sub != null)
                                {
                                    foreach (var sub in mod.sub)
                                    {
                                        TreeNode subNode = new TreeNode { Text = sub, Parent = modNode };
                                        modNode.Children.Add(subNode);
                                    }
                                }
                            }
                        }
                    }
                }

                return root;
            }
            catch (Exception ex)
            {
                throw new Exception("JSON解析失败，请检查格式是否正确。错误信息: " + ex.Message);
            }
        }

        private void CalculateSubtreeSizes(TreeNode node, double defaultWidth, double defaultHeight, double horSpacing, double verSpacing, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level)
        {
            double width;
            double height;
            GetShapeSize(level, defaultWidth, defaultHeight, optionsByLevel, out width, out height);
            double selfSize = isLeftToRight ? height : width;

            if (node.Children.Count == 0)
            {
                node.SubtreeSize = selfSize;
                return;
            }

            double childrenSize = 0;
            foreach (var child in node.Children)
            {
                CalculateSubtreeSizes(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1);
                childrenSize += child.SubtreeSize;
            }

            double spacing = isLeftToRight ? verSpacing : horSpacing;
            childrenSize += (node.Children.Count - 1) * spacing;

            node.SubtreeSize = Math.Max(selfSize, childrenSize);
        }

        private void AssignCoordinates(TreeNode node, double defaultWidth, double defaultHeight, double horSpacing, double verSpacing, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level)
        {
            if (node.Children.Count == 0) return;

            double spacing = isLeftToRight ? verSpacing : horSpacing;
            double totalSize = 0;
            foreach (var child in node.Children)
            {
                totalSize += child.SubtreeSize;
            }
            totalSize += (node.Children.Count - 1) * spacing;

            if (isLeftToRight)
            {
                double nodeWidth;
                double nodeHeight;
                GetShapeSize(level, defaultWidth, defaultHeight, optionsByLevel, out nodeWidth, out nodeHeight);

                // X axis goes deep, Y axis is distributed
                double currentY = node.Y + totalSize / 2.0;
                foreach (var child in node.Children)
                {
                    double childWidth;
                    double childHeight;
                    GetShapeSize(level + 1, defaultWidth, defaultHeight, optionsByLevel, out childWidth, out childHeight);
                    child.X = node.X + nodeWidth / 2.0 + horSpacing + childWidth / 2.0;
                    child.Y = currentY - child.SubtreeSize / 2.0;
                    currentY -= child.SubtreeSize + verSpacing;

                    AssignCoordinates(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1);
                }
            }
            else
            {
                double nodeWidth;
                double nodeHeight;
                GetShapeSize(level, defaultWidth, defaultHeight, optionsByLevel, out nodeWidth, out nodeHeight);

                // Y axis goes deep (down), X axis is distributed
                double currentX = node.X - totalSize / 2.0;
                foreach (var child in node.Children)
                {
                    double childWidth;
                    double childHeight;
                    GetShapeSize(level + 1, defaultWidth, defaultHeight, optionsByLevel, out childWidth, out childHeight);
                    child.X = currentX + child.SubtreeSize / 2.0;
                    child.Y = node.Y - nodeHeight / 2.0 - verSpacing - childHeight / 2.0;
                    currentX += child.SubtreeSize + horSpacing;

                    AssignCoordinates(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1);
                }
            }
        }

        private void DrawTree(Visio.Page page, TreeNode node, double defaultWidth, double defaultHeight, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level)
        {
            double width;
            double height;
            GetShapeSize(level, defaultWidth, defaultHeight, optionsByLevel, out width, out height);

            // Draw current node rectangle
            double xLeft = node.X - width / 2.0;
            double xRight = node.X + width / 2.0;
            double yBottom = node.Y - height / 2.0;
            double yTop = node.Y + height / 2.0;

            node.VisioShape = page.DrawRectangle(xLeft, yBottom, xRight, yTop);
            bool isVerticalText = ShouldUseVerticalText(level, optionsByLevel);
            node.VisioShape.Text = isVerticalText ? string.Empty : node.Text;
            ApplyAcademicShapeStyle(node.VisioShape);
            if (isVerticalText)
            {
                DrawTopDownCharacters(page, node.Text, node.X, node.Y, width, height);
            }

            // Connect parent to child
            if (node.Parent != null && node.Parent.VisioShape != null)
            {
                DrawConnector(page, node.Parent, node, defaultWidth, defaultHeight, optionsByLevel, isLeftToRight, level - 1, level);
            }

            // Draw children
            foreach (var child in node.Children)
            {
                DrawTree(page, child, defaultWidth, defaultHeight, optionsByLevel, isLeftToRight, level + 1);
            }
        }

        private bool ShouldUseVerticalText(int level, Dictionary<int, LevelOption> optionsByLevel)
        {
            LevelOption option;
            return optionsByLevel.TryGetValue(level, out option) && option.IsVertical;
        }

        private void ApplyAcademicShapeStyle(Visio.Shape shape)
        {
            shape.CellsU["FillPattern"].Formula = "0";
            shape.CellsU["LineColor"].Formula = "RGB(0, 0, 0)";
            shape.CellsU["LineWeight"].Formula = "0.75pt";
            shape.CellsU["Char.Color"].Formula = "RGB(0, 0, 0)";
            shape.CellsU["Char.Size"].FormulaU = $"{currentFontSizePt.ToString(CultureInfo.InvariantCulture)}pt";
            shape.CellsU["Char.Font"].FormulaU = $"\"{currentFontName}\"";
            shape.CellsU["Para.HorzAlign"].Formula = "1";
            shape.CellsU["VerticalAlign"].Formula = "1";
            TrySetFormula(shape, "TxtPinX", "Width*0.5");
            TrySetFormula(shape, "TxtPinY", "Height*0.5");
            TrySetFormula(shape, "TxtWidth", "Width");
            TrySetFormula(shape, "TxtHeight", "Height");
            TrySetFormula(shape, "LeftMargin", "0");
            TrySetFormula(shape, "RightMargin", "0");
            TrySetFormula(shape, "TopMargin", "0");
            TrySetFormula(shape, "BottomMargin", "0");
        }

        private void DrawTopDownCharacters(Visio.Page page, string text, double centerX, double centerY, double width, double height)
        {
            string value = (text ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return;
            }

            double charBoxHeight = Math.Min(0.22, height / value.Length);
            double totalHeight = charBoxHeight * value.Length;
            double top = centerY + totalHeight / 2.0;

            for (int i = 0; i < value.Length; i++)
            {
                double charY = top - charBoxHeight * (i + 0.5);
                Visio.Shape charShape = page.DrawRectangle(
                    centerX - width / 2.0,
                    charY - charBoxHeight / 2.0,
                    centerX + width / 2.0,
                    charY + charBoxHeight / 2.0);
                charShape.Text = value[i].ToString();
                charShape.CellsU["LinePattern"].Formula = "0";
                charShape.CellsU["FillPattern"].Formula = "0";
                charShape.CellsU["Char.Color"].Formula = "RGB(0, 0, 0)";
                charShape.CellsU["Char.Size"].FormulaU = $"{currentFontSizePt.ToString(CultureInfo.InvariantCulture)}pt";
                charShape.CellsU["Char.Font"].FormulaU = $"\"{currentFontName}\"";
                charShape.CellsU["Para.HorzAlign"].Formula = "1";
                charShape.CellsU["VerticalAlign"].Formula = "1";
                TrySetFormula(charShape, "TxtPinX", "Width*0.5");
                TrySetFormula(charShape, "TxtPinY", "Height*0.5");
                TrySetFormula(charShape, "TxtWidth", "Width");
                TrySetFormula(charShape, "TxtHeight", "Height");
                TrySetFormula(charShape, "LeftMargin", "0");
                TrySetFormula(charShape, "RightMargin", "0");
                TrySetFormula(charShape, "TopMargin", "0");
                TrySetFormula(charShape, "BottomMargin", "0");
            }
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

        private void DrawConnector(Visio.Page page, TreeNode parent, TreeNode child, double defaultWidth, double defaultHeight, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int parentLevel, int childLevel)
        {
            double parentWidth;
            double parentHeight;
            double childWidth;
            double childHeight;
            GetShapeSize(parentLevel, defaultWidth, defaultHeight, optionsByLevel, out parentWidth, out parentHeight);
            GetShapeSize(childLevel, defaultWidth, defaultHeight, optionsByLevel, out childWidth, out childHeight);

            double beginX;
            double beginY;
            double endX;
            double endY;

            if (isLeftToRight)
            {
                beginX = parent.X + parentWidth / 2.0;
                beginY = parent.Y;
                endX = child.X - childWidth / 2.0;
                endY = child.Y;
            }
            else
            {
                beginX = parent.X;
                beginY = parent.Y - parentHeight / 2.0;
                endX = child.X;
                endY = child.Y + childHeight / 2.0;
            }

            if (isLeftToRight)
            {
                double midX = (beginX + endX) / 2.0;
                DrawAcademicLine(page, beginX, beginY, midX, beginY);
                DrawAcademicLine(page, midX, beginY, midX, endY);
                Visio.Shape lastSegment = DrawAcademicLine(page, midX, endY, endX, endY);
                lastSegment.CellsU["EndArrow"].Formula = chkArrow.Checked ? "4" : "0";
            }
            else
            {
                double midY = beginY - Math.Max(0.12, (beginY - endY) * 0.45);
                DrawAcademicLine(page, beginX, beginY, beginX, midY);
                DrawAcademicLine(page, beginX, midY, endX, midY);
                Visio.Shape lastSegment = DrawAcademicLine(page, endX, midY, endX, endY);
                lastSegment.CellsU["EndArrow"].Formula = chkArrow.Checked ? "4" : "0";
            }
        }

        private void DrawTitleConnector(Visio.Page page, double beginX, double beginY, double endX, double endY)
        {
            if (Math.Abs(beginX - endX) < 0.001)
            {
                DrawAcademicLine(page, beginX, beginY, endX, endY);
                return;
            }

            double midY = (beginY + endY) / 2.0;
            DrawAcademicLine(page, beginX, beginY, beginX, midY);
            DrawAcademicLine(page, beginX, midY, endX, midY);
            DrawAcademicLine(page, endX, midY, endX, endY);
        }

        private Visio.Shape DrawAcademicLine(Visio.Page page, double beginX, double beginY, double endX, double endY)
        {
            Visio.Shape connector = page.DrawLine(beginX, beginY, endX, endY);
            connector.CellsU["LineColor"].Formula = "RGB(0, 0, 0)";
            connector.CellsU["LineWeight"].Formula = $"{currentLineWidthPt.ToString(CultureInfo.InvariantCulture)}pt";
            connector.CellsU["BeginArrow"].Formula = "0";
            connector.CellsU["EndArrow"].Formula = "0";
            return connector;
        }

        private void GetShapeSize(int level, double defaultWidth, double defaultHeight, Dictionary<int, LevelOption> optionsByLevel, out double width, out double height)
        {
            LevelOption option;
            if (optionsByLevel != null && optionsByLevel.TryGetValue(level, out option))
            {
                width = option.WidthMm / 25.4;
                height = option.HeightMm / 25.4;
                return;
            }

            width = defaultWidth;
            height = defaultHeight;
        }

        #endregion
    }
}
