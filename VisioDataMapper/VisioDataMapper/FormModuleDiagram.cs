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
        private Label lblDefaultHorizontalHeight;
        private TextBox txtDefaultHorizontalHeight;
        private Label lblDefaultVerticalSpacing;
        private TextBox txtDefaultVerticalSpacing;

        private Label lblLevelCount;
        private Label lblSelectLevel;
        private ComboBox cmbSelectLevel;
        private bool isUpdatingLevels = false;
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
        private double currentDefaultHorizontalShapeHeightChars = DefaultHorizontalShapeHeightChars;
        private double currentDefaultVerticalShapeSpacingPx = DefaultVerticalShapeSpacingPx;
        private const double RootTopMarginInch = 1.45;
        private const double DefaultRootTopExtraGapInch = 0.35;
        private const double DefaultFifthSizeFontPt = 10.5;
        private const double DefaultDpi = 96.0;
        private const double DefaultHorizontalPaddingChars = 2.0;
        private const double DefaultVerticalShapeWidthChars = 1.5;
        // 横着排列时表格“高(字)”的默认值。
        private const double DefaultHorizontalShapeHeightChars = 2.0;
        // 竖着排列时表格“形状间隔(px)”的默认值。
        private const double DefaultVerticalShapeSpacingPx = 10.0;
        private const double DefaultCharWidthMm = 3.7;
        private const double DefaultCharHeightMm = 5.8;
        
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

            lblDefaultHorizontalHeight = new Label { Text = "横排默认高(字):", Location = new Point(785, 570), Size = new Size(110, 22) };
            txtDefaultHorizontalHeight = new TextBox { Text = DefaultHorizontalShapeHeightChars.ToString("0.#", CultureInfo.InvariantCulture), Location = new Point(900, 566), Size = new Size(45, 25) };
            txtDefaultHorizontalHeight.TextChanged += DefaultLevelSetting_TextChanged;

            lblDefaultVerticalSpacing = new Label { Text = "竖排间隔(px):", Location = new Point(955, 570), Size = new Size(95, 22) };
            txtDefaultVerticalSpacing = new TextBox { Text = DefaultVerticalShapeSpacingPx.ToString("0.#", CultureInfo.InvariantCulture), Location = new Point(1055, 566), Size = new Size(45, 25) };
            txtDefaultVerticalSpacing.TextChanged += DefaultLevelSetting_TextChanged;

            lblFontName = new Label { Text = "字体:", Location = new Point(785, 605), Size = new Size(40, 22) };
            cmbFontName = new ComboBox { Location = new Point(830, 601), Size = new Size(110, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontName.Items.AddRange(new string[] { "宋体", "黑体", "仿宋", "楷体", "微软雅黑" });
            cmbFontName.SelectedIndex = 0;

            lblFontSize = new Label { Text = "字号:", Location = new Point(955, 605), Size = new Size(40, 22) };
            txtFontSize = new TextBox { Text = "10.5", Location = new Point(1000, 601), Size = new Size(55, 25) };

            lblLevelCount = new Label { Text = "共 0 个层级", Location = new Point(15, 605), Size = new Size(100, 22) };
            lblSelectLevel = new Label { Text = "生成层级选择:", Location = new Point(125, 605), Size = new Size(100, 22) };
            cmbSelectLevel = new ComboBox { Location = new Point(230, 601), Size = new Size(80, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbSelectLevel.SelectedIndexChanged += cmbSelectLevel_SelectedIndexChanged;

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
            dgvLevels.CurrentCellDirtyStateChanged += dgvLevels_CurrentCellDirtyStateChanged;
            dgvLevels.CellValueChanged += dgvLevels_CellValueChanged;

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
            this.Controls.Add(lblDefaultHorizontalHeight);
            this.Controls.Add(txtDefaultHorizontalHeight);
            this.Controls.Add(lblDefaultVerticalSpacing);
            this.Controls.Add(txtDefaultVerticalSpacing);

            this.Controls.Add(lblFontName);
            this.Controls.Add(cmbFontName);
            this.Controls.Add(lblFontSize);
            this.Controls.Add(txtFontSize);
            this.Controls.Add(lblLevelCount);
            this.Controls.Add(lblSelectLevel);
            this.Controls.Add(cmbSelectLevel);
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
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "RootTopExtraGapInch", HeaderText = "RootTopExtraGapInch", FillWeight = 95 });
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "ShapeSpacingPx", HeaderText = "形状间隔(px)", FillWeight = 85 });

            var directionColumn = new DataGridViewComboBoxColumn { Name = "Direction", HeaderText = "形状排列", FillWeight = 90 };
            directionColumn.Items.AddRange("横着排列", "竖着排列");
            dgvLevels.Columns.Add(directionColumn);

            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "Width", HeaderText = "宽(字)", FillWeight = 70 });
            dgvLevels.Columns.Add(new DataGridViewTextBoxColumn { Name = "Height", HeaderText = "高(字)", FillWeight = 70 });
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
            lblDefaultHorizontalHeight.Top = rowTop + 4;
            txtDefaultHorizontalHeight.Top = rowTop;
            lblDefaultVerticalSpacing.Top = rowTop + 4;
            txtDefaultVerticalSpacing.Top = rowTop;

            int secondRowTop = rowTop + 35;
            lblFontName.Top = secondRowTop + 4;
            cmbFontName.Top = secondRowTop;
            lblFontSize.Top = secondRowTop + 4;
            txtFontSize.Top = secondRowTop;

            int rightEdge = margin + width;
            txtDefaultVerticalSpacing.Left = rightEdge - txtDefaultVerticalSpacing.Width;
            lblDefaultVerticalSpacing.Left = txtDefaultVerticalSpacing.Left - lblDefaultVerticalSpacing.Width - 5;
            txtDefaultHorizontalHeight.Left = lblDefaultVerticalSpacing.Left - txtDefaultHorizontalHeight.Width - 15;
            lblDefaultHorizontalHeight.Left = txtDefaultHorizontalHeight.Left - lblDefaultHorizontalHeight.Width - 5;

            txtFontSize.Left = rightEdge - txtFontSize.Width;
            lblFontSize.Left = txtFontSize.Left - lblFontSize.Width - 5;
            cmbFontName.Left = lblFontSize.Left - cmbFontName.Width - 15;
            lblFontName.Left = cmbFontName.Left - lblFontName.Width - 5;

            lblLevelCount.Top = secondRowTop + 4;
            lblLevelCount.Width = 100;
            lblSelectLevel.Top = secondRowTop + 4;
            lblSelectLevel.Left = lblLevelCount.Right + 10;
            cmbSelectLevel.Top = secondRowTop;
            cmbSelectLevel.Left = lblSelectLevel.Right + 5;
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

        private void cmbSelectLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isUpdatingLevels) return;
            RefreshLevelOptions();
        }

        private void DefaultLevelSetting_TextChanged(object sender, EventArgs e)
        {
            if (isUpdatingLevels) return;
            RefreshLevelOptions();
        }

        private void dgvLevels_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgvLevels.IsCurrentCellDirty)
            {
                dgvLevels.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgvLevels_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (isUpdatingLevels || e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            if (dgvLevels.Columns[e.ColumnIndex].Name != "Direction")
            {
                return;
            }

            DataGridViewRow row = dgvLevels.Rows[e.RowIndex];
            string direction = Convert.ToString(row.Cells["Direction"].Value);
            string maxText = Convert.ToString(row.Cells["MaxText"].Value);
            bool isVertical = direction == "竖着排列";
            double widthChars;
            double heightChars;

            if (isVertical)
            {
                if (!TryParsePositiveNumber(Convert.ToString(row.Cells["Width"].Value), out widthChars) ||
                    IsLegacyVerticalWidthValue(widthChars, maxText))
                {
                    row.Cells["Width"].Value = GetDefaultVerticalShapeWidthChars().ToString("0.#", CultureInfo.InvariantCulture);
                }

                if (!TryParsePositiveNumber(Convert.ToString(row.Cells["Height"].Value), out heightChars) ||
                    IsLegacyHorizontalHeightValue(heightChars))
                {
                    row.Cells["Height"].Value = GetVerticalTextHeightChars(maxText).ToString("0.#", CultureInfo.InvariantCulture);
                }

                double spacingPx;
                if (!TryParsePositiveNumber(Convert.ToString(row.Cells["ShapeSpacingPx"].Value), out spacingPx))
                {
                    row.Cells["ShapeSpacingPx"].Value = currentDefaultVerticalShapeSpacingPx.ToString("0.#", CultureInfo.InvariantCulture);
                }
            }
            else
            {
                double horizontalWidthChars = GetHorizontalTextWidthChars(maxText);
                if (!TryParsePositiveNumber(Convert.ToString(row.Cells["Width"].Value), out widthChars) ||
                    widthChars < horizontalWidthChars ||
                    IsLegacyHorizontalWidthValue(widthChars, maxText))
                {
                    row.Cells["Width"].Value = horizontalWidthChars.ToString("0.#", CultureInfo.InvariantCulture);
                }

                if (!TryParsePositiveNumber(Convert.ToString(row.Cells["Height"].Value), out heightChars) ||
                    IsLegacyHorizontalHeightValue(heightChars))
                {
                    row.Cells["Height"].Value = currentDefaultHorizontalShapeHeightChars.ToString("0.#", CultureInfo.InvariantCulture);
                }
            }
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

            double parsedFontSize;
            if (txtFontSize != null && TryParsePositiveNumber(txtFontSize.Text, out parsedFontSize))
            {
                currentFontSizePt = parsedFontSize;
            }
            double previousDefaultHorizontalHeightChars = currentDefaultHorizontalShapeHeightChars;
            double previousDefaultVerticalSpacingPx = currentDefaultVerticalShapeSpacingPx;
            UpdateCurrentDefaultLevelSettings();

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



            int totalLevels = stats.Count;
            lblLevelCount.Text = $"共 {totalLevels} 个层级";

            isUpdatingLevels = true;
            try
            {
                if (cmbSelectLevel.Items.Count != totalLevels)
                {
                    cmbSelectLevel.Items.Clear();
                    for (int i = 1; i <= totalLevels; i++)
                    {
                        cmbSelectLevel.Items.Add(i);
                    }
                    if (totalLevels > 0)
                    {
                        cmbSelectLevel.SelectedIndex = totalLevels - 1;
                    }
                }
            }
            finally
            {
                isUpdatingLevels = false;
            }

            int selectedMaxLevel = totalLevels;
            if (cmbSelectLevel.SelectedItem != null)
            {
                selectedMaxLevel = Convert.ToInt32(cmbSelectLevel.SelectedItem);
            }

            dgvLevels.Rows.Clear();
            levelOptions.Clear();

            for (int level = 1; level <= selectedMaxLevel; level++)
            {
                if (!stats.ContainsKey(level)) continue;
                LevelStats stat = stats[level];
                bool isLastLevel = level == selectedMaxLevel;
                LevelOption existingOption;
                bool hasExistingOption = existingOptions.TryGetValue(level, out existingOption);
                // 末层强制竖排，宽高重新自适应；其他层从已有选项读取，用户可手动调整
                bool isVertical;
                double widthChars;
                double heightChars;
                if (isLastLevel)
                {
                    isVertical = true;
                    widthChars = hasExistingOption ? existingOption.WidthChars : GetDefaultVerticalShapeWidthChars();
                    heightChars = hasExistingOption ? existingOption.HeightChars : GetVerticalTextHeightChars(stat.LongestText);
                }
                else
                {
                    isVertical = hasExistingOption ? existingOption.IsVertical : false;
                    widthChars = hasExistingOption ? existingOption.WidthChars : (isVertical ? GetDefaultVerticalShapeWidthChars() : GetHorizontalTextWidthChars(stat.LongestText));
                    heightChars = hasExistingOption ? existingOption.HeightChars : (isVertical ? GetVerticalTextHeightChars(stat.LongestText) : currentDefaultHorizontalShapeHeightChars);
                }
                if (isVertical)
                {
                    widthChars = NormalizeVerticalShapeWidthChars(widthChars, stat.LongestText);
                    heightChars = NormalizeVerticalShapeHeightChars(heightChars, stat.LongestText);
                }
                else
                {
                    widthChars = NormalizeHorizontalShapeWidthChars(widthChars, stat.LongestText);
                    heightChars = NormalizeHorizontalShapeHeightChars(heightChars, previousDefaultHorizontalHeightChars);
                }
                double rootTopExtraGapInch = hasExistingOption ? existingOption.RootTopExtraGapInch : DefaultRootTopExtraGapInch;
                if (level != 1)
                {
                    rootTopExtraGapInch = 0;
                }
                double shapeSpacingPx = hasExistingOption ? existingOption.ShapeSpacingPx : (isVertical ? currentDefaultVerticalShapeSpacingPx : 0);
                if (isVertical && IsDefaultLikeVerticalShapeSpacingValue(shapeSpacingPx, previousDefaultVerticalSpacingPx))
                {
                    shapeSpacingPx = currentDefaultVerticalShapeSpacingPx;
                }

                var option = new LevelOption
                {
                    Level = level,
                    NodeCount = stat.NodeCount,
                    MaxText = stat.LongestText,
                    IsVertical = isVertical,
                    WidthChars = widthChars,
                    HeightChars = heightChars,
                    RootTopExtraGapInch = rootTopExtraGapInch,
                    ShapeSpacingPx = shapeSpacingPx
                };
                levelOptions.Add(option);

                dgvLevels.Rows.Add(
                    level,
                    stat.NodeCount,
                    stat.LongestText,
                    level == 1 ? rootTopExtraGapInch.ToString("0.###", CultureInfo.InvariantCulture) : string.Empty,
                    isVertical ? shapeSpacingPx.ToString("0.#", CultureInfo.InvariantCulture) : string.Empty,
                    isVertical ? "竖着排列" : "横着排列",
                    widthChars.ToString("0.#", CultureInfo.InvariantCulture),
                    heightChars.ToString("0.#", CultureInfo.InvariantCulture));
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

        private double GetHorizontalTextWidthChars(string text)
        {
            return Math.Max(4.0, GetChineseCharCount(text) + DefaultHorizontalPaddingChars);
        }

        private double GetVerticalTextHeightChars(string text)
        {
            return GetChineseCharCount(text);
        }

        private double GetDefaultVerticalShapeWidthChars()
        {
            return DefaultVerticalShapeWidthChars;
        }

        private double NormalizeVerticalShapeWidthChars(double widthChars, string text)
        {
            if (IsLegacyVerticalWidthValue(widthChars, text))
            {
                return GetDefaultVerticalShapeWidthChars();
            }

            return widthChars;
        }

        private double NormalizeVerticalShapeHeightChars(double heightChars, string text)
        {
            if (IsLegacyVerticalHeightValue(heightChars, text))
            {
                return GetVerticalTextHeightChars(text);
            }

            return heightChars;
        }

        private double NormalizeHorizontalShapeWidthChars(double widthChars, string text)
        {
            double textWidthChars = GetHorizontalTextWidthChars(text);
            if (widthChars < textWidthChars || IsLegacyHorizontalWidthValue(widthChars, text))
            {
                return textWidthChars;
            }

            return widthChars;
        }

        private double NormalizeHorizontalShapeHeightChars(double heightChars)
        {
            return NormalizeHorizontalShapeHeightChars(heightChars, DefaultHorizontalShapeHeightChars);
        }

        private double NormalizeHorizontalShapeHeightChars(double heightChars, double previousDefaultHeightChars)
        {
            return IsDefaultLikeHorizontalHeightValue(heightChars, previousDefaultHeightChars) ? currentDefaultHorizontalShapeHeightChars : heightChars;
        }

        private bool IsLegacyVerticalWidthValue(double widthValue, string text)
        {
            return Math.Abs(widthValue - 14.0) < 0.05 ||
                Math.Abs(widthValue - 5.5) < 0.15 ||
                Math.Abs(widthValue - 4.2) < 0.15 ||
                Math.Abs(widthValue - EstimateLegacyHorizontalWidthMm(text)) < 0.05;
        }

        private bool IsLegacyHorizontalWidthValue(double widthValue, string text)
        {
            return Math.Abs(widthValue - EstimateLegacyHorizontalWidthMm(text)) < 0.05;
        }

        private bool IsLegacyHorizontalHeightValue(double heightValue)
        {
            return Math.Abs(heightValue - 10.0) < 0.05;
        }

        private bool IsDefaultLikeHorizontalHeightValue(double heightValue, double previousDefaultHeightChars)
        {
            return IsLegacyHorizontalHeightValue(heightValue) ||
                Math.Abs(heightValue - previousDefaultHeightChars) < 0.05 ||
                Math.Abs(heightValue - DefaultHorizontalShapeHeightChars) < 0.05;
        }

        private bool IsLegacyVerticalHeightValue(double heightValue, string text)
        {
            return Math.Abs(heightValue - (GetChineseCharCount(text) + 2.0)) < 0.05;
        }

        private bool IsDefaultLikeVerticalShapeSpacingValue(double spacingPx, double previousDefaultSpacingPx)
        {
            return Math.Abs(spacingPx - 1.0) < 0.05 ||
                Math.Abs(spacingPx - 3.0) < 0.05 ||
                Math.Abs(spacingPx - previousDefaultSpacingPx) < 0.05 ||
                Math.Abs(spacingPx - DefaultVerticalShapeSpacingPx) < 0.05;
        }

        private void UpdateCurrentDefaultLevelSettings()
        {
            double parsedValue;
            if (txtDefaultHorizontalHeight != null && TryParsePositiveNumber(txtDefaultHorizontalHeight.Text, out parsedValue))
            {
                currentDefaultHorizontalShapeHeightChars = parsedValue;
            }

            if (txtDefaultVerticalSpacing != null && TryParsePositiveNumber(txtDefaultVerticalSpacing.Text, out parsedValue))
            {
                currentDefaultVerticalShapeSpacingPx = parsedValue;
            }
        }

        private double EstimateLegacyHorizontalWidthMm(string text)
        {
            return Math.Max(28, GetChineseCharCount(text) * 5.2 + 12);
        }

        private double WidthCharsToMm(double widthChars)
        {
            return Math.Max(DefaultCharWidthMm, widthChars * DefaultCharWidthMm);
        }

        private double HeightCharsToMm(double heightChars)
        {
            return Math.Max(DefaultCharHeightMm, heightChars * DefaultCharHeightMm);
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
                double widthChars;
                double heightChars;
                if (!int.TryParse(Convert.ToString(row.Cells["Level"].Value), out level) ||
                    !TryParsePositiveNumber(Convert.ToString(row.Cells["Width"].Value), out widthChars) ||
                    !TryParsePositiveNumber(Convert.ToString(row.Cells["Height"].Value), out heightChars))
                {
                    continue;
                }

                double rootTopExtraGapInch = 0;
                if (level == 1)
                {
                    double parsedRootTopExtraGapInch;
                    if (!TryParsePositiveNumber(Convert.ToString(row.Cells["RootTopExtraGapInch"].Value), out parsedRootTopExtraGapInch))
                    {
                        parsedRootTopExtraGapInch = DefaultRootTopExtraGapInch;
                    }

                    rootTopExtraGapInch = parsedRootTopExtraGapInch;
                }
                double shapeSpacingPx = 0;
                double parsedShapeSpacingPx;
                if (TryParsePositiveNumber(Convert.ToString(row.Cells["ShapeSpacingPx"].Value), out parsedShapeSpacingPx))
                {
                    shapeSpacingPx = parsedShapeSpacingPx;
                }

                result[level] = new LevelOption
                {
                    Level = level,
                    NodeCount = row.Cells["NodeCount"].Value == null ? 0 : Convert.ToInt32(row.Cells["NodeCount"].Value),
                    MaxText = Convert.ToString(row.Cells["MaxText"].Value),
                    IsVertical = Convert.ToString(row.Cells["Direction"].Value) == "竖着排列",
                    WidthChars = widthChars,
                    HeightChars = heightChars,
                    RootTopExtraGapInch = rootTopExtraGapInch,
                    ShapeSpacingPx = shapeSpacingPx
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

                // Determine the max rendering level from the combo selection
                int selectedMaxLevel = levelOptions.Count > 0 ? levelOptions[levelOptions.Count - 1].Level : int.MaxValue;
                if (cmbSelectLevel.SelectedItem != null)
                {
                    selectedMaxLevel = Convert.ToInt32(cmbSelectLevel.SelectedItem);
                }

                AppendLog("运行层级布局引擎算法...");
                // Step 1: Calculate subtree sizes
                CalculateSubtreeSizes(root, defaultShapeWidth, shapeHeightVal, horSpacingVal, verSpacingVal, optionsByLevel, isLeftToRight, 1, selectedMaxLevel);

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
                    root.Y = pageHeight - RootTopMarginInch; // leave room for the root node itself
                }

                // Step 2: Compute absolute coordinates
                AssignCoordinates(root, defaultShapeWidth, shapeHeightVal, horSpacingVal, verSpacingVal, optionsByLevel, isLeftToRight, 1, selectedMaxLevel);

                AppendLog("在 Visio 中绘制形状和连线...");
                // Draw nodes and connect them
                DrawTree(activePage, root, defaultShapeWidth, shapeHeightVal, optionsByLevel, isLeftToRight, 1, selectedMaxLevel);

                // The module diagram title is only used as form metadata; it is not rendered.

                AppendLog("生成绘图完成！");
                this.Close();
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
            UpdateCurrentDefaultLevelSettings();

            foreach (DataGridViewRow row in dgvLevels.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                int level = Convert.ToInt32(row.Cells["Level"].Value);
                string direction = Convert.ToString(row.Cells["Direction"].Value);
                string maxText = Convert.ToString(row.Cells["MaxText"].Value);
                double widthChars = ParsePositiveNumber(Convert.ToString(row.Cells["Width"].Value), $"第{level}层宽度字符数");
                double heightChars = ParsePositiveNumber(Convert.ToString(row.Cells["Height"].Value), $"第{level}层高度字符数");
                bool isVertical = direction == "竖着排列";
                if (isVertical)
                {
                    widthChars = NormalizeVerticalShapeWidthChars(widthChars, maxText);
                }
                else
                {
                    widthChars = NormalizeHorizontalShapeWidthChars(widthChars, maxText);
                    heightChars = NormalizeHorizontalShapeHeightChars(heightChars);
                }
                double rootTopExtraGapInch = 0;
                if (level == 1)
                {
                    rootTopExtraGapInch = ParsePositiveNumber(Convert.ToString(row.Cells["RootTopExtraGapInch"].Value), "RootTopExtraGapInch");
                }
                double shapeSpacingPx = 0;
                double parsedShapeSpacingPx;
                if (TryParsePositiveNumber(Convert.ToString(row.Cells["ShapeSpacingPx"].Value), out parsedShapeSpacingPx))
                {
                    shapeSpacingPx = parsedShapeSpacingPx;
                }
                else if (isVertical)
                {
                    shapeSpacingPx = currentDefaultVerticalShapeSpacingPx;
                }

                result[level] = new LevelOption
                {
                    Level = level,
                    NodeCount = Convert.ToInt32(row.Cells["NodeCount"].Value),
                    MaxText = maxText,
                    IsVertical = isVertical,
                    WidthChars = widthChars,
                    HeightChars = heightChars,
                    RootTopExtraGapInch = rootTopExtraGapInch,
                    ShapeSpacingPx = shapeSpacingPx
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
            public Visio.Shape RouteAnchorShape { get; set; }
        }

        private class LevelOption
        {
            public int Level { get; set; }
            public int NodeCount { get; set; }
            public string MaxText { get; set; }
            public bool IsVertical { get; set; }
            public double WidthChars { get; set; }
            public double HeightChars { get; set; }
            public double RootTopExtraGapInch { get; set; }
            public double ShapeSpacingPx { get; set; }
        }

        private class LevelStats
        {
            public int NodeCount { get; set; }
            public string LongestText { get; set; } = string.Empty;
        }

        private TreeNode ParseTree(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            try
            {
                var serializer = new JavaScriptSerializer();
                var data = serializer.Deserialize<System.Collections.Generic.Dictionary<string, object>>(text);
                if (data == null) return null;

                string sysName = data.ContainsKey("sys_name") && data["sys_name"] != null ? data["sys_name"].ToString() : "系统";

                // Set diagram title to sys_name if user hasn't modified the title box
                if (txtTitle.Text == "功能模块图" && data.ContainsKey("sys_name") && data["sys_name"] != null && !string.IsNullOrWhiteSpace(data["sys_name"].ToString()))
                {
                    txtTitle.Text = data["sys_name"].ToString();
                }

                TreeNode root = new TreeNode { Text = sysName };

                if (data.ContainsKey("platform") && data["platform"] is System.Collections.ArrayList platList)
                {
                    foreach (var platObj in platList)
                    {
                        if (platObj is System.Collections.Generic.Dictionary<string, object> plat)
                        {
                            string pName = plat.ContainsKey("name") && plat["name"] != null ? plat["name"].ToString() : "";
                            TreeNode platNode = new TreeNode { Text = pName, Parent = root };
                            root.Children.Add(platNode);

                            if (plat.ContainsKey("module") && plat["module"] is System.Collections.ArrayList modList)
                            {
                                foreach (var modObj in modList)
                                {
                                    if (modObj is string modName)
                                    {
                                        TreeNode modNode = new TreeNode { Text = modName, Parent = platNode };
                                        platNode.Children.Add(modNode);
                                    }
                                    else if (modObj is System.Collections.Generic.Dictionary<string, object> modDict)
                                    {
                                        string mName = modDict.ContainsKey("name") && modDict["name"] != null ? modDict["name"].ToString() : "";
                                        TreeNode modNode = new TreeNode { Text = mName, Parent = platNode };
                                        platNode.Children.Add(modNode);

                                        if (modDict.ContainsKey("sub") && modDict["sub"] is System.Collections.ArrayList subList)
                                        {
                                            foreach (var subObj in subList)
                                            {
                                                TreeNode subNode = new TreeNode { Text = subObj.ToString(), Parent = modNode };
                                                modNode.Children.Add(subNode);
                                            }
                                        }
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

        private void CalculateSubtreeSizes(TreeNode node, double defaultWidth, double defaultHeight, double horSpacing, double verSpacing, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level, int selectedMaxLevel = int.MaxValue)
        {
            double width;
            double height;
            GetShapeSize(level, defaultWidth, defaultHeight, optionsByLevel, out width, out height);
            double selfSize = isLeftToRight ? height : width;

            if (node.Children.Count == 0 || level >= selectedMaxLevel)
            {
                node.SubtreeSize = selfSize;
                return;
            }

            double childrenSize = 0;
            foreach (var child in node.Children)
            {
                CalculateSubtreeSizes(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1, selectedMaxLevel);
                childrenSize += child.SubtreeSize;
            }

            double spacing = GetSiblingSpacing(node, horSpacing, verSpacing, optionsByLevel, isLeftToRight);
            childrenSize += (node.Children.Count - 1) * spacing;

            node.SubtreeSize = Math.Max(selfSize, childrenSize);
        }

        private double GetSiblingSpacing(TreeNode parent, double horSpacing, double verSpacing, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight)
        {
            double defaultSpacing = isLeftToRight ? verSpacing : horSpacing;
            if (parent == null || parent.Children.Count == 0)
            {
                return defaultSpacing;
            }

            int childLevel = GetNodeDepth(parent.Children[0]);
            LevelOption option;
            if (optionsByLevel != null &&
                optionsByLevel.TryGetValue(childLevel, out option) &&
                option.IsVertical &&
                option.ShapeSpacingPx > 0)
            {
                return option.ShapeSpacingPx / DefaultDpi;
            }

            return defaultSpacing;
        }

        private int GetNodeDepth(TreeNode node)
        {
            int depth = 0;
            TreeNode current = node;
            while (current != null)
            {
                depth++;
                current = current.Parent;
            }

            return depth;
        }

        private void AssignCoordinates(TreeNode node, double defaultWidth, double defaultHeight, double horSpacing, double verSpacing, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level, int selectedMaxLevel = int.MaxValue)
        {
            if (node.Children.Count == 0 || level >= selectedMaxLevel) return;

            double spacing = GetSiblingSpacing(node, horSpacing, verSpacing, optionsByLevel, isLeftToRight);
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
                    currentY -= child.SubtreeSize + spacing;

                    AssignCoordinates(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1, selectedMaxLevel);
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
                    currentX += child.SubtreeSize + spacing;

                    AssignCoordinates(child, defaultWidth, defaultHeight, horSpacing, verSpacing, optionsByLevel, isLeftToRight, level + 1, selectedMaxLevel);
                }
            }
        }

        private void DrawTree(Visio.Page page, TreeNode node, double defaultWidth, double defaultHeight, Dictionary<int, LevelOption> optionsByLevel, bool isLeftToRight, int level, int selectedMaxLevel = int.MaxValue)
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
            node.VisioShape.Text = isVerticalText ? FormatTopDownText(node.Text) : node.Text;
            ApplyAcademicShapeStyle(node.VisioShape);
            if (isVerticalText)
            {
                ApplyTopDownTextBlockStyle(node.VisioShape, node.Text);
            }

            // Connect parent to child
            if (node.Parent != null && node.Parent.VisioShape != null)
            {
                DrawConnector(page, node.Parent, node, defaultWidth, defaultHeight, optionsByLevel, isLeftToRight, level - 1, level);
            }

            // Draw children
            if (level < selectedMaxLevel)
            {
                foreach (var child in node.Children)
                {
                    DrawTree(page, child, defaultWidth, defaultHeight, optionsByLevel, isLeftToRight, level + 1, selectedMaxLevel);
                }
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

        private string FormatTopDownText(string text)
        {
            string value = (text ?? string.Empty).Trim();
            if (value.Length == 0)
            {
                return string.Empty;
            }

            return string.Join("\n", value.ToCharArray());
        }

        private void ApplyTopDownTextBlockStyle(Visio.Shape shape, string text)
        {
            double textWidthInch = Math.Max(0.12, currentFontSizePt / 72.0 * 1.1);
            string textWidthFormula = $"{textWidthInch.ToString(CultureInfo.InvariantCulture)} in";

            TrySetFormula(shape, "TxtWidth", textWidthFormula);
            TrySetFormula(shape, "TxtHeight", "Height");
            TrySetFormula(shape, "TxtPinX", "Width*0.5");
            TrySetFormula(shape, "TxtPinY", "Height*0.5");
            TrySetFormula(shape, "TxtLocPinX", "TxtWidth*0.5");
            TrySetFormula(shape, "TxtLocPinY", "TxtHeight*0.5");
            TrySetFormula(shape, "Para.HorzAlign", "1");
            TrySetFormula(shape, "VerticalAlign", "1");
            TrySetFormula(shape, "Char.Spacing", "0 pt");
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
            if (parent.VisioShape == null || child.VisioShape == null) return;

            if (parentLevel == 1 && !isLeftToRight)
            {
                DrawRootTopConnector(page, parent, child, defaultWidth, defaultHeight, optionsByLevel, parentLevel, childLevel);
                return;
            }

            DrawAutoConnector(page, parent.VisioShape, child.VisioShape, isLeftToRight);
        }

        private Visio.Shape DrawAutoConnector(Visio.Page page, Visio.Shape beginShape, Visio.Shape endShape, bool isLeftToRight)
        {
            object connectorTool = page.Application.ConnectorToolDataObject;
            Visio.Shape connector = page.Drop(connectorTool, 0, 0);

            // Set styles for the connector
            connector.CellsU["LineColor"].Formula = "RGB(0, 0, 0)";
            connector.CellsU["LineWeight"].Formula = $"{currentLineWidthPt.ToString(CultureInfo.InvariantCulture)}pt";
            connector.CellsU["BeginArrow"].Formula = "0";
            connector.CellsU["EndArrow"].Formula = chkArrow.Checked ? "4" : "0";

            // Force tree (OrgChart) routing style to merge trunks
            TrySetFormula(connector, "RouteStyle", isLeftToRight ? "6" : "5");
            TrySetFormula(connector, "ShapeRouteStyle", isLeftToRight ? "6" : "5");
            TrySetFormula(connector, "ConLineRouteExt", isLeftToRight ? "4" : "3");

            // Glue parent and child shape to the connector endpoints
            Visio.Cell beginCell = connector.CellsU["BeginX"];
            beginCell.GlueTo(beginShape.CellsU["PinX"]);

            Visio.Cell endCell = connector.CellsU["EndX"];
            endCell.GlueTo(endShape.CellsU["PinX"]);

            return connector;
        }

        private void DrawRootTopConnector(Visio.Page page, TreeNode parent, TreeNode child, double defaultWidth, double defaultHeight, Dictionary<int, LevelOption> optionsByLevel, int parentLevel, int childLevel)
        {
            double parentWidth;
            double parentHeight;
            double childWidth;
            double childHeight;
            GetShapeSize(parentLevel, defaultWidth, defaultHeight, optionsByLevel, out parentWidth, out parentHeight);
            GetShapeSize(childLevel, defaultWidth, defaultHeight, optionsByLevel, out childWidth, out childHeight);

            double beginX = parent.X;
            double beginY = parent.Y - parentHeight / 2.0;
            double branchY = beginY - GetRootTopExtraGapInch(optionsByLevel);

            if (parent.RouteAnchorShape == null)
            {
                parent.RouteAnchorShape = CreateHiddenRouteAnchor(page, beginX, branchY);
                DrawAutoConnector(page, parent.VisioShape, parent.RouteAnchorShape, false);
            }

            DrawAutoConnector(page, parent.RouteAnchorShape, child.VisioShape, false);
        }

        private Visio.Shape CreateHiddenRouteAnchor(Visio.Page page, double centerX, double centerY)
        {
            double size = 0.02;
            Visio.Shape anchor = page.DrawRectangle(centerX - size / 2.0, centerY - size / 2.0, centerX + size / 2.0, centerY + size / 2.0);
            anchor.Text = string.Empty;
            anchor.CellsU["FillPattern"].Formula = "0";
            anchor.CellsU["LinePattern"].Formula = "0";
            anchor.CellsU["Char.Color"].Formula = "RGB(255, 255, 255)";
            anchor.CellsU["Char.Size"].FormulaU = "1pt";
            TrySetFormula(anchor, "TxtPinX", "Width*0.5");
            TrySetFormula(anchor, "TxtPinY", "Height*0.5");
            TrySetFormula(anchor, "TxtWidth", "Width");
            TrySetFormula(anchor, "TxtHeight", "Height");
            TrySetFormula(anchor, "LeftMargin", "0");
            TrySetFormula(anchor, "RightMargin", "0");
            TrySetFormula(anchor, "TopMargin", "0");
            TrySetFormula(anchor, "BottomMargin", "0");
            return anchor;
        }

        private double GetRootTopExtraGapInch(Dictionary<int, LevelOption> optionsByLevel)
        {
            LevelOption option;
            if (optionsByLevel != null && optionsByLevel.TryGetValue(1, out option) && option.RootTopExtraGapInch > 0)
            {
                return option.RootTopExtraGapInch;
            }

            return DefaultRootTopExtraGapInch;
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
                width = WidthCharsToMm(option.WidthChars) / 25.4;
                height = HeightCharsToMm(option.HeightChars) / 25.4;
                return;
            }

            width = defaultWidth;
            height = defaultHeight;
        }

        #endregion
    }
}
