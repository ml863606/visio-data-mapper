using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormUseCaseDiagram : Form
    {
        private Button btnImportJson;
        private Button btnGenerate;
        private Button btnClose;
        private TextBox txtTitle;
        private Panel pnlTitleBorder;
        private Label lblTitle;
        private TextBox txtStatus;
        private Panel pnlStatusBorder;
        private TabControl tabActors;
        private Panel pnlTabContainer;
        private TabPage tabJson;
        private TabPage tabPlantUml;
        private JsonVisualEditorControl jsonEditor;
        private TextBox txtPlantUmlInput;
        private Panel pnlOptions;
        private ComboBox cmbActorShape;
        private ComboBox cmbFontName;
        private ComboBox cmbFontSize;
        private ComboBox cmbLayout;
        private TextBox txtHorizontalSpacing;
        private TextBox txtVerticalSpacing;
        private CheckBox chkEnglishRelationText;
        private CheckBox chkAssociationArrow;
        private CheckBox chkShowFunctionNodes;
        private CheckBox chkUseCaseOutline;
        private Label lblActorShape;
        private Label lblFontName;
        private Label lblFontSize;
        private Label lblLayout;
        private Label lblSpacing;
        private Label lblMillimeter;

        private string systemName = "系统用例图";
        private double currentLineWidth = 0.75;
        private string currentFontName = "宋体";
        private double currentFontSize = 10.5;
        private bool isInternalTextChange = false;

        public FormUseCaseDiagram()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "智能画图-用例图";
            Size = new Size(900, 850);
            MinimumSize = new Size(900, 780);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            Font defaultFont = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            Font = defaultFont;
            this.BackColor = SystemColors.Control;

            Label lblTip = new Label
            {
                Text = "请将UML用例JSON或PlantUML复制到剪贴板，点击导入文本后自动渲染。",
                Location = new Point(15, 20),
                AutoSize = true,
                ForeColor = SystemColors.ControlText,
                Font = defaultFont
            };

            btnImportJson = new Button 
            { 
                Text = "导入文本", 
                Location = new Point(15, 50), 
                Size = new Size(130, 36),
                Font = defaultFont
            };
            btnImportJson.Click += btnImportJson_Click;

            lblTitle = new Label 
            { 
                Text = "大标题:", 
                AutoSize = true,
                ForeColor = SystemColors.ControlText,
                Font = defaultFont
            };
            
            pnlTitleBorder = new Panel
            {
                BackColor = SystemColors.Control,
                Padding = new Padding(0),
                Location = new Point(535, 54),
                Size = new Size(340, 28)
            };
            txtTitle = new TextBox 
            { 
                Text = systemName,
                BorderStyle = BorderStyle.Fixed3D,
                Dock = DockStyle.Fill,
                Font = defaultFont,
                BackColor = Color.White
            };
            pnlTitleBorder.Controls.Add(txtTitle);

            pnlTabContainer = new Panel
            {
                Location = new Point(15, 100),
                Size = new Size(860, 455),
                BorderStyle = BorderStyle.None
            };

            tabActors = new TabControl
            {
                Location = new Point(-2, -2),
                Size = new Size(864, 459),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular),
                Padding = new Point(15, 6)
            };
            pnlTabContainer.Controls.Add(tabActors);

            tabJson = new TabPage("JSON数据");
            tabJson.BackColor = Color.White;
            tabJson.Text = "JSON数据";
            jsonEditor = new JsonVisualEditorControl();
            jsonEditor.SetJsonText(GetSampleJson(), true);
            jsonEditor.JsonTextChanged += jsonEditor_JsonTextChanged;
            tabJson.Controls.Add(jsonEditor);
            tabActors.TabPages.Add(tabJson);

            tabPlantUml = new TabPage("PlantUML");
            tabPlantUml.BackColor = Color.White;
            txtPlantUmlInput = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F, FontStyle.Regular),
                Text = GetSamplePlantUml(),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            txtPlantUmlInput.TextChanged += txtPlantUmlInput_TextChanged;
            tabPlantUml.Controls.Add(txtPlantUmlInput);
            tabActors.TabPages.Add(tabPlantUml);

            pnlStatusBorder = new Panel
            {
                BackColor = SystemColors.Control,
                Padding = new Padding(0),
                Location = new Point(15, 580),
                Size = new Size(860, 78)
            };
            txtStatus = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = SystemColors.Control,
                ForeColor = Color.DarkGray,
                BorderStyle = BorderStyle.Fixed3D,
                Dock = DockStyle.Fill,
                Text = $"{DateTime.Now:HH:mm:ss} 请在上方“JSON数据”文本框中粘贴JSON进行渲染。"
            };
            pnlStatusBorder.Controls.Add(txtStatus);

            lblActorShape = new Label { Text = "参与者形状:", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(3, 6, 3, 3) };
            cmbActorShape = new ComboBox { Size = new Size(110, 26), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular), Margin = new Padding(3, 3, 15, 3) };
            cmbActorShape.Items.AddRange(new string[] { "Draw.Io风格", "Visio自带" });
            cmbActorShape.SelectedIndex = 0;

            chkEnglishRelationText = new CheckBox { Text = "关系线条使用英文字符", AutoSize = true, Checked = true, ForeColor = SystemColors.ControlText, Margin = new Padding(5, 5, 15, 3) };
            chkShowFunctionNodes = new CheckBox { Text = "显示功能节点", AutoSize = true, Checked = true, ForeColor = SystemColors.ControlText, Margin = new Padding(5, 5, 15, 3) };
            chkUseCaseOutline = new CheckBox { Text = "模块/功能添加外边框", AutoSize = true, Checked = true, ForeColor = SystemColors.ControlText, Margin = new Padding(5, 5, 15, 3) };

            lblLayout = new Label { Text = "排版:", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(3, 6, 3, 3) };
            cmbLayout = new ComboBox { Size = new Size(130, 26), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular), Margin = new Padding(3, 3, 15, 3) };
            cmbLayout.Items.AddRange(new string[] { "参与者均在左侧", "参与者均在右侧", "参与者左右分布" });
            cmbLayout.SelectedIndex = 0;

            lblFontName = new Label { Text = "字体:", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(3, 6, 3, 3) };
            cmbFontName = CreateFontCombo(new Point(0, 0));
            cmbFontName.Margin = new Padding(3, 3, 15, 3);
            lblFontSize = new Label { Text = "字号:", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(3, 6, 3, 3) };
            cmbFontSize = CreateFontSizeCombo(new Point(0, 0));
            cmbFontSize.Margin = new Padding(3, 3, 15, 3);

            lblSpacing = new Label { Text = "横纵间距:", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(3, 6, 3, 3) };
            
            Panel pnlHorSpacingBorder = new Panel { BackColor = SystemColors.Control, Padding = new Padding(0), Size = new Size(40, 24), Margin = new Padding(3, 3, 5, 3) };
            txtHorizontalSpacing = new TextBox { Text = "15", BorderStyle = BorderStyle.Fixed3D, Dock = DockStyle.Fill, Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular), TextAlign = HorizontalAlignment.Center };
            pnlHorSpacingBorder.Controls.Add(txtHorizontalSpacing);

            Panel pnlVerSpacingBorder = new Panel { BackColor = SystemColors.Control, Padding = new Padding(0), Size = new Size(40, 24), Margin = new Padding(3, 3, 5, 3) };
            txtVerticalSpacing = new TextBox { Text = "8", BorderStyle = BorderStyle.Fixed3D, Dock = DockStyle.Fill, Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular), TextAlign = HorizontalAlignment.Center };
            pnlVerSpacingBorder.Controls.Add(txtVerticalSpacing);

            lblMillimeter = new Label { Text = "mm", AutoSize = true, ForeColor = SystemColors.ControlText, Margin = new Padding(0, 6, 15, 3) };
            chkAssociationArrow = new CheckBox { Text = "关联连线画箭头", AutoSize = true, Checked = false, ForeColor = SystemColors.ControlText, Margin = new Padding(5, 5, 15, 3) };

            btnGenerate = new Button 
            { 
                Text = "生成绘图", 
                Size = new Size(120, 36), 
                BackColor = Color.LightSkyBlue, 
                ForeColor = SystemColors.ControlText, 
                Font = new Font(defaultFont, FontStyle.Bold)
            };
            btnGenerate.Click += btnGenerate_Click;

            btnClose = new Button 
            { 
                Text = "关闭", 
                Size = new Size(100, 36),
                Font = defaultFont
            };
            btnClose.Click += btnClose_Click;

            pnlOptions = new Panel
            {
                BackColor = SystemColors.Control,
                BorderStyle = BorderStyle.None,
                Height = 135
            };

            FlowLayoutPanel row1 = new FlowLayoutPanel { Location = new Point(5, 5), Size = new Size(700, 36), Padding = new Padding(0) };
            row1.Controls.Add(lblActorShape);
            row1.Controls.Add(cmbActorShape);
            row1.Controls.Add(chkEnglishRelationText);
            row1.Controls.Add(chkShowFunctionNodes);
            row1.Controls.Add(chkUseCaseOutline);

            FlowLayoutPanel row2 = new FlowLayoutPanel { Location = new Point(5, 45), Size = new Size(700, 36), Padding = new Padding(0) };
            row2.Controls.Add(lblLayout);
            row2.Controls.Add(cmbLayout);
            row2.Controls.Add(lblFontName);
            row2.Controls.Add(cmbFontName);
            row2.Controls.Add(lblFontSize);
            row2.Controls.Add(cmbFontSize);

            FlowLayoutPanel row3 = new FlowLayoutPanel { Location = new Point(5, 85), Size = new Size(500, 40), Padding = new Padding(0) };
            row3.Controls.Add(lblSpacing);
            row3.Controls.Add(pnlHorSpacingBorder);
            row3.Controls.Add(pnlVerSpacingBorder);
            row3.Controls.Add(lblMillimeter);
            row3.Controls.Add(chkAssociationArrow);

            pnlOptions.Controls.Add(row1);
            pnlOptions.Controls.Add(row2);
            pnlOptions.Controls.Add(row3);
            pnlOptions.Controls.Add(btnGenerate);
            pnlOptions.Controls.Add(btnClose);

            Controls.Add(lblTip);
            Controls.Add(btnImportJson);
            Controls.Add(lblTitle);
            Controls.Add(pnlTitleBorder);
            Controls.Add(pnlTabContainer);
            Controls.Add(pnlStatusBorder);
            Controls.Add(pnlOptions);

            Resize += FormUseCaseDiagram_Resize;
            LayoutControls();
        }

        private DataGridView CreateActorGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersHeight = 30,
                RowTemplate = { Height = 28 },
                BackgroundColor = Color.White
            };
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Index", HeaderText = "序号", FillWeight = 42 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Module", HeaderText = "模块", FillWeight = 130 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Function", HeaderText = "功能", FillWeight = 170 });
            ApplyGridStyle(grid);
            return grid;
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

        private ComboBox CreateFontCombo(Point location)
        {
            var combo = new ComboBox { Location = location, Size = new Size(105, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(new string[] { "宋体", "黑体", "仿宋", "楷体", "微软雅黑" });
            combo.SelectedIndex = 0;
            return combo;
        }

        private ComboBox CreateFontSizeCombo(Point location)
        {
            var combo = new ComboBox { Location = location, Size = new Size(80, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(new string[] { "小三", "四号", "小四", "五号", "小五" });
            combo.SelectedIndex = 3;
            return combo;
        }

        private void LayoutControls()
        {
            if (tabActors == null || pnlOptions == null)
            {
                return;
            }

            int margin = 15;
            int width = Math.Max(820, ClientSize.Width - margin * 2);

            // Reposition top elements
            pnlTitleBorder.Left = Math.Max(535, margin + width - pnlTitleBorder.Width);
            lblTitle.Top = pnlTitleBorder.Top + 4;
            lblTitle.Left = pnlTitleBorder.Left - lblTitle.Width - 10;

            // Reposition panel at the bottom
            pnlOptions.Width = width;
            pnlOptions.Location = new Point(margin, ClientSize.Height - pnlOptions.Height - margin);

            // Reposition log box just above pnlOptions
            pnlStatusBorder.Width = width;
            pnlStatusBorder.Height = 65;
            pnlStatusBorder.Top = pnlOptions.Top - pnlStatusBorder.Height - margin;

            // Reposition tab control in the middle
            pnlTabContainer.Width = width;
            pnlTabContainer.Height = Math.Max(200, pnlStatusBorder.Top - pnlTabContainer.Top - margin);

            // Inside panel, reposition buttons dynamically to match right side
            btnClose.Top = 85;
            btnClose.Left = pnlOptions.Width - 15 - btnClose.Width;
            btnGenerate.Top = 85;
            btnGenerate.Left = btnClose.Left - 15 - btnGenerate.Width;
        }

        private void FormUseCaseDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private string GetSampleJson()
        {
            return
@"{
  ""system"": ""话费充值系统"",
  ""actors"": [
    {
      ""name"": ""用户"",
      ""modules"": [
        { ""name"": ""账户登录"", ""functions"": [] },
        { ""name"": ""查看话费余额"", ""functions"": [] },
        { ""name"": ""选择充值面额"", ""functions"": [] },
        { ""name"": ""输入充值手机号"", ""functions"": [] },
        { ""name"": ""核对充值信息"", ""functions"": [] },
        { ""name"": ""发起支付"", ""functions"": [] },
        { ""name"": ""查看充值订单记录"", ""functions"": [] },
        { ""name"": ""查询充值到账状态"", ""functions"": [] },
        { ""name"": ""申请充值退款"", ""functions"": [] }
      ]
    }
  ]
}";
        }

        private string GetSamplePlantUml()
        {
            return
@"@startuml 话费充值系统用例图
left to right direction
actor ""用户"" as User
rectangle 话费充值系统 {
    (账户登录)
    (查看话费余额)
    (选择充值面额)
    (输入充值手机号)
    (核对充值信息)
    (发起支付)
    (查看充值订单记录)
    (查询充值到账状态)
    (申请充值退款)
}
User -- (账户登录)
User -- (查看话费余额)
User -- (选择充值面额)
User -- (输入充值手机号)
User -- (核对充值信息)
User -- (发起支付)
User -- (查看充值订单记录)
User -- (查询充值到账状态)
User -- (申请充值退款)
@enduml";
        }

        private void btnImportJson_Click(object sender, EventArgs e)
        {
            try
            {
                if (!Clipboard.ContainsText())
                {
                    MessageBox.Show("剪贴板中没有JSON文本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string inputText = Clipboard.GetText();
                ImportInput(inputText);

                isInternalTextChange = true;
                try
                {
                    if (LooksLikePlantUml(inputText))
                    {
                        txtPlantUmlInput.Text = inputText;
                        tabActors.SelectedTab = tabPlantUml;
                    }
                    else
                    {
                        jsonEditor.SetJsonText(inputText, true);
                        tabActors.SelectedTab = tabJson;
                    }
                }
                finally
                {
                    isInternalTextChange = false;
                }

                RenderDiagram();
                MessageBox.Show("用例图生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog($"导入失败: {ex.Message}");
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportInput(string input)
        {
            string text = (input ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new InvalidOperationException("导入内容不能为空。");
            }

            if (LooksLikePlantUml(text))
            {
                ImportPlantUml(text);
                return;
            }

            ImportJson(text);
        }

        private bool LooksLikePlantUml(string text)
        {
            return text.IndexOf("@startuml", StringComparison.OrdinalIgnoreCase) >= 0
                || Regex.IsMatch(text, @"(?im)^\s*actor\s+")
                || Regex.IsMatch(text, @"\([^)]+\)\s*(?:as\s+\w+)?")
                || Regex.IsMatch(text, @"--\s*\(");
        }

        private void ImportJson(string json)
        {
            var serializer = new JavaScriptSerializer();
            UseCaseJsonRoot root = serializer.Deserialize<UseCaseJsonRoot>(json);
            if (root == null || root.actors == null || root.actors.Count == 0)
            {
                throw new InvalidOperationException("JSON中未找到actors数据。");
            }

            // Remove all tab pages except tabJson
            for (int i = tabActors.TabPages.Count - 1; i >= 0; i--)
            {
                if (tabActors.TabPages[i] != tabJson && tabActors.TabPages[i] != tabPlantUml)
                {
                    tabActors.TabPages.RemoveAt(i);
                }
            }

            systemName = string.IsNullOrWhiteSpace(root.system) ? "系统用例图" : root.system.Trim();
            txtTitle.Text = systemName + "用例图";

            int objectCount = 0;
            int relationCount = 0;
            foreach (UseCaseActorJson actor in root.actors)
            {
                if (actor.hidden) continue;
                string actorName = NormalizeName(actor.name, "参与者");
                var page = new TabPage(actorName);
                var grid = CreateActorGrid();
                page.Controls.Add(grid);
                tabActors.TabPages.Add(page);
                objectCount++;

                if (actor.modules == null)
                {
                    continue;
                }

                int index = 1;
                foreach (UseCaseModuleJson module in actor.modules)
                {
                    if (module.hidden) continue;
                    string moduleName = NormalizeName(module.name, "模块");
                    objectCount++;
                    relationCount++;

                    if (module.functions == null || module.functions.Count == 0)
                    {
                        grid.Rows.Add(index++, moduleName, string.Empty);
                        continue;
                    }

                    for (int i = 0; i < module.functions.Count; i++)
                    {
                        UseCaseFunctionJson function = module.functions[i];
                        if (function.hidden) continue;
                        string functionName = NormalizeName(function.name, "功能");
                        objectCount++;
                        relationCount++;
                        int rowIndex = grid.Rows.Add(index++, moduleName, functionName);
                        if (i == 0)
                        {
                            grid.Rows[rowIndex].DefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252);
                        }
                    }
                }
            }

            if (tabActors.TabPages.Count > 1)
            {
                tabActors.SelectedIndex = 1; // Select the first actor tab
            }

            AppendLog($"已成功导入{txtTitle.Text} {tabActors.TabPages.Count - 1}个参与者 {objectCount}个对象 {relationCount}组关系");
        }

        private void ImportPlantUml(string uml)
        {
            PlantUmlUseCaseModel model = ParsePlantUmlUseCase(uml);
            if (model.Actors.Count == 0)
            {
                model.Actors["Actor"] = "参与者";
            }
            if (model.UseCases.Count == 0)
            {
                throw new InvalidOperationException("PlantUML中未找到用例。");
            }

            ClearActorTabs();
            systemName = string.IsNullOrWhiteSpace(model.SystemName) ? "系统" : model.SystemName;
            txtTitle.Text = systemName + "用例图";

            Dictionary<string, List<string>> actorUseCases = BuildActorUseCaseMap(model);
            foreach (KeyValuePair<string, List<string>> pair in actorUseCases)
            {
                string actorName = NormalizeName(pair.Key, "参与者");
                var page = new TabPage(actorName);
                var grid = CreateActorGrid();
                page.Controls.Add(grid);
                tabActors.TabPages.Add(page);

                int index = 1;
                foreach (string useCase in pair.Value)
                {
                    grid.Rows.Add(index++, useCase, string.Empty);
                }
            }

            if (tabActors.TabPages.Count > 1)
            {
                tabActors.SelectedIndex = 1;
            }

            AppendLog($"已导入PlantUML用例图：{actorUseCases.Count}个参与者，{model.UseCases.Count}个用例。");
        }

        private PlantUmlUseCaseModel ParsePlantUmlUseCase(string uml)
        {
            var model = new PlantUmlUseCaseModel();
            var aliasToActor = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var aliasToUseCase = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            string[] lines = uml.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            foreach (string rawLine in lines)
            {
                string line = StripPlantUmlComment(rawLine).Trim();
                if (string.IsNullOrWhiteSpace(line)) continue;
                if (line.StartsWith("@", StringComparison.OrdinalIgnoreCase)) continue;
                if (line.Equals("left to right direction", StringComparison.OrdinalIgnoreCase)) continue;
                if (line == "{" || line == "}") continue;

                Match rectangle = Regex.Match(line, @"^\s*rectangle\s+""?([^""{]+)""?\s*\{?", RegexOptions.IgnoreCase);
                if (rectangle.Success)
                {
                    model.SystemName = rectangle.Groups[1].Value.Trim();
                    continue;
                }

                Match actor = Regex.Match(line, @"^\s*actor\s+""([^""]+)""\s*(?:as\s+([A-Za-z_][\w]*))?", RegexOptions.IgnoreCase);
                if (!actor.Success)
                {
                    actor = Regex.Match(line, @"^\s*actor\s+([^\s]+)\s*(?:as\s+([A-Za-z_][\w]*))?", RegexOptions.IgnoreCase);
                }
                if (actor.Success)
                {
                    string actorName = actor.Groups[1].Value.Trim();
                    string alias = actor.Groups[2].Success ? actor.Groups[2].Value.Trim() : actorName;
                    model.Actors[alias] = actorName;
                    aliasToActor[alias] = actorName;
                    continue;
                }

                Match useCaseAlias = Regex.Match(line, @"\(""([^""]+)""\)\s*(?:as\s+([A-Za-z_][\w]*))?", RegexOptions.IgnoreCase);
                if (!useCaseAlias.Success)
                {
                    useCaseAlias = Regex.Match(line, @"\(([^)]+)\)\s*(?:as\s+([A-Za-z_][\w]*))?", RegexOptions.IgnoreCase);
                }
                if (useCaseAlias.Success)
                {
                    string useCaseName = useCaseAlias.Groups[1].Value.Trim();
                    model.UseCases.Add(useCaseName);
                    if (useCaseAlias.Groups[2].Success)
                    {
                        aliasToUseCase[useCaseAlias.Groups[2].Value.Trim()] = useCaseName;
                    }
                }

                Match relation = Regex.Match(line, @"^\s*([A-Za-z_][\w]*|""[^""]+"")\s*(?:--|-\S*->|<-\S*-)\s*(\([^)]+\)|[A-Za-z_][\w]*|""[^""]+"")", RegexOptions.IgnoreCase);
                if (relation.Success)
                {
                    string left = ResolvePlantUmlName(relation.Groups[1].Value.Trim(), aliasToActor, aliasToUseCase);
                    string right = ResolvePlantUmlName(relation.Groups[2].Value.Trim(), aliasToActor, aliasToUseCase);
                    bool leftIsActor = aliasToActor.Values.Contains(left);
                    bool rightIsActor = aliasToActor.Values.Contains(right);
                    string actorName = leftIsActor ? left : (rightIsActor ? right : left);
                    string useCaseName = actorName == left ? right : left;
                    if (!string.IsNullOrWhiteSpace(actorName) && !string.IsNullOrWhiteSpace(useCaseName))
                    {
                        model.Relations.Add(new PlantUmlUseCaseRelation { Actor = actorName, UseCase = useCaseName });
                        model.UseCases.Add(useCaseName);
                    }
                }
            }

            return model;
        }

        private void ClearActorTabs()
        {
            for (int i = tabActors.TabPages.Count - 1; i >= 0; i--)
            {
                if (tabActors.TabPages[i] != tabJson && tabActors.TabPages[i] != tabPlantUml)
                {
                    tabActors.TabPages.RemoveAt(i);
                }
            }
        }

        private string StripPlantUmlComment(string line)
        {
            int quoteCount = 0;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"') quoteCount++;
                if (line[i] == '\'' && quoteCount % 2 == 0)
                {
                    return line.Substring(0, i);
                }
            }
            return line;
        }

        private string ResolvePlantUmlName(string token, Dictionary<string, string> actors, Dictionary<string, string> useCases)
        {
            string value = token.Trim().Trim('"');
            if (value.StartsWith("(") && value.EndsWith(")"))
            {
                value = value.Substring(1, value.Length - 2).Trim().Trim('"');
            }
            if (actors.ContainsKey(value)) return actors[value];
            if (useCases.ContainsKey(value)) return useCases[value];
            return value;
        }

        private Dictionary<string, List<string>> BuildActorUseCaseMap(PlantUmlUseCaseModel model)
        {
            var map = new Dictionary<string, List<string>>();
            foreach (string actor in model.Actors.Values.Distinct())
            {
                map[actor] = new List<string>();
            }

            foreach (PlantUmlUseCaseRelation relation in model.Relations)
            {
                string actor = string.IsNullOrWhiteSpace(relation.Actor) ? model.Actors.Values.FirstOrDefault() : relation.Actor;
                if (string.IsNullOrWhiteSpace(actor)) actor = "参与者";
                if (!map.ContainsKey(actor)) map[actor] = new List<string>();
                if (!map[actor].Contains(relation.UseCase)) map[actor].Add(relation.UseCase);
            }

            if (map.Count == 1 && model.Relations.Count == 0)
            {
                string actor = map.Keys.First();
                foreach (string useCase in model.UseCases)
                {
                    if (!map[actor].Contains(useCase)) map[actor].Add(useCase);
                }
            }

            return map;
        }

        private string NormalizeName(string value, string fallback)
        {
            string result = (value ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(result) ? fallback : result;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                RenderDiagram();
                MessageBox.Show("用例图生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog($"生成失败: {ex.Message}");
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderDiagram()
        {
            List<UseCaseRelation> relations = ReadRelations();
            if (relations.Count == 0)
            {
                throw new InvalidOperationException("请先导入JSON或填写关系表。");
            }

            HashSet<string> actorSet = ReadActorSet();
            currentFontName = cmbFontName.SelectedItem == null ? "宋体" : cmbFontName.SelectedItem.ToString();
            currentFontSize = ParseFontSize(cmbFontSize.SelectedItem);
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            Visio.Page page = GetOrCreateActivePage(visioApp);
            ClearPageShapes(page);
            DrawUseCaseDiagram(page, relations, actorSet);
            AppendLog("用例图生成完成。");
        }

        private void jsonEditor_JsonTextChanged(object sender, EventArgs e)
        {
            if (isInternalTextChange)
            {
                return;
            }

            string text = jsonEditor.JsonText.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            // Check if it matches a JSON object format roughly before parsing
            if (text.StartsWith("{") && text.EndsWith("}"))
            {
                try
                {
                    ImportJson(text);
                }
                catch
                {
                    // Do not show errors to avoid interrupting user paste/edit
                }
            }
        }

        private void txtPlantUmlInput_TextChanged(object sender, EventArgs e)
        {
            if (isInternalTextChange)
            {
                return;
            }

            string text = txtPlantUmlInput.Text.Trim();
            if (string.IsNullOrEmpty(text) || !LooksLikePlantUml(text))
            {
                return;
            }

            try
            {
                ImportPlantUml(text);
            }
            catch
            {
                // Do not show errors to avoid interrupting user paste/edit
            }
        }

        private List<UseCaseRelation> ReadRelations()
        {
            var relations = new List<UseCaseRelation>();
            foreach (TabPage page in tabActors.TabPages)
            {
                if (page == tabJson || page == tabPlantUml)
                {
                    continue;
                }
                string actorName = NormalizeName(page.Text, "参与者");
                DataGridView grid = page.Controls.OfType<DataGridView>().FirstOrDefault();
                if (grid == null)
                {
                    continue;
                }

                string currentModule = string.Empty;
                var actorModules = new HashSet<string>();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    string module = CellText(row, "Module");
                    string function = CellText(row, "Function");
                    if (!string.IsNullOrWhiteSpace(module))
                    {
                        currentModule = module;
                        if (!actorModules.Contains(currentModule))
                        {
                            relations.Add(new UseCaseRelation
                            {
                                Source = actorName,
                                Target = currentModule,
                                Relation = "关联",
                                Text = string.Empty
                            });
                            actorModules.Add(currentModule);
                        }
                    }

                    if (chkShowFunctionNodes.Checked && !string.IsNullOrWhiteSpace(currentModule) && !string.IsNullOrWhiteSpace(function))
                    {
                        relations.Add(new UseCaseRelation
                        {
                            Source = currentModule,
                            Target = function,
                            Relation = "包含",
                            Text = string.Empty
                        });
                    }
                }
            }

            return relations;
        }

        private HashSet<string> ReadActorSet()
        {
            var actorSet = new HashSet<string>();
            foreach (TabPage page in tabActors.TabPages)
            {
                if (page == tabJson || page == tabPlantUml)
                {
                    continue;
                }
                string name = Convert.ToString(page.Text);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    actorSet.Add(name);
                }
            }

            return actorSet;
        }

        private string CellText(DataGridViewRow row, string columnName)
        {
            object value = row.Cells[columnName].Value;
            return value == null ? string.Empty : Convert.ToString(value).Trim();
        }

        private void DrawUseCaseDiagram(Visio.Page page, List<UseCaseRelation> relations, HashSet<string> actorSet)
        {
            double pageWidth = page.PageSheet.CellsU["PageWidth"].Result["in"];
            double pageHeight = page.PageSheet.CellsU["PageHeight"].Result["in"];
            double horizontalGap = ParsePositiveNumber(txtHorizontalSpacing.Text, "横向间距") / 25.4;
            double verticalGap = ParsePositiveNumber(txtVerticalSpacing.Text, "纵向间距") / 25.4;
            double titleY = pageHeight - 0.35;

            string title = string.IsNullOrWhiteSpace(txtTitle.Text) ? systemName + "用例图" : txtTitle.Text.Trim();
            Visio.Shape titleShape = page.DrawRectangle(pageWidth / 2.0 - 1.35, titleY - 0.18, pageWidth / 2.0 + 1.35, titleY + 0.18);
            titleShape.Text = title;
            ApplyTextShapeStyle(titleShape, currentFontName, currentFontSize + 1.5, true);
            ApplyBoxStyle(titleShape);

            List<string> allObjects = GetObjectsFromRelations(relations);
            List<string> actors = allObjects.Where(actorSet.Contains).ToList();
            List<string> useCases = allObjects.Where(name => !actorSet.Contains(name)).ToList();

            Dictionary<string, ShapeAnchor> anchors = new Dictionary<string, ShapeAnchor>();
            double boundaryLeft = 2.15;
            double boundaryRight = pageWidth - 0.75;
            double boundaryTop = pageHeight - 0.9;
            double boundaryBottom = 0.65;

            if (cmbLayout.SelectedItem != null && cmbLayout.SelectedItem.ToString() == "参与者均在右侧")
            {
                boundaryLeft = 0.75;
                boundaryRight = pageWidth - 2.15;
            }
            else if (cmbLayout.SelectedItem != null && cmbLayout.SelectedItem.ToString() == "参与者左右分布")
            {
                boundaryLeft = 1.65;
                boundaryRight = pageWidth - 1.65;
            }

            Visio.Shape boundary = page.DrawRectangle(boundaryLeft, boundaryBottom, boundaryRight, boundaryTop);
            boundary.Text = systemName;
            ApplyBoundaryStyle(boundary);

            DrawActors(page, actors, anchors, boundaryLeft, boundaryRight, boundaryBottom, boundaryTop);
            DrawUseCases(page, useCases, relations, anchors, boundaryLeft, boundaryRight, boundaryBottom, boundaryTop, horizontalGap, verticalGap);

            foreach (UseCaseRelation relation in relations)
            {
                if (!anchors.ContainsKey(relation.Source) || !anchors.ContainsKey(relation.Target))
                {
                    continue;
                }

                DrawRelation(page, anchors[relation.Source], anchors[relation.Target], relation);
            }
        }

        private List<string> GetObjectsFromRelations(List<UseCaseRelation> relations)
        {
            var names = new List<string>();
            foreach (UseCaseRelation relation in relations)
            {
                if (!names.Contains(relation.Source))
                {
                    names.Add(relation.Source);
                }

                if (!names.Contains(relation.Target))
                {
                    names.Add(relation.Target);
                }
            }

            return names;
        }

        private void DrawActors(Visio.Page page, List<string> actors, Dictionary<string, ShapeAnchor> anchors, double boundaryLeft, double boundaryRight, double boundaryBottom, double boundaryTop)
        {
            if (actors.Count == 0)
            {
                return;
            }

            string layout = cmbLayout.SelectedItem == null ? "参与者均在左侧" : cmbLayout.SelectedItem.ToString();
            List<string> leftActors = new List<string>();
            List<string> rightActors = new List<string>();
            if (layout == "参与者均在右侧")
            {
                rightActors.AddRange(actors);
            }
            else if (layout == "参与者左右分布")
            {
                for (int i = 0; i < actors.Count; i++)
                {
                    if (i % 2 == 0)
                    {
                        leftActors.Add(actors[i]);
                    }
                    else
                    {
                        rightActors.Add(actors[i]);
                    }
                }
            }
            else
            {
                leftActors.AddRange(actors);
            }

            DrawActorColumn(page, leftActors, anchors, boundaryLeft - 0.9, boundaryBottom, boundaryTop);
            DrawActorColumn(page, rightActors, anchors, boundaryRight + 0.9, boundaryBottom, boundaryTop);
        }

        private void DrawActorColumn(Visio.Page page, List<string> actors, Dictionary<string, ShapeAnchor> anchors, double x, double bottom, double top)
        {
            if (actors.Count == 0)
            {
                return;
            }

            double usableHeight = Math.Max(1.0, top - bottom - 0.6);
            double gap = actors.Count == 1 ? 0 : usableHeight / (actors.Count - 1);
            double startY = actors.Count == 1 ? (top + bottom) / 2.0 : top - 0.45;
            for (int i = 0; i < actors.Count; i++)
            {
                double y = startY - i * gap;
                Visio.Shape actorShape;
                if (cmbActorShape.SelectedItem != null && cmbActorShape.SelectedItem.ToString() == "Visio自带")
                {
                    actorShape = DrawVisioActor(page, actors[i], x, y);
                }
                else
                {
                    actorShape = DrawStickActor(page, actors[i], x, y);
                }

                anchors[actors[i]] = new ShapeAnchor { Name = actors[i], X = x, Y = y, Width = 0.62, Height = 0.95, IsActor = true, Shape = actorShape };
            }
        }

        private void DrawUseCases(Visio.Page page, List<string> useCases, List<UseCaseRelation> relations, Dictionary<string, ShapeAnchor> anchors, double left, double right, double bottom, double top, double horizontalGap, double verticalGap)
        {
            if (useCases.Count == 0)
            {
                return;
            }

            HashSet<string> moduleNames = new HashSet<string>(relations.Where(r => r.Relation == "包含").Select(r => r.Source));
            List<string> modules = useCases.Where(moduleNames.Contains).ToList();
            List<string> functions = useCases.Where(name => !moduleNames.Contains(name)).ToList();
            double moduleX = left + (right - left) * 0.36;
            double functionX = left + (right - left) * 0.68;
            if (functions.Count == 0)
            {
                moduleX = (left + right) / 2.0;
            }

            PlaceUseCaseColumn(page, modules.Count > 0 ? modules : useCases, anchors, moduleX, bottom, top, verticalGap);
            if (modules.Count > 0)
            {
                PlaceUseCaseColumn(page, functions, anchors, functionX + Math.Max(0, horizontalGap - 0.3), bottom, top, verticalGap);
            }
        }

        private void PlaceUseCaseColumn(Visio.Page page, List<string> names, Dictionary<string, ShapeAnchor> anchors, double x, double bottom, double top, double verticalGap)
        {
            if (names.Count == 0)
            {
                return;
            }

            double ellipseHeight = 0.42;
            double requiredHeight = names.Count * ellipseHeight + Math.Max(0, names.Count - 1) * Math.Max(0.16, verticalGap);
            double startY = (top + bottom) / 2.0 + requiredHeight / 2.0 - ellipseHeight / 2.0;
            for (int i = 0; i < names.Count; i++)
            {
                double y = startY - i * (ellipseHeight + Math.Max(0.16, verticalGap));
                double width = EstimateTextWidthInches(names[i], currentFontSize, 0.95);
                Visio.Shape shape = page.DrawOval(x - width / 2.0, y - ellipseHeight / 2.0, x + width / 2.0, y + ellipseHeight / 2.0);
                shape.Text = names[i];
                ApplyTextShapeStyle(shape, currentFontName, currentFontSize, false);
                ApplyUseCaseStyle(shape);
                anchors[names[i]] = new ShapeAnchor { Name = names[i], X = x, Y = y, Width = width, Height = ellipseHeight, IsActor = false, Shape = shape };
            }
        }

        private Visio.Shape DrawVisioActor(Visio.Page page, string name, double x, double y)
        {
            var shapes = new List<Visio.Shape>();
            double headR = 0.12;
            Visio.Shape head = page.DrawOval(x - headR, y + 0.28, x + headR, y + 0.52);
            ApplyBoxStyle(head);
            shapes.Add(head);
            shapes.Add(DrawStyledLine(page, x, y + 0.28, x, y - 0.17, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x - 0.28, y + 0.08, x + 0.28, y + 0.08, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x, y - 0.17, x - 0.26, y - 0.5, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x, y - 0.17, x + 0.26, y - 0.5, false, string.Empty));
            Visio.Shape label = page.DrawRectangle(x - 0.52, y - 0.82, x + 0.52, y - 0.56);
            label.Text = name;
            ApplyTextShapeStyle(label, currentFontName, currentFontSize, false);
            TrySetFormula(label, "FillPattern", "0");
            TrySetFormula(label, "LinePattern", "0");
            shapes.Add(label);
            return GroupShapes(page, shapes);
        }

        private Visio.Shape DrawStickActor(Visio.Page page, string name, double x, double y)
        {
            var shapes = new List<Visio.Shape>();
            double headR = 0.12;
            Visio.Shape head = page.DrawOval(x - headR, y + 0.24, x + headR, y + 0.48);
            ApplyBoxStyle(head);
            shapes.Add(head);
            shapes.Add(DrawStyledLine(page, x, y + 0.24, x, y - 0.18, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x - 0.25, y + 0.08, x + 0.25, y + 0.08, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x, y - 0.18, x - 0.24, y - 0.48, false, string.Empty));
            shapes.Add(DrawStyledLine(page, x, y - 0.18, x + 0.24, y - 0.48, false, string.Empty));
            Visio.Shape label = page.DrawRectangle(x - 0.48, y - 0.78, x + 0.48, y - 0.52);
            label.Text = name;
            ApplyTextShapeStyle(label, currentFontName, currentFontSize, false);
            TrySetFormula(label, "FillPattern", "0");
            TrySetFormula(label, "LinePattern", "0");
            shapes.Add(label);
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

        private void DrawRelation(Visio.Page page, ShapeAnchor source, ShapeAnchor target, UseCaseRelation relation)
        {
            double x1;
            double y1;
            double x2;
            double y2;
            GetUnifiedRelationPoint(source, target, relation, true, out x1, out y1);
            GetUnifiedRelationPoint(target, source, relation, false, out x2, out y2);

            bool dashed = relation.Relation == "包含" || relation.Relation == "扩展" || relation.Relation == "泛化";
            string text = BuildLineText(relation);
            Visio.Shape line = DrawDynamicConnector(page, source, target, x1, y1, x2, y2, dashed, text);
            if (relation.Relation == "关联")
            {
                TrySetFormula(line, "EndArrow", chkAssociationArrow.Checked ? "4" : "0");
            }
            else
            {
                TrySetFormula(line, "EndArrow", "4");
            }
        }

        private void GetUnifiedRelationPoint(ShapeAnchor shape, ShapeAnchor other, UseCaseRelation relation, bool isSource, out double x, out double y)
        {
            if (relation.Relation == "关联")
            {
                if (shape.IsActor)
                {
                    x = shape.X + (other.X >= shape.X ? shape.Width / 2.0 : -shape.Width / 2.0);
                    y = shape.Y;
                    return;
                }

                GetShapeEdgePoint(shape, other.X, other.Y, out x, out y);
                return;
            }

            if (isSource)
            {
                x = shape.X + (other.X >= shape.X ? shape.Width / 2.0 : -shape.Width / 2.0);
                y = shape.Y;
                return;
            }

            GetShapeEdgePoint(shape, other.X, other.Y, out x, out y);
        }

        private string BuildLineText(UseCaseRelation relation)
        {
            if (!string.IsNullOrWhiteSpace(relation.Text))
            {
                return relation.Text.Trim();
            }

            if (relation.Relation == "包含")
            {
                return chkEnglishRelationText.Checked ? "<<include>>" : "《包含》";
            }

            if (relation.Relation == "扩展")
            {
                return chkEnglishRelationText.Checked ? "<<extend>>" : "《扩展》";
            }

            if (relation.Relation == "泛化")
            {
                return chkEnglishRelationText.Checked ? "generalization" : "泛化";
            }

            return string.Empty;
        }

        private Visio.Shape DrawStyledLine(Visio.Page page, double x1, double y1, double x2, double y2, bool dashed, string text)
        {
            Visio.Shape line = page.DrawLine(x1, y1, x2, y2);
            line.Text = text ?? string.Empty;
            TrySetFormula(line, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(line, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(line, "LinePattern", dashed ? "2" : "1");
            TrySetFormula(line, "BeginArrow", "0");
            TrySetFormula(line, "EndArrow", "0");
            TrySetFormula(line, "Char.Font", $"\"{currentFontName}\"");
            TrySetFormula(line, "Char.Size", $"{Math.Max(8, currentFontSize - 1).ToString(CultureInfo.InvariantCulture)}pt");
            return line;
        }

        private Visio.Shape DrawDynamicConnector(Visio.Page page, ShapeAnchor source, ShapeAnchor target, double x1, double y1, double x2, double y2, bool dashed, string text)
        {
            Visio.Shape connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0);
            connector.Text = text ?? string.Empty;
            TrySetFormula(connector, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(connector, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(connector, "LinePattern", dashed ? "2" : "1");
            TrySetFormula(connector, "BeginArrow", "0");
            TrySetFormula(connector, "EndArrow", "0");
            TrySetFormula(connector, "Char.Font", $"\"{currentFontName}\"");
            TrySetFormula(connector, "Char.Size", $"{Math.Max(8, currentFontSize - 1).ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(connector, "BeginX", $"{x1.ToString(CultureInfo.InvariantCulture)} in");
            TrySetFormula(connector, "BeginY", $"{y1.ToString(CultureInfo.InvariantCulture)} in");
            TrySetFormula(connector, "EndX", $"{x2.ToString(CultureInfo.InvariantCulture)} in");
            TrySetFormula(connector, "EndY", $"{y2.ToString(CultureInfo.InvariantCulture)} in");
            TrySetFormula(connector, "ConLineRouteExt", "1");
            TrySetFormula(connector, "ShapeRouteStyle", "16");

            GlueConnector(connector, source, target, x1, y1, x2, y2);
            return connector;
        }

        private void GlueConnector(Visio.Shape connector, ShapeAnchor source, ShapeAnchor target, double x1, double y1, double x2, double y2)
        {
            try
            {
                if (source.Shape != null)
                {
                    Visio.Cell sourceCell = EnsureConnectionPoint(source.Shape, x1 - source.X, y1 - source.Y);
                    connector.CellsU["BeginX"].GlueTo(sourceCell);
                }

                if (target.Shape != null)
                {
                    Visio.Cell targetCell = EnsureConnectionPoint(target.Shape, x2 - target.X, y2 - target.Y);
                    connector.CellsU["EndX"].GlueTo(targetCell);
                }
            }
            catch
            {
            }
        }

        private Visio.Cell EnsureConnectionPoint(Visio.Shape shape, double offsetX, double offsetY)
        {
            short row = shape.AddRow(
                (short)Visio.VisSectionIndices.visSectionConnectionPts,
                (short)Visio.VisRowIndices.visRowLast,
                (short)Visio.VisRowTags.visTagCnnctPt);
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, row, (short)Visio.VisCellIndices.visCnnctX].FormulaU =
                $"Width*0.5+{offsetX.ToString(CultureInfo.InvariantCulture)} in";
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, row, (short)Visio.VisCellIndices.visCnnctY].FormulaU =
                $"Height*0.5+{offsetY.ToString(CultureInfo.InvariantCulture)} in";
            return shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, row, (short)Visio.VisCellIndices.visCnnctX];
        }

        private void GetShapeEdgePoint(ShapeAnchor shape, double targetX, double targetY, out double edgeX, out double edgeY)
        {
            double dx = targetX - shape.X;
            double dy = targetY - shape.Y;
            if (Math.Abs(dx) < 0.0001 && Math.Abs(dy) < 0.0001)
            {
                edgeX = shape.X;
                edgeY = shape.Y;
                return;
            }

            if (!shape.IsActor)
            {
                // For oval (ellipse) shapes, use ellipse boundary intersection to prevent gaps
                double a = shape.Width / 2.0;
                double b = shape.Height / 2.0;
                double denominator = Math.Sqrt(b * b * dx * dx + a * a * dy * dy);
                if (denominator > 0.0001)
                {
                    edgeX = shape.X + (a * b * dx) / denominator;
                    edgeY = shape.Y + (a * b * dy) / denominator;
                    return;
                }
            }

            // Fallback for rectangular shapes or stick actors
            double scaleX = Math.Abs(dx) < 0.0001 ? double.MaxValue : shape.Width / 2.0 / Math.Abs(dx);
            double scaleY = Math.Abs(dy) < 0.0001 ? double.MaxValue : shape.Height / 2.0 / Math.Abs(dy);
            double scale = Math.Min(scaleX, scaleY);
            edgeX = shape.X + dx * scale;
            edgeY = shape.Y + dy * scale;
        }

        private double EstimateTextWidthInches(string text, double fontSize, double minWidth)
        {
            int count = Math.Max(1, (text ?? string.Empty).Length);
            return Math.Max(minWidth, count * fontSize * 0.0105 + 0.36);
        }

        private void ApplyTextShapeStyle(Visio.Shape shape, string fontName, double fontSize, bool bold)
        {
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSize.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(0, 0, 0)");
            TrySetFormula(shape, "Char.Style", bold ? "1" : "0");
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

        private void ApplyUseCaseStyle(Visio.Shape shape)
        {
            TrySetFormula(shape, "FillPattern", "0");
            if (chkUseCaseOutline.Checked)
            {
                TrySetFormula(shape, "LinePattern", "1");
                TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
                TrySetFormula(shape, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            }
            else
            {
                TrySetFormula(shape, "LinePattern", "0");
            }
        }

        private void ApplyBoundaryStyle(Visio.Shape shape)
        {
            TrySetFormula(shape, "FillPattern", "0");
            TrySetFormula(shape, "LineColor", "RGB(0, 0, 0)");
            TrySetFormula(shape, "LineWeight", $"{currentLineWidth.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Font", $"\"{currentFontName}\"");
            TrySetFormula(shape, "Char.Size", $"{currentFontSize.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Para.HorzAlign", "0");
            TrySetFormula(shape, "VerticalAlign", "0");
            TrySetFormula(shape, "TxtPinX", "Width*0.08");
            TrySetFormula(shape, "TxtPinY", "Height-0.18");
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

        private double ParseFontSize(object selectedItem)
        {
            string fontSizeName = selectedItem == null ? "五号" : selectedItem.ToString();
            switch (fontSizeName)
            {
                case "小三":
                    return 15;
                case "四号":
                    return 14;
                case "小四":
                    return 12;
                case "小五":
                    return 9;
                case "五号":
                default:
                    return 10.5;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
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

        private class UseCaseJsonRoot
        {
            public string system { get; set; }
            public List<UseCaseActorJson> actors { get; set; }
        }

        private class UseCaseActorJson
        {
            public string name { get; set; }
            public bool hidden { get; set; }
            public List<UseCaseModuleJson> modules { get; set; }
        }

        private class UseCaseModuleJson
        {
            public string name { get; set; }
            public bool hidden { get; set; }
            public List<UseCaseFunctionJson> functions { get; set; }
        }

        private class UseCaseFunctionJson
        {
            public string name { get; set; }
            public bool hidden { get; set; }
        }

        private class UseCaseRelation
        {
            public string Source { get; set; }
            public string Target { get; set; }
            public string Relation { get; set; }
            public string Text { get; set; }
        }

        private class PlantUmlUseCaseModel
        {
            public string SystemName { get; set; }
            public Dictionary<string, string> Actors { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public HashSet<string> UseCases { get; } = new HashSet<string>();
            public List<PlantUmlUseCaseRelation> Relations { get; } = new List<PlantUmlUseCaseRelation>();
        }

        private class PlantUmlUseCaseRelation
        {
            public string Actor { get; set; }
            public string UseCase { get; set; }
        }

        private class ShapeAnchor
        {
            public string Name { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public bool IsActor { get; set; }
            public Visio.Shape Shape { get; set; }
        }
    }
}
