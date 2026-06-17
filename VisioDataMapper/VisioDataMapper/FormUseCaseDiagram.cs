using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
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
        private TextBox txtStatus;
        private TabControl tabActors;
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
            Size = new Size(960, 900);
            MinimumSize = new Size(900, 820);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            Font defaultFont = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            Font = defaultFont;

            Label lblTip = new Label
            {
                Text = "请将UML用例JSON复制到剪贴板，点击导入JSON后按参与者生成多页表格",
                Location = new Point(15, 15),
                Size = new Size(470, 25),
                ForeColor = Color.DarkRed
            };

            btnImportJson = new Button { Text = "导入JSON", Location = new Point(15, 50), Size = new Size(130, 44) };
            btnImportJson.Click += btnImportJson_Click;

            Label lblTitle = new Label { Text = "大标题:", Location = new Point(455, 61), Size = new Size(70, 24), TextAlign = ContentAlignment.MiddleRight };
            txtTitle = new TextBox { Location = new Point(535, 59), Size = new Size(340, 26), Text = systemName };

            tabActors = new TabControl
            {
                Location = new Point(15, 110),
                Size = new Size(860, 455)
            };

            txtStatus = new TextBox
            {
                Location = new Point(15, 580),
                Size = new Size(860, 78),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = false,
                BackColor = Color.White,
                Text = $"{DateTime.Now:HH:mm:ss} 请复制JSON后点击导入JSON，或在此粘贴JSON直接渲染。"
            };
            txtStatus.TextChanged += txtStatus_TextChanged;

            lblActorShape = new Label { Text = "参与者形状:", Location = new Point(20, 700), Size = new Size(90, 24), TextAlign = ContentAlignment.MiddleLeft };
            cmbActorShape = new ComboBox { Location = new Point(120, 697), Size = new Size(170, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbActorShape.Items.AddRange(new string[] { "Draw.Io风格", "Visio自带" });
            cmbActorShape.SelectedIndex = 0;

            chkEnglishRelationText = new CheckBox { Text = "关系线条使用英文字符", Location = new Point(310, 700), Size = new Size(170, 24), Checked = true };
            chkShowFunctionNodes = new CheckBox { Text = "显示功能节点", Location = new Point(500, 700), Size = new Size(125, 24), Checked = true };
            chkUseCaseOutline = new CheckBox { Text = "模块/功能添加外边框", Location = new Point(645, 700), Size = new Size(165, 24), Checked = true };

            lblLayout = new Label { Text = "排版:", Location = new Point(20, 742), Size = new Size(70, 24), TextAlign = ContentAlignment.MiddleLeft };
            cmbLayout = new ComboBox { Location = new Point(80, 739), Size = new Size(190, 26), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLayout.Items.AddRange(new string[] { "参与者均在左侧", "参与者均在右侧", "参与者左右分布" });
            cmbLayout.SelectedIndex = 0;

            lblFontName = new Label { Text = "字体:", Location = new Point(290, 742), Size = new Size(45, 24), TextAlign = ContentAlignment.MiddleLeft };
            cmbFontName = CreateFontCombo(new Point(335, 739));
            lblFontSize = new Label { Text = "字号:", Location = new Point(455, 742), Size = new Size(45, 24), TextAlign = ContentAlignment.MiddleLeft };
            cmbFontSize = CreateFontSizeCombo(new Point(500, 739));

            lblSpacing = new Label { Text = "横纵间距:", Location = new Point(20, 784), Size = new Size(78, 24), TextAlign = ContentAlignment.MiddleLeft };
            txtHorizontalSpacing = new TextBox { Text = "15", Location = new Point(100, 781), Size = new Size(48, 26) };
            txtVerticalSpacing = new TextBox { Text = "8", Location = new Point(160, 781), Size = new Size(48, 26) };
            lblMillimeter = new Label { Text = "mm", Location = new Point(216, 784), Size = new Size(35, 24) };
            chkAssociationArrow = new CheckBox { Text = "关联连线画箭头", Location = new Point(310, 777), Size = new Size(145, 24), Checked = false };

            btnGenerate = new Button { Text = "生成绘图", Location = new Point(610, 744), Size = new Size(120, 48), BackColor = Color.LightSkyBlue, Font = new Font(defaultFont, FontStyle.Bold) };
            btnGenerate.Click += btnGenerate_Click;
            btnClose = new Button { Text = "关闭", Location = new Point(760, 744), Size = new Size(115, 48) };
            btnClose.Click += btnClose_Click;

            Controls.Add(lblTip);
            Controls.Add(btnImportJson);
            Controls.Add(lblTitle);
            Controls.Add(txtTitle);
            Controls.Add(tabActors);
            Controls.Add(txtStatus);
            Controls.Add(lblActorShape);
            Controls.Add(cmbActorShape);
            Controls.Add(chkEnglishRelationText);
            Controls.Add(chkShowFunctionNodes);
            Controls.Add(chkUseCaseOutline);
            Controls.Add(lblLayout);
            Controls.Add(cmbLayout);
            Controls.Add(lblFontName);
            Controls.Add(cmbFontName);
            Controls.Add(lblFontSize);
            Controls.Add(cmbFontSize);
            Controls.Add(lblSpacing);
            Controls.Add(txtHorizontalSpacing);
            Controls.Add(txtVerticalSpacing);
            Controls.Add(lblMillimeter);
            Controls.Add(chkAssociationArrow);
            Controls.Add(btnGenerate);
            Controls.Add(btnClose);

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
            return grid;
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
            if (tabActors == null)
            {
                return;
            }

            int margin = 15;
            int width = Math.Max(820, ClientSize.Width - margin * 2);
            int buttonTop = ClientSize.Height - 62;
            int optionRow2Top = buttonTop - 45;
            int optionRow1Top = optionRow2Top - 42;
            int statusBottom = optionRow1Top - 15;
            txtTitle.Left = Math.Max(535, margin + width - txtTitle.Width);
            tabActors.Width = width;
            tabActors.Height = Math.Max(300, statusBottom - 105 - tabActors.Top);
            txtStatus.Top = tabActors.Bottom + 15;
            txtStatus.Width = width;
            txtStatus.Height = Math.Max(54, statusBottom - txtStatus.Top);

            lblActorShape.Top = optionRow1Top;
            cmbActorShape.Top = optionRow1Top - 3;
            chkEnglishRelationText.Top = optionRow1Top;
            chkShowFunctionNodes.Top = optionRow1Top;
            chkUseCaseOutline.Top = optionRow1Top;
            lblLayout.Top = optionRow2Top;
            cmbLayout.Top = optionRow2Top - 3;
            lblFontName.Top = optionRow2Top;
            cmbFontName.Top = optionRow2Top - 3;
            lblFontSize.Top = optionRow2Top;
            cmbFontSize.Top = optionRow2Top - 3;
            lblSpacing.Top = buttonTop + 12;
            txtHorizontalSpacing.Top = buttonTop + 9;
            txtVerticalSpacing.Top = buttonTop + 9;
            lblMillimeter.Top = buttonTop + 12;
            chkAssociationArrow.Top = buttonTop + 12;
            btnGenerate.Top = buttonTop;
            btnClose.Top = buttonTop;
            btnClose.Left = margin + width - btnClose.Width;
            btnGenerate.Left = btnClose.Left - btnGenerate.Width - 25;
        }

        private void FormUseCaseDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
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

                string jsonText = Clipboard.GetText();
                ImportJson(jsonText);

                isInternalTextChange = true;
                try
                {
                    txtStatus.Text = jsonText;
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

        private void ImportJson(string json)
        {
            var serializer = new JavaScriptSerializer();
            UseCaseJsonRoot root = serializer.Deserialize<UseCaseJsonRoot>(json);
            if (root == null || root.actors == null || root.actors.Count == 0)
            {
                throw new InvalidOperationException("JSON中未找到actors数据。");
            }

            tabActors.TabPages.Clear();

            systemName = string.IsNullOrWhiteSpace(root.system) ? "系统用例图" : root.system.Trim();
            txtTitle.Text = systemName + "用例图";

            int objectCount = 0;
            int relationCount = 0;
            foreach (UseCaseActorJson actor in root.actors)
            {
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

            AppendLog($"已成功导入{txtTitle.Text} {tabActors.TabPages.Count}个参与者 {objectCount}个对象 {relationCount}组关系");
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

        private void txtStatus_TextChanged(object sender, EventArgs e)
        {
            if (isInternalTextChange)
            {
                return;
            }

            string text = txtStatus.Text.Trim();
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
                    RenderDiagram();
                }
                catch
                {
                    // Do not show errors to avoid interrupting user paste/edit
                }
            }
        }

        private List<UseCaseRelation> ReadRelations()
        {
            var relations = new List<UseCaseRelation>();
            foreach (TabPage page in tabActors.TabPages)
            {
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
            public List<UseCaseModuleJson> modules { get; set; }
        }

        private class UseCaseModuleJson
        {
            public string name { get; set; }
            public List<UseCaseFunctionJson> functions { get; set; }
        }

        private class UseCaseFunctionJson
        {
            public string name { get; set; }
        }

        private class UseCaseRelation
        {
            public string Source { get; set; }
            public string Target { get; set; }
            public string Relation { get; set; }
            public string Text { get; set; }
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
