using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Visio = Microsoft.Office.Interop.Visio;

namespace VisioDataMapper
{
    public class FormSystemArchitectureDiagram : Form
    {
        private Button btnAiPrompt;
        private Button btnImportJson;
        private Button btnGenerate;
        private Button btnClose;
        private Label lblTitle;
        private TextBox txtTitle;
        private TabControl tabControl;
        private TabPage tabJson;
        private TabPage tabNodes;
        private TabPage tabConnections;
        private TextBox txtJson;
        private DataGridView gridNodes;
        private DataGridView gridConnections;
        private TextBox txtStatus;
        private Panel pnlOptions;
        private ComboBox cmbFontName;
        private ComboBox cmbFontSize;
        private TextBox txtLayerSpacing;
        private TextBox txtModuleSpacing;
        private CheckBox chkCompactLayout;

        private const string DefaultUiFontName = "Microsoft YaHei";
        private const string DefaultDrawingFontName = "宋体";
        private const double DefaultDrawingFontSizePt = 10.5;
        private const double DefaultLayerSpacingMm = 6.0;
        private const double DefaultModuleSpacingMm = 5.0;
        private const bool DefaultCompactLayoutChecked = true;
        private const double DefaultPageWidthInch = 7.4;
        private const double DefaultLeftMarginInch = 0.42;
        private const double DefaultTopMarginInch = 0.38;
        private const double LayerLabelWidthInch = 0.72;
        private const double ModuleHeightInch = 0.32;
        private const double DatabaseWidthInch = 1.34;
        private const double DatabaseHeightInch = 0.46;
        private const double LayerInnerPaddingInch = 0.14;
        private const double GroupTitleHeightInch = 0.30;
        private const string AcademicBorderColor = "RGB(47, 55, 68)";
        private const string AcademicConnectorColor = "RGB(58, 67, 82)";

        public FormSystemArchitectureDiagram()
        {
            InitializeComponent();
            LoadSampleData();
        }

        private void InitializeComponent()
        {
            Text = "智能画图-系统架构图";
            Size = new Size(980, 820);
            MinimumSize = new Size(900, 720);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = false;
            BackColor = Color.FromArgb(240, 244, 248);
            Font = new Font(DefaultUiFontName, 9F, FontStyle.Regular);

            Label lblTip = new Label
            {
                Text = "请将系统说明交给 AI 生成规范 JSON，再导入 JSON 生成系统架构图。",
                Location = new Point(15, 14),
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104),
                Font = new Font(DefaultUiFontName, 9.5F, FontStyle.Regular)
            };

            btnAiPrompt = CreatePrimaryButton("AI生成提示词", new Point(15, 42), new Size(120, 32), Color.FromArgb(91, 76, 196));
            btnAiPrompt.Click += btnAiPrompt_Click;

            btnImportJson = CreatePrimaryButton("导入JSON", new Point(145, 42), new Size(105, 32), Color.FromArgb(0, 122, 255));
            btnImportJson.Click += btnImportJson_Click;

            lblTitle = new Label { Text = "架构图标题:", Location = new Point(285, 48), AutoSize = true, ForeColor = Color.FromArgb(74, 85, 104) };
            txtTitle = new TextBox { Location = new Point(375, 44), Size = new Size(560, 25), Font = new Font(DefaultUiFontName, 10F, FontStyle.Regular) };

            tabControl = new TabControl { Location = new Point(15, 88), Size = new Size(930, 500), Font = new Font(DefaultUiFontName, 9F, FontStyle.Regular) };
            tabJson = new TabPage("JSON代码") { BackColor = Color.White };
            txtJson = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Dock = DockStyle.Fill,
                Font = new Font("Consolas", 10F, FontStyle.Regular),
                BorderStyle = BorderStyle.None,
                BackColor = Color.White
            };
            tabJson.Controls.Add(txtJson);
            tabControl.TabPages.Add(tabJson);

            tabNodes = new TabPage("节点表") { BackColor = Color.White };
            gridNodes = CreateGrid();
            BuildNodeGridColumns();
            tabNodes.Controls.Add(gridNodes);
            tabControl.TabPages.Add(tabNodes);

            tabConnections = new TabPage("连接表") { BackColor = Color.White };
            gridConnections = CreateGrid();
            BuildConnectionGridColumns();
            tabConnections.Controls.Add(gridConnections);
            tabControl.TabPages.Add(tabConnections);

            txtStatus = new TextBox
            {
                Location = new Point(15, 600),
                Size = new Size(930, 62),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(248, 250, 252),
                ForeColor = Color.FromArgb(71, 85, 105),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 9F, FontStyle.Regular),
                Text = $"{DateTime.Now:HH:mm:ss} 已启动系统架构图。可粘贴 JSON 后点击“导入JSON”或直接生成。"
            };

            pnlOptions = new Panel { Location = new Point(15, 675), Size = new Size(930, 92), BackColor = Color.White };
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
            Controls.Add(btnImportJson);
            Controls.Add(lblTitle);
            Controls.Add(txtTitle);
            Controls.Add(tabControl);
            Controls.Add(txtStatus);
            Controls.Add(pnlOptions);

            Resize += FormSystemArchitectureDiagram_Resize;
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
                Font = new Font(DefaultUiFontName, 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersWidth = 42,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                Font = new Font(DefaultUiFontName, 9F, FontStyle.Regular)
            };
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 41, 55);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font(DefaultUiFontName, 9F, FontStyle.Bold);
            grid.EnableHeadersVisualStyles = false;
            return grid;
        }

        private void BuildNodeGridColumns()
        {
            gridNodes.Columns.Clear();
            AddTextColumn(gridNodes, "id", "ID", 95);
            AddTextColumn(gridNodes, "parentId", "父级ID", 95);
            AddComboColumn(gridNodes, "kind", "类型", 80, new[] { "layer", "group", "sublayer", "module" });
            AddTextColumn(gridNodes, "text", "文本", 150);
            AddComboColumn(gridNodes, "shape", "形状", 82, new[] { "container", "module", "database" });
            AddTextColumn(gridNodes, "order", "顺序", 55);
            AddTextColumn(gridNodes, "row", "行", 45);
            AddTextColumn(gridNodes, "col", "列", 45);
            AddTextColumn(gridNodes, "colSpan", "跨列", 55);
            AddTextColumn(gridNodes, "width", "宽(in)", 62);
            AddTextColumn(gridNodes, "height", "高(in)", 62);
            AddTextColumn(gridNodes, "color", "颜色", 82);
        }

        private void BuildConnectionGridColumns()
        {
            gridConnections.Columns.Clear();
            AddTextColumn(gridConnections, "from", "From", 130);
            AddTextColumn(gridConnections, "to", "To", 130);
            AddComboColumn(gridConnections, "type", "类型", 90, new[] { "bidirectional", "arrow" });
            AddTextColumn(gridConnections, "xRatio", "X比例", 70);
        }

        private void AddTextColumn(DataGridView grid, string name, string header, int width)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                Width = width,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        }

        private void AddComboColumn(DataGridView grid, string name, string header, int width, string[] items)
        {
            var col = new DataGridViewComboBoxColumn
            {
                Name = name,
                HeaderText = header,
                Width = width,
                FlatStyle = FlatStyle.Flat,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            col.Items.AddRange(items);
            grid.Columns.Add(col);
        }

        private void BuildOptionsPanel()
        {
            var row = new FlowLayoutPanel
            {
                Location = new Point(15, 15),
                Size = new Size(620, 32),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent
            };

            txtLayerSpacing = new TextBox { Text = DefaultLayerSpacingMm.ToString("0.#", CultureInfo.InvariantCulture), Size = new Size(55, 25), TextAlign = HorizontalAlignment.Center };
            txtModuleSpacing = new TextBox { Text = DefaultModuleSpacingMm.ToString("0.#", CultureInfo.InvariantCulture), Size = new Size(55, 25), TextAlign = HorizontalAlignment.Center };
            chkCompactLayout = new CheckBox { Text = "紧凑排版", Size = new Size(95, 24), AutoSize = false, Checked = DefaultCompactLayoutChecked, Margin = new Padding(18, 2, 0, 0) };

            row.Controls.Add(CreateLabeledOption("层间距(mm):", txtLayerSpacing));
            row.Controls.Add(CreateLabeledOption("模块间距(mm):", txtModuleSpacing));
            row.Controls.Add(chkCompactLayout);

            var row2 = new FlowLayoutPanel
            {
                Location = new Point(15, 53),
                Size = new Size(620, 32),
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = Color.Transparent
            };
            cmbFontName = new ComboBox { Size = new Size(120, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontName.Items.AddRange(new[] { "宋体", "微软雅黑", "黑体", "楷体", "仿宋" });
            cmbFontName.SelectedItem = DefaultDrawingFontName;

            cmbFontSize = new ComboBox { Size = new Size(85, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFontSize.Items.AddRange(new[] { "三号", "小三", "四号", "小四", "五号", "小五" });
            cmbFontSize.SelectedItem = "五号";
            row2.Controls.Add(CreateLabeledOption("字体:", cmbFontName));
            row2.Controls.Add(CreateLabeledOption("字号:", cmbFontSize));

            btnGenerate = CreatePrimaryButton("生成架构图", new Point(680, 32), new Size(120, 36), Color.FromArgb(40, 167, 69));
            btnGenerate.Click += btnGenerate_Click;
            btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(815, 32),
                Size = new Size(90, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(233, 236, 239),
                ForeColor = Color.FromArgb(73, 80, 87),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += (s, e) => Close();

            pnlOptions.Controls.Add(row);
            pnlOptions.Controls.Add(row2);
            pnlOptions.Controls.Add(btnGenerate);
            pnlOptions.Controls.Add(btnClose);
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
            var label = new Label { Text = labelText, AutoSize = true, Margin = new Padding(0, 6, 6, 0) };
            editor.Margin = new Padding(0, 2, 0, 0);
            panel.Controls.Add(label);
            panel.Controls.Add(editor);
            return panel;
        }

        private void LayoutControls()
        {
            int margin = 15;
            int width = Math.Max(850, ClientSize.Width - margin * 2);
            tabControl.Width = width;
            pnlOptions.Width = width;
            pnlOptions.Top = ClientSize.Height - pnlOptions.Height - margin;
            txtStatus.Width = width;
            txtStatus.Top = pnlOptions.Top - txtStatus.Height - 12;
            tabControl.Height = Math.Max(260, txtStatus.Top - tabControl.Top - 12);
            btnClose.Left = pnlOptions.Width - btnClose.Width - 20;
            btnGenerate.Left = btnClose.Left - btnGenerate.Width - 15;
            txtTitle.Width = Math.Max(260, width - txtTitle.Left + margin);
        }

        private void FormSystemArchitectureDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void AppendLog(string message)
        {
            txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now:HH:mm:ss} {message}");
        }

        private void LoadSampleData()
        {
            txtTitle.Text = "图书管理系统架构图";
            txtJson.Text =
@"{
  ""title"": ""图书管理系统架构图"",
  ""layers"": [
    {
      ""id"": ""user"",
      ""name"": ""用户层"",
      ""items"": [
        { ""id"": ""frontend"", ""text"": ""Vue3 + CSS3"", ""type"": ""module"", ""span"": ""full"", ""color"": ""#c9ffc9"" }
      ]
    },
    {
      ""id"": ""backend"",
      ""name"": ""Django REST Framework"",
      ""style"": ""group"",
      ""children"": [
        {
          ""id"": ""view"",
          ""name"": ""表示层"",
          ""color"": ""#8fdada"",
          ""items"": [
            { ""id"": ""cat_view"", ""text"": ""分类管理"" },
            { ""id"": ""book_view"", ""text"": ""图书管理"" },
            { ""id"": ""user_view"", ""text"": ""用户管理"" },
            { ""id"": ""fav_view"", ""text"": ""图书收藏"" },
            { ""id"": ""rate_view"", ""text"": ""图书评分"" },
            { ""id"": ""recommend_view"", ""text"": ""图书推荐"" }
          ]
        },
        {
          ""id"": ""biz"",
          ""name"": ""业务层"",
          ""color"": ""#e4ce26"",
          ""items"": [
            { ""id"": ""cat_biz"", ""text"": ""分类管理"" },
            { ""id"": ""book_biz"", ""text"": ""图书管理"" },
            { ""id"": ""user_biz"", ""text"": ""用户管理"" },
            { ""id"": ""fav_biz"", ""text"": ""图书收藏"" },
            { ""id"": ""rate_biz"", ""text"": ""图书评分"" },
            { ""id"": ""recommend_biz"", ""text"": ""图书推荐"" }
          ]
        },
        {
          ""id"": ""dao"",
          ""name"": ""数据访问层"",
          ""color"": ""#78d63c"",
          ""items"": [
            { ""id"": ""orm"", ""text"": ""Django ORM"", ""type"": ""module"", ""span"": ""wide"" }
          ]
        }
      ]
    },
    {
      ""id"": ""database"",
      ""name"": ""数据层"",
      ""items"": [
        { ""id"": ""mysql"", ""text"": ""MySQL"", ""type"": ""database"" }
      ]
    },
    {
      ""id"": ""os"",
      ""name"": ""操作系统"",
      ""items"": [
        { ""id"": ""windows"", ""text"": ""Windows"", ""type"": ""module"", ""span"": ""wide"", ""color"": ""#c9ffc9"" }
      ]
    }
  ],
  ""connections"": [
    { ""from"": ""user"", ""to"": ""backend"", ""type"": ""bidirectional"" },
    { ""from"": ""view"", ""to"": ""biz"", ""type"": ""bidirectional"" },
    { ""from"": ""biz"", ""to"": ""dao"", ""type"": ""bidirectional"" },
    { ""from"": ""backend"", ""to"": ""database"", ""type"": ""bidirectional"" }
  ]
}";
            LoadJsonToTables();
        }

        private void btnAiPrompt_Click(object sender, EventArgs e)
        {
            string prompt =
@"请根据下面的系统说明，输出系统架构图 JSON，只输出 JSON，不要解释。
要求：
1. layers 表示从上到下的架构层。
2. 每层可以包含 items，也可以包含 children 子层。
3. 模块 type 只能是 module、database、service、external。
4. span 只能是 normal、wide、full。
5. connections 只描述层或子层之间的关系，不要输出坐标。
6. 不要生成 x、y、width、height。

系统说明：
";
            Clipboard.SetText(prompt);
            AppendLog("已复制 AI 生成系统架构 JSON 的提示词到剪贴板。");
        }

        private void btnImportJson_Click(object sender, EventArgs e)
        {
            if (!Clipboard.ContainsText())
            {
                MessageBox.Show("剪贴板中没有文本内容。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            txtJson.Text = Clipboard.GetText().Trim();
            LoadJsonToTables();
            AppendLog("已从剪贴板导入 JSON 文本。");
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                RenderArchitectureDiagram();
                AppendLog("系统架构图生成成功。");
                Close();
            }
            catch (Exception ex)
            {
                AppendLog("生成失败: " + ex.Message);
                MessageBox.Show("生成失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadJsonToTables()
        {
            ArchitectureDiagram diagram = ParseDiagram(txtJson.Text);
            gridNodes.Rows.Clear();
            gridConnections.Rows.Clear();

            int order = 1;
            foreach (ArchitectureLayer layer in diagram.layers)
            {
                AddLayerRows(layer, string.Empty, ref order);
            }

            foreach (ArchitectureConnection conn in diagram.connections)
            {
                int rowIndex = gridConnections.Rows.Add();
                DataGridViewRow row = gridConnections.Rows[rowIndex];
                row.Cells["from"].Value = conn.from;
                row.Cells["to"].Value = conn.to;
                row.Cells["type"].Value = string.IsNullOrWhiteSpace(conn.type) ? "bidirectional" : conn.type;
                row.Cells["xRatio"].Value = string.IsNullOrWhiteSpace(conn.xRatio) ? "0.5" : conn.xRatio;
            }
            tabControl.SelectedTab = tabNodes;
        }

        private void AddLayerRows(ArchitectureLayer layer, string parentId, ref int order)
        {
            bool hasChildren = layer.children != null && layer.children.Count > 0;
            string kind = hasChildren ? "group" : (string.IsNullOrWhiteSpace(parentId) ? "layer" : "sublayer");
            AddNodeRow(layer.id, parentId, kind, layer.name, "container", order++, 0, 0, 1, string.Empty, string.Empty, layer.color);

            int childOrder = 1;
            foreach (ArchitectureLayer child in layer.children ?? new List<ArchitectureLayer>())
            {
                AddLayerRows(child, layer.id, ref childOrder);
            }

            int itemIndex = 0;
            foreach (ArchitectureItem item in layer.items ?? new List<ArchitectureItem>())
            {
                int columns = Math.Min(4, Math.Max(1, layer.items.Count));
                int row = itemIndex / columns + 1;
                int col = itemIndex % columns + 1;
                int colSpan = GetSpanColumns(item.span, columns);
                string shape = item.type == "database" ? "database" : "module";
                string width = item.type == "database" ? DatabaseWidthInch.ToString("0.##", CultureInfo.InvariantCulture) : string.Empty;
                string height = item.type == "database" ? DatabaseHeightInch.ToString("0.##", CultureInfo.InvariantCulture) : string.Empty;
                AddNodeRow(item.id, layer.id, "module", item.text, shape, itemIndex + 1, row, col, colSpan, width, height, item.color);
                itemIndex++;
            }
        }

        private void AddNodeRow(string id, string parentId, string kind, string text, string shape, int order, int row, int col, int colSpan, string width, string height, string color)
        {
            int rowIndex = gridNodes.Rows.Add();
            DataGridViewRow gridRow = gridNodes.Rows[rowIndex];
            gridRow.Cells["id"].Value = id;
            gridRow.Cells["parentId"].Value = parentId;
            gridRow.Cells["kind"].Value = kind;
            gridRow.Cells["text"].Value = text;
            gridRow.Cells["shape"].Value = shape;
            gridRow.Cells["order"].Value = order;
            gridRow.Cells["row"].Value = row > 0 ? row.ToString(CultureInfo.InvariantCulture) : string.Empty;
            gridRow.Cells["col"].Value = col > 0 ? col.ToString(CultureInfo.InvariantCulture) : string.Empty;
            gridRow.Cells["colSpan"].Value = colSpan.ToString(CultureInfo.InvariantCulture);
            gridRow.Cells["width"].Value = width;
            gridRow.Cells["height"].Value = height;
            gridRow.Cells["color"].Value = color;
        }

        private int GetSpanColumns(string span, int columns)
        {
            if (span == "full") return columns;
            if (span == "wide") return Math.Min(columns, 2);
            return 1;
        }

        private List<TableNode> ReadNodesFromGrid()
        {
            var nodes = new List<TableNode>();
            foreach (DataGridViewRow row in gridNodes.Rows)
            {
                if (row.IsNewRow) continue;
                string id = GetCellText(row, "id");
                if (string.IsNullOrWhiteSpace(id)) continue;
                nodes.Add(new TableNode
                {
                    Id = id,
                    ParentId = GetCellText(row, "parentId"),
                    Kind = WithDefault(GetCellText(row, "kind"), "module"),
                    Text = WithDefault(GetCellText(row, "text"), id),
                    Shape = WithDefault(GetCellText(row, "shape"), "module"),
                    Order = ParseInt(GetCellText(row, "order"), nodes.Count + 1),
                    Row = ParseInt(GetCellText(row, "row"), 0),
                    Col = ParseInt(GetCellText(row, "col"), 0),
                    ColSpan = Math.Max(1, ParseInt(GetCellText(row, "colSpan"), 1)),
                    Width = ParseOptionalDouble(GetCellText(row, "width")),
                    Height = ParseOptionalDouble(GetCellText(row, "height")),
                    Color = GetCellText(row, "color")
                });
            }

            if (nodes.Count == 0)
            {
                throw new ArgumentException("节点表不能为空。");
            }
            return nodes;
        }

        private List<TableConnection> ReadConnectionsFromGrid()
        {
            var connections = new List<TableConnection>();
            foreach (DataGridViewRow row in gridConnections.Rows)
            {
                if (row.IsNewRow) continue;
                string from = GetCellText(row, "from");
                string to = GetCellText(row, "to");
                if (string.IsNullOrWhiteSpace(from) || string.IsNullOrWhiteSpace(to)) continue;
                connections.Add(new TableConnection
                {
                    From = from,
                    To = to,
                    Type = WithDefault(GetCellText(row, "type"), "bidirectional"),
                    XRatio = ParseOptionalDouble(GetCellText(row, "xRatio"))
                });
            }
            return connections;
        }

        private string GetCellText(DataGridViewRow row, string name)
        {
            object value = row.Cells[name].Value;
            return value == null ? string.Empty : value.ToString().Trim();
        }

        private string WithDefault(string value, string defaultValue)
        {
            return string.IsNullOrWhiteSpace(value) ? defaultValue : value;
        }

        private int ParseInt(string value, int defaultValue)
        {
            int parsed;
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : defaultValue;
        }

        private double? ParseOptionalDouble(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            double parsed;
            string normalized = value.Trim().Replace('，', '.');
            if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) ||
                double.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
            {
                return parsed;
            }
            return null;
        }

        private void RenderArchitectureDiagram()
        {
            List<TableNode> nodes = ReadNodesFromGrid();
            List<TableConnection> connections = ReadConnectionsFromGrid();
            double layerSpacing = ParsePositiveNumber(txtLayerSpacing.Text, "层间距") / 25.4;
            double moduleSpacing = ParsePositiveNumber(txtModuleSpacing.Text, "模块间距") / 25.4;
            if (chkCompactLayout.Checked)
            {
                layerSpacing = Math.Min(layerSpacing, 0.24);
                moduleSpacing = Math.Min(moduleSpacing, 0.18);
            }

            string fontName = GetSelectedFontName();
            double fontSizePt = GetSelectedFontSizePt();
            Visio.Application app = Globals.ThisAddIn.Application;
            Visio.Page page = GetOrCreateActivePage(app);
            ClearPageShapes(page);

            double pageWidth = DefaultPageWidthInch;
            double contentWidth = pageWidth - DefaultLeftMarginInch * 2;
            double yTop = CalculateTotalHeight(nodes, contentWidth, moduleSpacing, layerSpacing) + DefaultTopMarginInch;
            page.PageSheet.CellsU["PageWidth"].FormulaU = $"{pageWidth.ToString(CultureInfo.InvariantCulture)} in";
            page.PageSheet.CellsU["PageHeight"].FormulaU = $"{(yTop + DefaultTopMarginInch).ToString(CultureInfo.InvariantCulture)} in";

            var anchors = new Dictionary<string, RectangleD>();
            double currentTop = yTop;
            foreach (TableNode layer in GetTopNodes(nodes))
            {
                double height = CalculateLayerHeight(layer, nodes, contentWidth, moduleSpacing);
                DrawTopLayer(page, layer, nodes, DefaultLeftMarginInch, currentTop - height, contentWidth, height, moduleSpacing, fontName, fontSizePt, anchors);
                currentTop -= height + layerSpacing;
            }

            DrawConnections(page, connections, anchors);
        }

        private ArchitectureDiagram ParseDiagram(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentException("JSON 不能为空。");
            }
            var serializer = new JavaScriptSerializer();
            ArchitectureDiagram diagram = serializer.Deserialize<ArchitectureDiagram>(json);
            if (diagram == null || diagram.layers == null || diagram.layers.Count == 0)
            {
                throw new ArgumentException("JSON 中没有 layers。");
            }
            if (diagram.connections == null)
            {
                diagram.connections = new List<ArchitectureConnection>();
            }
            NormalizeLayers(diagram.layers);
            return diagram;
        }

        private void NormalizeLayers(List<ArchitectureLayer> layers)
        {
            foreach (ArchitectureLayer layer in layers)
            {
                if (string.IsNullOrWhiteSpace(layer.id)) layer.id = Guid.NewGuid().ToString("N");
                if (string.IsNullOrWhiteSpace(layer.name)) layer.name = layer.id;
                if (layer.items == null) layer.items = new List<ArchitectureItem>();
                if (layer.children == null) layer.children = new List<ArchitectureLayer>();
                foreach (ArchitectureItem item in layer.items)
                {
                    if (string.IsNullOrWhiteSpace(item.id)) item.id = Guid.NewGuid().ToString("N");
                    if (string.IsNullOrWhiteSpace(item.text)) item.text = item.id;
                    if (string.IsNullOrWhiteSpace(item.type)) item.type = "module";
                    if (string.IsNullOrWhiteSpace(item.span)) item.span = "normal";
                }
                NormalizeLayers(layer.children);
            }
        }

        private IEnumerable<TableNode> GetTopNodes(List<TableNode> nodes)
        {
            return nodes
                .Where(n => n.Kind != "module" && string.IsNullOrWhiteSpace(n.ParentId))
                .OrderBy(n => n.Order);
        }

        private List<TableNode> GetChildren(List<TableNode> nodes, string parentId, string kind = null)
        {
            return nodes
                .Where(n => string.Equals(n.ParentId, parentId, StringComparison.OrdinalIgnoreCase) &&
                    (kind == null || string.Equals(n.Kind, kind, StringComparison.OrdinalIgnoreCase)))
                .OrderBy(n => n.Order)
                .ToList();
        }

        private double CalculateTotalHeight(List<TableNode> nodes, double width, double moduleSpacing, double layerSpacing)
        {
            double total = 0;
            foreach (TableNode layer in GetTopNodes(nodes))
            {
                total += CalculateLayerHeight(layer, nodes, width, moduleSpacing) + layerSpacing;
            }
            return Math.Max(4.0, total);
        }

        private double CalculateLayerHeight(TableNode layer, List<TableNode> nodes, double width, double moduleSpacing)
        {
            List<TableNode> childLayers = GetChildren(nodes, layer.Id).Where(n => n.Kind != "module").ToList();
            if (childLayers.Count > 0)
            {
                double childWidth = width - LayerInnerPaddingInch * 2.0;
                return childLayers.Sum(child => CalculateChildLayerHeight(child, nodes, childWidth, moduleSpacing))
                    + moduleSpacing * Math.Max(0, childLayers.Count - 1)
                    + LayerInnerPaddingInch * 2.0
                    + GroupTitleHeightInch;
            }
            return CalculateChildLayerHeight(layer, nodes, width, moduleSpacing);
        }

        private double CalculateChildLayerHeight(TableNode layer, List<TableNode> nodes, double width, double moduleSpacing)
        {
            List<TableNode> items = GetChildren(nodes, layer.Id, "module");
            int maxRow = Math.Max(1, items.Count == 0 ? 1 : items.Max(item => item.Row > 0 ? item.Row : 1));
            double itemH = Math.Max(ModuleHeightInch, items.Count == 0 ? ModuleHeightInch : items.Max(item => item.Height ?? ModuleHeightInch));
            return Math.Max(0.64, maxRow * itemH + Math.Max(0, maxRow - 1) * moduleSpacing + LayerInnerPaddingInch * 2.0);
        }

        private void DrawTopLayer(Visio.Page page, TableNode layer, List<TableNode> nodes, double left, double bottom, double width, double height, double moduleSpacing, string fontName, double fontSizePt, Dictionary<string, RectangleD> anchors)
        {
            Visio.Shape outer = page.DrawRectangle(left, bottom, left + width, bottom + height);
            ApplyContainerStyle(outer, "#ffffff", fontName, fontSizePt);
            anchors[layer.Id] = new RectangleD(left, bottom, width, height);

            List<TableNode> childLayers = GetChildren(nodes, layer.Id).Where(n => n.Kind != "module").ToList();
            if (childLayers.Count > 0)
            {
                DrawPlainText(page, layer.Text, left + 0.18, bottom + height - 0.28, left + 3.1, bottom + height - 0.07, fontName, fontSizePt, false);
                double childTop = bottom + height - GroupTitleHeightInch - LayerInnerPaddingInch;
                foreach (TableNode child in childLayers)
                {
                    double childHeight = CalculateChildLayerHeight(child, nodes, width - LayerInnerPaddingInch * 2.0, moduleSpacing);
                    DrawChildLayer(page, child, nodes, left + LayerInnerPaddingInch, childTop - childHeight, width - LayerInnerPaddingInch * 2.0, childHeight, moduleSpacing, fontName, fontSizePt, anchors, true);
                    childTop -= childHeight + moduleSpacing;
                }
            }
            else
            {
                DrawPlainText(page, layer.Text, left + 0.10, bottom + height / 2.0 - 0.17, left + LayerLabelWidthInch, bottom + height / 2.0 + 0.17, fontName, fontSizePt, true);
                anchors[layer.Id] = new RectangleD(left, bottom, width, height);
                DrawItems(page, GetChildren(nodes, layer.Id, "module"), left + LayerLabelWidthInch + 0.10, bottom + LayerInnerPaddingInch, width - LayerLabelWidthInch - 0.24, height - LayerInnerPaddingInch * 2.0, moduleSpacing, fontName, fontSizePt, anchors);
            }
        }

        private void DrawChildLayer(Visio.Page page, TableNode layer, List<TableNode> nodes, double left, double bottom, double width, double height, double moduleSpacing, string fontName, double fontSizePt, Dictionary<string, RectangleD> anchors, bool drawBackground)
        {
            string fill = string.IsNullOrWhiteSpace(layer.Color) ? "#ffffff" : SoftenColor(layer.Color);
            if (drawBackground)
            {
                Visio.Shape bg = page.DrawRectangle(left, bottom, left + width, bottom + height);
                ApplyContainerStyle(bg, fill, fontName, fontSizePt);
            }
            anchors[layer.Id] = new RectangleD(left, bottom, width, height);

            DrawPlainText(page, layer.Text, left + 0.06, bottom + height / 2.0 - 0.22, left + LayerLabelWidthInch - 0.04, bottom + height / 2.0 + 0.22, fontName, fontSizePt, true);

            double contentLeft = left + LayerLabelWidthInch + 0.10;
            double contentWidth = width - LayerLabelWidthInch - 0.24;
            DrawItems(page, GetChildren(nodes, layer.Id, "module"), contentLeft, bottom + LayerInnerPaddingInch, contentWidth, height - LayerInnerPaddingInch * 2.0, moduleSpacing, fontName, fontSizePt, anchors);
        }

        private void DrawItems(Visio.Page page, List<TableNode> items, double left, double bottom, double width, double height, double spacing, string fontName, double fontSizePt, Dictionary<string, RectangleD> anchors)
        {
            if (items == null || items.Count == 0) return;
            int columns = Math.Max(1, items.Max(item => Math.Max(item.Col, 1) + Math.Max(item.ColSpan, 1) - 1));
            double defaultItemH = Math.Max(ModuleHeightInch, items.Max(item => item.Height ?? ModuleHeightInch));
            double normalW = (width - spacing * (columns - 1)) / columns;
            foreach (TableNode item in items.OrderBy(item => item.Row <= 0 ? 1 : item.Row).ThenBy(item => item.Col <= 0 ? item.Order : item.Col).ThenBy(item => item.Order))
            {
                int row = Math.Max(1, item.Row) - 1;
                int col = Math.Max(1, item.Col) - 1;
                int colSpan = Math.Max(1, Math.Min(item.ColSpan, columns - col));
                double itemH = item.Height ?? defaultItemH;
                double itemW = item.Width ?? (normalW * colSpan + spacing * (colSpan - 1));
                if (item.Shape == "database" && !item.Width.HasValue) itemW = Math.Min(DatabaseWidthInch, width * 0.30);
                double cellLeft = left + col * (normalW + spacing);
                double cellWidth = normalW * colSpan + spacing * (colSpan - 1);
                double x = cellLeft + (cellWidth - itemW) / 2.0;
                double yTop = bottom + height - row * (defaultItemH + spacing);
                double yBottom = yTop - itemH;
                Visio.Shape shape = item.Shape == "database"
                    ? DrawDatabase(page, item.Text, x, yBottom - 0.03, itemW, item.Height ?? DatabaseHeightInch, fontName, fontSizePt)
                    : page.DrawRectangle(x, yBottom, x + itemW, yTop);
                if (item.Shape != "database")
                {
                    shape.Text = item.Text;
                    ApplyModuleStyle(shape, string.IsNullOrWhiteSpace(item.Color) ? "#ffffff" : SoftenColor(item.Color), fontName, fontSizePt);
                }
                anchors[item.Id] = new RectangleD(x, yBottom, itemW, itemH);
            }
        }

        private Visio.Shape DrawDatabase(Visio.Page page, string text, double left, double bottom, double width, double height, string fontName, double fontSizePt)
        {
            Visio.Shape body = page.DrawRectangle(left, bottom, left + width, bottom + height);
            body.Text = text;
            ApplyModuleStyle(body, "#DDEFF6", fontName, fontSizePt);
            TrySetFormula(body, "LineColor", "RGB(55, 105, 135)");
            TrySetFormula(body, "LineWeight", "0.65pt");
            Visio.Shape top = page.DrawOval(left, bottom + height - 0.09, left + width, bottom + height + 0.09);
            ApplyLineOnlyStyle(top, "RGB(55, 105, 135)");
            return body;
        }

        private void DrawConnections(Visio.Page page, List<TableConnection> connections, Dictionary<string, RectangleD> anchors)
        {
            if (connections == null) return;
            foreach (TableConnection conn in connections)
            {
                if (!anchors.ContainsKey(conn.From) || !anchors.ContainsKey(conn.To)) continue;
                RectangleD from = anchors[conn.From];
                RectangleD to = anchors[conn.To];
                double x = GetConnectorX(from, to, conn);
                double y1 = from.Bottom < to.Bottom ? from.Top : from.Bottom;
                double y2 = from.Bottom < to.Bottom ? to.Bottom : to.Top;
                if (Math.Abs(y2 - y1) > 0.08)
                {
                    double trim = 0.025;
                    if (y1 < y2)
                    {
                        y1 += trim;
                        y2 -= trim;
                    }
                    else
                    {
                        y1 -= trim;
                        y2 += trim;
                    }
                }
                Visio.Shape line = page.DrawLine(x, y1, x, y2);
                ApplyConnectorStyle(line, conn.Type == "bidirectional");
            }
        }

        private void ApplyContainerStyle(Visio.Shape shape, string fillHex, string fontName, double fontSizePt)
        {
            TrySetFormula(shape, "FillForegnd", ToRgbFormula(fillHex));
            TrySetFormula(shape, "FillPattern", "1");
            TrySetFormula(shape, "LineColor", AcademicBorderColor);
            TrySetFormula(shape, "LineWeight", "0.65pt");
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSizePt.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(26, 32, 44)");
        }

        private void ApplyModuleStyle(Visio.Shape shape, string fillHex, string fontName, double fontSizePt)
        {
            TrySetFormula(shape, "FillForegnd", ToRgbFormula(fillHex));
            TrySetFormula(shape, "FillPattern", "1");
            TrySetFormula(shape, "LineColor", AcademicBorderColor);
            TrySetFormula(shape, "LineWeight", "0.6pt");
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSizePt.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(22, 28, 38)");
            TrySetFormula(shape, "Para.HorzAlign", "1");
            TrySetFormula(shape, "VerticalAlign", "1");
        }

        private void ApplyLineOnlyStyle(Visio.Shape shape, string lineColor)
        {
            TrySetFormula(shape, "FillPattern", "0");
            TrySetFormula(shape, "LineColor", lineColor);
            TrySetFormula(shape, "LineWeight", "0.65pt");
        }

        private void ApplyConnectorStyle(Visio.Shape shape, bool bidirectional)
        {
            TrySetFormula(shape, "LineColor", AcademicConnectorColor);
            TrySetFormula(shape, "LineWeight", "0.65pt");
            TrySetFormula(shape, "BeginArrow", bidirectional ? "4" : "0");
            TrySetFormula(shape, "EndArrow", "4");
        }

        private void DrawPlainText(Visio.Page page, string text, double left, double bottom, double right, double top, string fontName, double fontSizePt, bool verticalLabel)
        {
            Visio.Shape shape = page.DrawRectangle(left, bottom, right, top);
            shape.Text = verticalLabel ? FormatVerticalLabel(text) : (text ?? string.Empty);
            TrySetFormula(shape, "FillPattern", "0");
            TrySetFormula(shape, "LinePattern", "0");
            TrySetFormula(shape, "Char.Font", $"\"{fontName}\"");
            TrySetFormula(shape, "Char.Size", $"{fontSizePt.ToString(CultureInfo.InvariantCulture)}pt");
            TrySetFormula(shape, "Char.Color", "RGB(22, 28, 38)");
            TrySetFormula(shape, "Para.HorzAlign", "1");
            TrySetFormula(shape, "VerticalAlign", "1");
        }

        private string FormatVerticalLabel(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return string.Empty;
            if (text.Length <= 3) return text;
            return string.Join(Environment.NewLine, text.ToCharArray());
        }

        private string ToRgbFormula(string hex)
        {
            Color color = ColorTranslator.FromHtml(string.IsNullOrWhiteSpace(hex) ? "#ffffff" : hex);
            return $"RGB({color.R}, {color.G}, {color.B})";
        }

        private string SoftenColor(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return "#ffffff";
            try
            {
                Color color = ColorTranslator.FromHtml(hex);
                const double whiteMix = 0.58;
                int r = (int)Math.Round(color.R + (255 - color.R) * whiteMix);
                int g = (int)Math.Round(color.G + (255 - color.G) * whiteMix);
                int b = (int)Math.Round(color.B + (255 - color.B) * whiteMix);
                return $"#{r:X2}{g:X2}{b:X2}";
            }
            catch
            {
                return "#ffffff";
            }
        }

        private double GetConnectorX(RectangleD from, RectangleD to, TableConnection conn)
        {
            double left = Math.Max(from.Left, to.Left);
            double right = Math.Min(from.Right, to.Right);
            if (right <= left)
            {
                return (from.CenterX + to.CenterX) / 2.0;
            }

            double ratio = Math.Max(0, Math.Min(1, conn.XRatio ?? 0.5));
            return left + (right - left) * ratio;
        }

        private string GetSelectedFontName()
        {
            return cmbFontName == null || cmbFontName.SelectedItem == null ? DefaultDrawingFontName : cmbFontName.SelectedItem.ToString();
        }

        private double GetSelectedFontSizePt()
        {
            string selected = cmbFontSize == null || cmbFontSize.SelectedItem == null ? "五号" : cmbFontSize.SelectedItem.ToString();
            switch (selected)
            {
                case "三号": return 16.0;
                case "小三": return 15.0;
                case "四号": return 14.0;
                case "小四": return 12.0;
                case "小五": return 9.0;
                case "五号":
                default: return DefaultDrawingFontSizePt;
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

        private void TrySetFormula(Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                shape.CellsU[cellName].FormulaU = formula;
            }
            catch { }
        }

        private class RectangleD
        {
            public RectangleD(double left, double bottom, double width, double height)
            {
                Left = left;
                Bottom = bottom;
                Width = width;
                Height = height;
            }
            public double Left { get; private set; }
            public double Bottom { get; private set; }
            public double Width { get; private set; }
            public double Height { get; private set; }
            public double Right => Left + Width;
            public double Top => Bottom + Height;
            public double CenterX => Left + Width / 2.0;
        }

        private class TableNode
        {
            public string Id { get; set; }
            public string ParentId { get; set; }
            public string Kind { get; set; }
            public string Text { get; set; }
            public string Shape { get; set; }
            public int Order { get; set; }
            public int Row { get; set; }
            public int Col { get; set; }
            public int ColSpan { get; set; }
            public double? Width { get; set; }
            public double? Height { get; set; }
            public string Color { get; set; }
        }

        private class TableConnection
        {
            public string From { get; set; }
            public string To { get; set; }
            public string Type { get; set; }
            public double? XRatio { get; set; }
        }

        private class ArchitectureDiagram
        {
            public string title { get; set; }
            public List<ArchitectureLayer> layers { get; set; }
            public List<ArchitectureConnection> connections { get; set; }
        }

        private class ArchitectureLayer
        {
            public string id { get; set; }
            public string name { get; set; }
            public string style { get; set; }
            public string color { get; set; }
            public List<ArchitectureItem> items { get; set; }
            public List<ArchitectureLayer> children { get; set; }
        }

        private class ArchitectureItem
        {
            public string id { get; set; }
            public string text { get; set; }
            public string type { get; set; }
            public string span { get; set; }
            public string color { get; set; }
        }

        private class ArchitectureConnection
        {
            public string from { get; set; }
            public string to { get; set; }
            public string type { get; set; }
            public string xRatio { get; set; }
        }
    }
}
