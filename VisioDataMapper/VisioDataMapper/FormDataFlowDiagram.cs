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
    public class FormDataFlowDiagram : Form
    {
        // UI 控件
        private Button btnImportTable;
        private Label lblTitle;
        private TextBox txtTitle;
        private DataGridView dgvData;
        private TextBox txtStatus;

        private Label lblEntityStyle;
        private ComboBox cmbEntityStyle;
        private Label lblProcessStyle;
        private ComboBox cmbProcessStyle;
        private Label lblStoreStyle;
        private ComboBox cmbStoreStyle;
        private Label lblConnectorStyle;
        private ComboBox cmbConnectorStyle;

        private Label lblSpacing;
        private TextBox txtHorSpacing;
        private TextBox txtVerSpacing;
        private Label lblSpacingUnit;

        private Button btnGenerate;
        private Button btnClose;

        private Label lblLayoutScheme;
        private ComboBox cmbLayoutScheme;

        private string currentFontName = "宋体";
        private double currentFontSizePt = 10.5;
        private readonly DataFlowDiagramLevel diagramLevel;

        public enum DataFlowDiagramLevel
        {
            Top,
            Level1,
            Level2
        }

        // 默认内置的 Graphviz DOT 示例文本（即用户给出的系统通勤/管理数据）
        private const string DefaultDotSample = @"digraph AdminDFD {
    // 全局样式，线条清晰不交叉
    rankdir=LR;
    node [fontname=""SimHei"", fontsize=10];
    edge [fontname=""SimHei"", fontsize=9];

    // 1.外部实体：管理员（普通矩形）
    Admin [label=""管理员"", shape=box, style=filled, fillcolor=""#e8f4ff""];

    // 2.所有加工：圆角矩形
    P01 [label=""P0.1 用户信息管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];
    P02 [label=""P0.2 图书分类管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];
    P03 [label=""P0.3 公告分类管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];
    P04 [label=""P0.4 系统日志管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];
    P05 [label=""P0.5 系统参数管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];
    P06 [label=""P0.6 全流程业务管理"", shape=roundedbox, style=filled, fillcolor=""#d7f5e8""];

    // 3.数据存储：record 双线数据表（DFD标准存储样式，独立分开，不打包）
    D1 [label=""D1\n用户信息表"", shape=record, style=filled, fillcolor=""#fff2cc""];
    D2 [label=""D2\n图书分类表"", shape=record, style=filled, fillcolor=""#fff2cc""];
    D3 [label=""D3\n公告分类表"", shape=record, style=filled, fillcolor=""#fff2cc""];
    D4 [label=""D4\n系统日志表"", shape=record, style=filled, fillcolor=""#fff2cc""];
    D5 [label=""D5\n系统参数表"", shape=record, style=filled, fillcolor=""#fff2cc""];
    D6 [label=""D6\n图书信息表"", shape=record, style=filled, fillcolor=""#fff2cc""];

    // ========== 1.管理员 → 各加工 数据流 ==========
    Admin -> P01 [label=""用户维护表单""];
    Admin -> P02 [label=""图书分类操作指令""];
    Admin -> P03 [label=""公告分类编辑表单""];
    Admin -> P04 [label=""日志查询筛选条件""];
    Admin -> P05 [label=""系统参数配置单""];
    Admin -> P06 [label=""图书业务操作指令""];

    // ========== 2.P0.1 用户信息管理 ↔ D1 用户信息表 ==========
    P01 -> D1 [label=""更新用户数据""];
    D1 -> P01 [label=""用户基础数据""];
    P01 -> P04 [label=""用户操作记录""];

    // ========== 3.P0.2 图书分类管理 ↔ D2、关联D6 ==========
    P02 -> D2 [label=""新增/修改分类""];
    D2 -> P02 [label=""分类基础数据""];
    P02 -> P04 [label=""分类操作记录""];
    D2 -> P06 [label=""分类编码信息""];
    P06 -> D2 [label=""图书分类关联数据""];

    // ========== 4.P0.3 公告分类管理 ↔ D3 ==========
    P03 -> D3 [label=""公告分类修改数据""];
    D3 -> P03 [label=""公告分类原始数据""];
    P03 -> P04 [label=""公告分类操作日志""];

    // ========== 5.P0.4 系统日志管理 ↔ D4 ==========
    P04 -> D4 [label=""新增操作日志记录""];
    D4 -> P04 [label=""历史日志数据集""];

    // ========== 6.P0.5 系统参数管理 ↔ D5 ==========
    P05 -> D5 [label=""修改后全局参数""];
    D5 -> P05 [label=""默认系统参数""];
    P05 -> P04 [label=""参数变更日志""];
    D5 -> P06 [label=""全局配置参数""];

    // ========== 7.P0.6 全流程业务管理 ↔ D6 ==========
    P06 -> D6 [label=""图书借还业务数据""];
    D6 -> P06 [label=""完整图书业务信息""];
    P06 -> P04 [label=""图书业务操作记录""];
}";

        private const string TopLevelDotSample = @"digraph LibraryTopDFD {
    rankdir=LR;
    node [fontname=""SimHei"", fontsize=10];
    edge [fontname=""SimHei"", fontsize=9];

    User [label=""读者"", shape=box];
    Admin [label=""管理员"", shape=box];
    Librarian [label=""图书管理员"", shape=box];
    Payment [label=""支付平台"", shape=box];

    P0 [label=""P0 图书借阅管理系统"", shape=roundedbox];

    D1 [label=""D1\n用户信息表"", shape=record];
    D2 [label=""D2\n图书信息表"", shape=record];
    D3 [label=""D3\n借阅记录表"", shape=record];
    D4 [label=""D4\n支付记录表"", shape=record];

    User -> P0 [label=""注册登录信息、查询条件、借阅申请""];
    P0 -> User [label=""图书检索结果、借阅状态、通知信息""];

    Admin -> P0 [label=""系统配置、用户管理指令""];
    P0 -> Admin [label=""统计报表、系统运行日志""];

    Librarian -> P0 [label=""图书上架、归还确认、库存维护""];
    P0 -> Librarian [label=""待处理借阅单、库存预警""];

    P0 -> Payment [label=""支付请求""];
    Payment -> P0 [label=""支付结果""];

    P0 -> D1 [label=""维护用户资料""];
    D1 -> P0 [label=""用户基础数据""];
    P0 -> D2 [label=""维护图书资料""];
    D2 -> P0 [label=""图书库存数据""];
    P0 -> D3 [label=""写入借阅记录""];
    D3 -> P0 [label=""历史借阅数据""];
    P0 -> D4 [label=""写入支付记录""];
    D4 -> P0 [label=""支付流水数据""];
}";

        public FormDataFlowDiagram() : this(DataFlowDiagramLevel.Level2)
        {
        }

        public FormDataFlowDiagram(DataFlowDiagramLevel level)
        {
            diagramLevel = level;
            InitializeComponent();
            LoadSampleData();
        }

        private void InitializeComponent()
        {
            this.Text = "智能画图-" + GetDiagramLevelName();
            this.Size = new Size(920, 750);
            this.MinimumSize = new Size(920, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = false;

            Font defaultFont = new Font("Microsoft YaHei", 9F, FontStyle.Regular);
            this.Font = defaultFont;
            this.BackColor = Color.FromArgb(240, 244, 248);

            // 头部提示与按钮
            Label lblTip = new Label
            {
                Text = $"请将您的{GetDiagramLevelName()}信息导入到表格，内容格式参考→ 双击此处，查看AI智能画图教程",
                Location = new Point(15, 15),
                AutoSize = true,
                ForeColor = Color.FromArgb(90, 107, 124),
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Regular)
            };

            btnImportTable = new Button
            {
                Text = "导入表格",
                Location = new Point(15, 45),
                Size = new Size(110, 32),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 122, 255),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnImportTable.FlatAppearance.BorderSize = 0;
            btnImportTable.Click += btnImportTable_Click;

            lblTitle = new Label
            {
                Text = "大标题:",
                Location = new Point(280, 50),
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104)
            };

            txtTitle = new TextBox
            {
                Location = new Point(340, 47),
                Size = new Size(400, 25),
                Font = new Font("Microsoft YaHei", 10F, FontStyle.Regular)
            };

            // 中部 DataGridView
            dgvData = new DataGridView
            {
                Location = new Point(15, 95),
                Size = new Size(875, 360),
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                GridColor = Color.FromArgb(230, 235, 240),
                CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 28 }
            };

            dgvData.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(240, 244, 248);
            dgvData.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(73, 80, 87);
            dgvData.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold);

            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "Index", HeaderText = "序号", FillWeight = 40 });
            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "Text", HeaderText = "形状文本", FillWeight = 140 });

            var shapeTypeCol = new DataGridViewComboBoxColumn { Name = "ShapeType", HeaderText = "使用形状", FillWeight = 100 };
            shapeTypeCol.Items.AddRange("实体", "加工", "数据存储");
            dgvData.Columns.Add(shapeTypeCol);

            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "Next", HeaderText = "下一步", FillWeight = 90 });
            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "LineText", HeaderText = "连接线文字", FillWeight = 160 });
            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "XPercent", HeaderText = "X(%)", FillWeight = 55 });
            dgvData.Columns.Add(new DataGridViewTextBoxColumn { Name = "YPercent", HeaderText = "Y(%)", FillWeight = 55 });

            // 日志文本框
            txtStatus = new TextBox
            {
                Location = new Point(15, 470),
                Size = new Size(875, 65),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.FromArgb(248, 250, 252),
                ForeColor = Color.FromArgb(100, 116, 139),
                BorderStyle = BorderStyle.None,
                Text = $"{DateTime.Now:HH:mm:ss} 初次启动{GetDiagramLevelName()}。建议通过剪贴板复制 DOT 文本，然后点击“导入表格”自动解析。"
            };

            // 底部配置面板
            Panel pnlOptions = new Panel
            {
                Location = new Point(15, 550),
                Size = new Size(875, 130),
                BackColor = Color.White
            };
            pnlOptions.Paint += (s, pe) =>
            {
                using (var pen = new Pen(Color.FromArgb(218, 224, 233), 1))
                {
                    pe.Graphics.DrawRectangle(pen, 0, 0, pnlOptions.Width - 1, pnlOptions.Height - 1);
                }
            };

            lblEntityStyle = new Label { Text = "外部实体:", Location = new Point(15, 20), AutoSize = true };
            cmbEntityStyle = new ComboBox { Location = new Point(80, 16), Size = new Size(100, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbEntityStyle.Items.AddRange(new string[] { "矩形" });
            cmbEntityStyle.SelectedIndex = 0;

            lblProcessStyle = new Label { Text = "数据处理:", Location = new Point(200, 20), AutoSize = true };
            cmbProcessStyle = new ComboBox { Location = new Point(265, 16), Size = new Size(110, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbProcessStyle.Items.AddRange(new string[] { "正方形-带标", "圆角矩形" });
            cmbProcessStyle.SelectedIndex = 0;

            lblStoreStyle = new Label { Text = "数据存储:", Location = new Point(400, 20), AutoSize = true };
            cmbStoreStyle = new ComboBox { Location = new Point(465, 16), Size = new Size(140, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbStoreStyle.Items.AddRange(new string[] { "小正方形+两横线", "三边矩形(右开口)" });
            cmbStoreStyle.SelectedIndex = 0;

            lblLayoutScheme = new Label { Text = "排版方案:", Location = new Point(620, 20), AutoSize = true };
            cmbLayoutScheme = new ComboBox { Location = new Point(690, 16), Size = new Size(160, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLayoutScheme.Items.AddRange(new string[] { "方案 A (Visio避让)", "方案 B (拓扑分层布局)", "方案 C (DFD语义分区)", "方案 E (智能语义拓扑布局)" });
            cmbLayoutScheme.SelectedIndex = diagramLevel == DataFlowDiagramLevel.Top ? 2 : diagramLevel == DataFlowDiagramLevel.Level2 ? 3 : 0;

            lblConnectorStyle = new Label { Text = "连接线:", Location = new Point(15, 65), AutoSize = true };
            cmbConnectorStyle = new ComboBox { Location = new Point(80, 61), Size = new Size(140, 25), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbConnectorStyle.Items.AddRange(new string[] { "绑定形状", "绑定新增连接点", "普通动态连线" });
            cmbConnectorStyle.SelectedIndex = 0;

            lblSpacing = new Label { Text = "横纵向间距:", Location = new Point(245, 65), AutoSize = true };
            txtHorSpacing = new TextBox { Text = "30", Location = new Point(325, 61), Size = new Size(40, 25), TextAlign = HorizontalAlignment.Center };
            txtVerSpacing = new TextBox { Text = "35", Location = new Point(375, 61), Size = new Size(40, 25), TextAlign = HorizontalAlignment.Center };
            lblSpacingUnit = new Label { Text = "mm", Location = new Point(420, 65), AutoSize = true };

            btnGenerate = new Button
            {
                Text = "生成绘图",
                Location = new Point(620, 45),
                Size = new Size(110, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(40, 167, 69),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnGenerate.FlatAppearance.BorderSize = 0;
            btnGenerate.Click += btnGenerate_Click;

            btnClose = new Button
            {
                Text = "关闭",
                Location = new Point(745, 45),
                Size = new Size(100, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(233, 236, 239),
                ForeColor = Color.FromArgb(73, 80, 87),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.Click += btnClose_Click;

            pnlOptions.Controls.Add(lblEntityStyle);
            pnlOptions.Controls.Add(cmbEntityStyle);
            pnlOptions.Controls.Add(lblProcessStyle);
            pnlOptions.Controls.Add(cmbProcessStyle);
            pnlOptions.Controls.Add(lblStoreStyle);
            pnlOptions.Controls.Add(cmbStoreStyle);
            pnlOptions.Controls.Add(lblLayoutScheme);
            pnlOptions.Controls.Add(cmbLayoutScheme);
            pnlOptions.Controls.Add(lblConnectorStyle);
            pnlOptions.Controls.Add(cmbConnectorStyle);
            pnlOptions.Controls.Add(lblSpacing);
            pnlOptions.Controls.Add(txtHorSpacing);
            pnlOptions.Controls.Add(txtVerSpacing);
            pnlOptions.Controls.Add(lblSpacingUnit);
            pnlOptions.Controls.Add(btnGenerate);
            pnlOptions.Controls.Add(btnClose);

            this.Controls.Add(lblTip);
            this.Controls.Add(btnImportTable);
            this.Controls.Add(lblTitle);
            this.Controls.Add(txtTitle);
            this.Controls.Add(dgvData);
            this.Controls.Add(txtStatus);
            this.Controls.Add(pnlOptions);

            this.Resize += FormDataFlowDiagram_Resize;
            LayoutControls();
        }

        private void LayoutControls()
        {
            int margin = 15;
            int width = Math.Max(800, ClientSize.Width - margin * 2);

            dgvData.Width = width;
            txtStatus.Width = width;

            Panel pnlOptions = Controls.OfType<Panel>().FirstOrDefault();
            if (pnlOptions != null)
            {
                pnlOptions.Width = width;
                pnlOptions.Top = ClientSize.Height - pnlOptions.Height - margin;
                txtStatus.Top = pnlOptions.Top - txtStatus.Height - 15;
                dgvData.Height = Math.Max(150, txtStatus.Top - dgvData.Top - 15);
                
                btnGenerate.Left = pnlOptions.Width - btnClose.Width - btnGenerate.Width - 30;
                btnClose.Left = pnlOptions.Width - btnClose.Width - 15;
            }
        }

        private void FormDataFlowDiagram_Resize(object sender, EventArgs e)
        {
            LayoutControls();
        }

        private void AppendLog(string message)
        {
            txtStatus.AppendText(Environment.NewLine + $"{DateTime.Now:HH:mm:ss} {message}");
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private string GetDiagramLevelName()
        {
            switch (diagramLevel)
            {
                case DataFlowDiagramLevel.Top:
                    return "顶层数据流图";
                case DataFlowDiagramLevel.Level1:
                    return "1层数据流图";
                default:
                    return "2层数据流图";
            }
        }

        private string GetDefaultDotSample()
        {
            return diagramLevel == DataFlowDiagramLevel.Top ? TopLevelDotSample : DefaultDotSample;
        }

        private void LoadSampleData()
        {
            try
            {
                ParseAndFillDot(GetDefaultDotSample());
                AppendLog($"已成功加载内置默认{GetDiagramLevelName()}示例！");
            }
            catch (Exception ex)
            {
                AppendLog("加载默认示例失败：" + ex.Message);
            }
        }

        private void btnImportTable_Click(object sender, EventArgs e)
        {
            if (!Clipboard.ContainsText())
            {
                AppendLog("未获取到表格或Graphviz DOT 文本，请重新复制！");
                MessageBox.Show("剪贴板中没有文本内容，请先复制 DOT 格式文本！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string text = Clipboard.GetText();
            try
            {
                if (LooksLikeJson(text))
                {
                    ParseAndFillJson(text);
                }
                else
                {
                    ParseAndFillDot(text);
                }
                MessageBox.Show("数据导入并填入表格成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                AppendLog("解析剪贴板文本失败：" + ex.Message);
                MessageBox.Show("解析剪贴板文本失败！请检查是否为规范的 Graphviz DOT 格式有向图。错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // --- DOT 格式解析器核心方法 ---
        private bool LooksLikeJson(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string trimmed = text.TrimStart();
            return trimmed.StartsWith("{") || trimmed.StartsWith("[");
        }

        private void ParseAndFillJson(string jsonText)
        {
            var serializer = new JavaScriptSerializer();
            var root = serializer.DeserializeObject(jsonText) as Dictionary<string, object>;
            if (root == null)
            {
                throw new InvalidOperationException("JSON 根节点必须是对象。");
            }

            if (root.ContainsKey("title"))
            {
                txtTitle.Text = Convert.ToString(root["title"]);
            }

            var nodesRaw = root.ContainsKey("nodes") ? root["nodes"] as object[] : null;
            var edgesRaw = root.ContainsKey("edges") ? root["edges"] as object[] : null;
            if (nodesRaw == null || nodesRaw.Length == 0)
            {
                throw new InvalidOperationException("JSON 中缺少 nodes 节点数组。");
            }

            var nodeMap = new Dictionary<string, JsonDfdNode>();
            foreach (var item in nodesRaw)
            {
                var dict = item as Dictionary<string, object>;
                if (dict == null) continue;

                string id = GetJsonString(dict, "id");
                if (string.IsNullOrWhiteSpace(id)) continue;

                nodeMap[id] = new JsonDfdNode
                {
                    Id = id,
                    Text = GetJsonString(dict, "text", id),
                    Type = NormalizeDfdType(GetJsonString(dict, "type", GuessTypeFromId(id))),
                    XPercent = GetJsonNullableDouble(dict, "x"),
                    YPercent = GetJsonNullableDouble(dict, "y")
                };
            }

            var orderedKeys = nodeMap.Keys.ToList();
            var idToIndexMap = new Dictionary<string, int>();
            for (int i = 0; i < orderedKeys.Count; i++)
            {
                idToIndexMap[orderedKeys[i]] = i + 1;
            }

            var nextMap = orderedKeys.ToDictionary(k => k, k => new List<string>());
            var lineTextMap = orderedKeys.ToDictionary(k => k, k => new List<string>());
            if (edgesRaw != null)
            {
                foreach (var item in edgesRaw)
                {
                    var dict = item as Dictionary<string, object>;
                    if (dict == null) continue;

                    string from = GetJsonString(dict, "from");
                    string to = GetJsonString(dict, "to");
                    if (!idToIndexMap.ContainsKey(from) || !idToIndexMap.ContainsKey(to)) continue;

                    nextMap[from].Add(idToIndexMap[to].ToString(CultureInfo.InvariantCulture));
                    lineTextMap[from].Add(GetJsonString(dict, "text"));
                }
            }

            dgvData.Rows.Clear();
            foreach (string key in orderedKeys)
            {
                var node = nodeMap[key];
                dgvData.Rows.Add(
                    idToIndexMap[key],
                    node.Text,
                    node.Type,
                    string.Join(", ", nextMap[key]),
                    string.Join(", ", lineTextMap[key]),
                    FormatPercent(node.XPercent),
                    FormatPercent(node.YPercent));
            }

            AppendLog($"已成功导入 JSON 坐标化数据流图数据 {orderedKeys.Count}步");
        }

        private void ParseAndFillDot(string dotText)
        {
            if (string.IsNullOrWhiteSpace(dotText)) return;

            // 提取大标题名称
            string title = "数据流图";
            Match titleMatch = Regex.Match(dotText, @"(?:digraph|graph)\s+([a-zA-Z0-9_]+)");
            if (titleMatch.Success)
            {
                title = titleMatch.Groups[1].Value;
            }
            txtTitle.Text = title;

            // 用于保存提取出来的节点和边
            var nodeMap = new Dictionary<string, DotNode>();
            var edgeList = new List<DotEdge>();

            // 按行读取解析
            string[] lines = dotText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // 正则匹配
            // 匹配节点： ID [label="xxx", shape=xxx]
            Regex nodeRegex = new Regex(@"^\s*([a-zA-Z0-9_]+)\s*\[(.*?)\]");
            // 匹配边： ID1 -> ID2 [label="xxx"]
            Regex edgeRegex = new Regex(@"^\s*([a-zA-Z0-9_]+)\s*->\s*([a-zA-Z0-9_]+)\s*(?:\[(.*?)\])?");

            foreach (var line in lines)
            {
                string trimmed = line.Trim();
                if (trimmed.StartsWith("//") || trimmed.StartsWith("#") || string.IsNullOrEmpty(trimmed))
                    continue;

                // 匹配边关系
                Match edgeM = edgeRegex.Match(trimmed);
                if (edgeM.Success)
                {
                    string from = edgeM.Groups[1].Value;
                    string to = edgeM.Groups[2].Value;
                    string attrs = edgeM.Groups[3].Value;

                    string label = "";
                    Match labelMatch = Regex.Match(attrs, @"label\s*=\s*""(.*?)""");
                    if (labelMatch.Success)
                    {
                        label = labelMatch.Groups[1].Value;
                    }

                    edgeList.Add(new DotEdge { From = from, To = to, Label = label });
                    
                    // 如果节点在边中出现但没被单独定义，动态补齐
                    if (!nodeMap.ContainsKey(from)) nodeMap[from] = new DotNode { Id = from, Label = from };
                    if (!nodeMap.ContainsKey(to)) nodeMap[to] = new DotNode { Id = to, Label = to };
                    continue;
                }

                Match nodeM = nodeRegex.Match(trimmed);
                if (nodeM.Success)
                {
                    string id = nodeM.Groups[1].Value;
                    string idLower = id.ToLower();
                    if (idLower == "node" || idLower == "edge" || idLower == "graph" || idLower == "digraph")
                        continue;

                    string attrs = nodeM.Groups[2].Value;

                    string label = id;
                    Match labelMatch = Regex.Match(attrs, @"label\s*=\s*""(.*?)""");
                    if (labelMatch.Success)
                    {
                        label = labelMatch.Groups[1].Value;
                        // 支持 \n 换行符的转义处理
                        label = label.Replace("\\n", "\n");
                    }

                    string shape = "";
                    Match shapeMatch = Regex.Match(attrs, @"shape\s*=\s*([a-zA-Z0-9_]+)");
                    if (shapeMatch.Success)
                    {
                        shape = shapeMatch.Groups[1].Value;
                    }

                    nodeMap[id] = new DotNode
                    {
                        Id = id,
                        Label = label,
                        Shape = shape,
                        XPercent = ExtractDotNumber(attrs, "x"),
                        YPercent = ExtractDotNumber(attrs, "y")
                    };
                }
            }

            // 转换成表格所需的数字序号
            var orderedKeys = nodeMap.Keys.ToList();
            var idToIndexMap = new Dictionary<string, int>();
            for (int i = 0; i < orderedKeys.Count; i++)
            {
                idToIndexMap[orderedKeys[i]] = i + 1;
            }

            // 汇总每个节点的下一步和连接线文本
            var nextMap = new Dictionary<string, List<string>>();
            var lineTextMap = new Dictionary<string, List<string>>();

            foreach (var nodeKey in orderedKeys)
            {
                nextMap[nodeKey] = new List<string>();
                lineTextMap[nodeKey] = new List<string>();
            }

            foreach (var edge in edgeList)
            {
                if (nextMap.ContainsKey(edge.From))
                {
                    int targetIndex = idToIndexMap[edge.To];
                    nextMap[edge.From].Add(targetIndex.ToString());
                    lineTextMap[edge.From].Add(edge.Label);
                }
            }

            // 填充到 DataGridView
            dgvData.Rows.Clear();
            foreach (var nodeKey in orderedKeys)
            {
                var node = nodeMap[nodeKey];
                int index = idToIndexMap[nodeKey];

                // 智能分类：P开头 = 加工，D开头 = 数据存储，其他 = 实体
                string guessType = "实体";
                string idUpper = node.Id.ToUpper();

                if (idUpper.StartsWith("P"))
                {
                    guessType = "加工";
                }
                else if (idUpper.StartsWith("D"))
                {
                    guessType = "数据存储";
                }

                string nextJoined = string.Join(", ", nextMap[nodeKey]);
                string lineTextJoined = string.Join(", ", lineTextMap[nodeKey]);

                dgvData.Rows.Add(
                    index,
                    node.Label,
                    guessType,
                    nextJoined,
                    lineTextJoined,
                    FormatPercent(node.XPercent),
                    FormatPercent(node.YPercent));
            }

            AppendLog($"已成功添加数据流图数据 {orderedKeys.Count}步");
        }

        private string GetJsonString(Dictionary<string, object> dict, string key, string fallback = "")
        {
            if (!dict.ContainsKey(key) || dict[key] == null) return fallback;
            return Convert.ToString(dict[key]);
        }

        private double? GetJsonNullableDouble(Dictionary<string, object> dict, string key)
        {
            if (!dict.ContainsKey(key) || dict[key] == null) return null;
            if (double.TryParse(Convert.ToString(dict[key]), NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                return ClampPercent(value);
            }
            return null;
        }

        private double? ExtractDotNumber(string attrs, string name)
        {
            Match match = Regex.Match(attrs, $@"(?:^|[,;\s]){Regex.Escape(name)}\s*=\s*""?(-?\d+(?:\.\d+)?)""?", RegexOptions.IgnoreCase);
            if (!match.Success) return null;
            if (double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                return ClampPercent(value);
            }
            return null;
        }

        private string FormatPercent(double? value)
        {
            return value.HasValue ? value.Value.ToString("0.##", CultureInfo.InvariantCulture) : "";
        }

        private double ClampPercent(double value)
        {
            if (value < 0) return 0;
            if (value > 100) return 100;
            return value;
        }

        private string GuessTypeFromId(string id)
        {
            string idUpper = (id ?? string.Empty).ToUpperInvariant();
            if (idUpper.StartsWith("P")) return "加工";
            if (idUpper.StartsWith("D")) return "数据存储";
            return "实体";
        }

        private string NormalizeDfdType(string type)
        {
            if (type == "实体" || type == "加工" || type == "数据存储") return type;
            return "实体";
        }

        private class DotNode
        {
            public string Id { get; set; }
            public string Label { get; set; }
            public string Shape { get; set; }
            public double? XPercent { get; set; }
            public double? YPercent { get; set; }
        }

        private class JsonDfdNode
        {
            public string Id { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
            public double? XPercent { get; set; }
            public double? YPercent { get; set; }
        }

        private class DotEdge
        {
            public string From { get; set; }
            public string To { get; set; }
            public string Label { get; set; }
        }

        // --- Visio 生成与排版核心方法 ---
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            try
            {
                RenderDataFlowDiagram();
            }
            catch (Exception ex)
            {
                AppendLog("绘图生成失败: " + ex.Message);
                MessageBox.Show("绘图生成过程中发生错误: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenderDataFlowDiagram()
        {
            Visio.Application visioApp = Globals.ThisAddIn.Application;
            Visio.Page activePage = GetOrCreateActivePage(visioApp);
            ClearPageShapes(activePage);

            double pageWidth = activePage.PageSheet.CellsU["PageWidth"].Result["in"];
            double pageHeight = activePage.PageSheet.CellsU["PageHeight"].Result["in"];

            // 读取界面间距设置
            double horSpacingInch = 1.2;
            double verSpacingInch = 1.4;
            double.TryParse(txtHorSpacing.Text, out double horMm);
            double.TryParse(txtVerSpacing.Text, out double verMm);
            if (horMm > 0) horSpacingInch = horMm / 25.4;
            if (verMm > 0) verSpacingInch = verMm / 25.4;

            // 1. 读取表格中的所有行数据，构建拓扑节点
            var nodeList = new List<DFDNode>();
            var rows = dgvData.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToList();

            foreach (var row in rows)
            {
                if (row.Cells["Index"].Value == null) continue;

                int index = Convert.ToInt32(row.Cells["Index"].Value);
                string text = Convert.ToString(row.Cells["Text"].Value) ?? "";
                string type = Convert.ToString(row.Cells["ShapeType"].Value) ?? "实体";
                string nextStr = Convert.ToString(row.Cells["Next"].Value) ?? "";
                string lineTextStr = Convert.ToString(row.Cells["LineText"].Value) ?? "";
                double? xPercent = ParsePercentCell(row.Cells["XPercent"].Value);
                double? yPercent = ParsePercentCell(row.Cells["YPercent"].Value);

                // 解析下一步序号列表
                var nextIndices = new List<int>();
                if (!string.IsNullOrWhiteSpace(nextStr))
                {
                    var parts = nextStr.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var part in parts)
                    {
                        if (int.TryParse(part, out int parsedIndex))
                        {
                            nextIndices.Add(parsedIndex);
                        }
                    }
                }

                // 解析连接线描述
                var lineTexts = new List<string>();
                if (!string.IsNullOrWhiteSpace(lineTextStr))
                {
                    var parts = lineTextStr.Split(new[] { ',' }, StringSplitOptions.None);
                    foreach (var part in parts)
                    {
                        lineTexts.Add(part.Trim());
                    }
                }

                nodeList.Add(new DFDNode
                {
                    Index = index,
                    Text = text,
                    Type = type,
                    NextIndices = nextIndices,
                    LineTexts = lineTexts,
                    XPercent = xPercent,
                    YPercent = yPercent
                });
            }

            if (nodeList.Count == 0)
            {
                throw new InvalidOperationException("表格中没有可供绘制的数据步骤！");
            }

            // 建立快速索引字典
            var indexToNodeMap = nodeList.ToDictionary(n => n.Index);
            var edgeList = BuildDfdEdges(nodeList, indexToNodeMap);

            // 如果选择了方案 B/C/E，则使用对应的排版算法提前计算绝对坐标
            Dictionary<int, Tuple<double, double>> sugiyamaCoords = null;
            string selectedScheme = cmbLayoutScheme.SelectedItem?.ToString() ?? "方案 A (Visio避让)";

            if (selectedScheme.Contains("方案 E"))
            {
                try
                {
                    var academicEdges = BuildAcademicDfdEdges(nodeList, edgeList);
                    var layout = new DfdSmartSemanticTopologyLayout();
                    var layoutNodes = nodeList.Select(n => new DfdSmartSemanticTopologyLayout.LayoutNode
                    {
                        Index = n.Index,
                        Text = n.Text,
                        Type = n.Type
                    }).ToList();
                    var layoutEdges = academicEdges.Select(e => new DfdSmartSemanticTopologyLayout.LayoutEdge
                    {
                        From = e.From,
                        To = e.To
                    }).ToList();

                    var result = layout.CalculateLayout(layoutNodes, layoutEdges, pageWidth, pageHeight, horSpacingInch, verSpacingInch);
                    sugiyamaCoords = result.Coordinates;
                    pageWidth = result.PageWidth;
                    pageHeight = result.PageHeight;
                    activePage.PageSheet.CellsU["PageWidth"].FormulaU = $"{pageWidth.ToString(CultureInfo.InvariantCulture)} in";
                    activePage.PageSheet.CellsU["PageHeight"].FormulaU = $"{pageHeight.ToString(CultureInfo.InvariantCulture)} in";
                    foreach (var warning in result.Warnings)
                    {
                        AppendLog("[建议] " + warning);
                    }
                    int hiddenEdgeCount = edgeList.Count - academicEdges.Count;
                    if (hiddenEdgeCount > 0)
                    {
                        AppendLog($"[学术版] 已隐藏 {hiddenEdgeCount} 条日志/记录类辅助长线，避免主图过度交叉。");
                    }
                    ApplyAcademicEdgesToNodes(nodeList, academicEdges);
                }
                catch (Exception ex)
                {
                    AppendLog($"[警告] 方案 E 布局计算异常，已退回到简易网格: {ex.Message}");
                }
            }
            else if (selectedScheme.Contains("方案 B"))
            {
                try
                {
                    var layout = new DfdSugiyamaLayout();
                    var layoutNodes = nodeList.Select(n => new DfdSugiyamaLayout.LayoutNode
                    {
                        Index = n.Index,
                        Text = n.Text,
                        Type = n.Type
                    }).ToList();

                    var layoutEdges = new List<DfdSugiyamaLayout.LayoutEdge>();
                    foreach (var n in nodeList)
                    {
                        foreach (var nextIdx in n.NextIndices)
                        {
                            layoutEdges.Add(new DfdSugiyamaLayout.LayoutEdge { From = n.Index, To = nextIdx });
                        }
                    }

                    // 进行 Sugiyama 层次坐标分析（左右 LR 拓扑流向）
                    sugiyamaCoords = layout.CalculateLayout(
                        layoutNodes,
                        layoutEdges,
                        pageWidth,
                        pageHeight,
                        1.8 + horSpacingInch,
                        1.2 + verSpacingInch,
                        "LR"
                    );
                }
                catch (Exception ex)
                {
                    AppendLog($"[警告] 方案 B 布局计算异常，已退回到简易网格: {ex.Message}");
                }
            }
            else if (selectedScheme.Contains("方案 C"))
            {
                try
                {
                    var layout = new DfdSemanticLayout();
                    var layoutNodes = nodeList.Select(n => new DfdSemanticLayout.LayoutNode
                    {
                        Index = n.Index,
                        Text = n.Text,
                        Type = n.Type
                    }).ToList();

                    var layoutEdges = new List<DfdSemanticLayout.LayoutEdge>();
                    foreach (var n in nodeList)
                    {
                        foreach (var nextIdx in n.NextIndices)
                        {
                            layoutEdges.Add(new DfdSemanticLayout.LayoutEdge { From = n.Index, To = nextIdx });
                        }
                    }

                    // 进行 DFD 语义分区坐标计算
                    sugiyamaCoords = layout.CalculateLayout(
                        layoutNodes,
                        layoutEdges,
                        pageWidth,
                        pageHeight,
                        horSpacingInch,
                        verSpacingInch
                    );

                    // 语义布局算法通过特殊键 -1/-2 传递计算出的有效画布尺寸
                    if (sugiyamaCoords.ContainsKey(-1))
                    {
                        double newWidth = sugiyamaCoords[-1].Item1;
                        double newHeight = sugiyamaCoords[-1].Item2;
                        sugiyamaCoords.Remove(-1); // 移除特殊键，不影响后续节点遍历

                        if (newWidth > pageWidth)
                        {
                            pageWidth = newWidth;
                            activePage.PageSheet.CellsU["PageWidth"].FormulaU = $"{pageWidth} in";
                        }
                        if (newHeight > pageHeight)
                        {
                            pageHeight = newHeight;
                            activePage.PageSheet.CellsU["PageHeight"].FormulaU = $"{pageHeight} in";
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"[警告] 方案 C 布局计算异常，已退回到简易网格: {ex.Message}");
                }
            }

            // 动态扩增画布大小，确保拓扑分层图有足够的空间排布，绝对不发生重合挤压
            if (sugiyamaCoords != null && sugiyamaCoords.Count > 0)
            {
                try
                {
                    double maxX = sugiyamaCoords.Values.Max(c => c.Item1);
                    double maxY = sugiyamaCoords.Values.Max(c => c.Item2);

                    if (maxX > pageWidth - 1.5)
                    {
                        pageWidth = maxX + 2.0;
                        activePage.PageSheet.CellsU["PageWidth"].FormulaU = $"{pageWidth} in";
                    }
                    if (maxY > pageHeight - 1.5)
                    {
                        pageHeight = maxY + 2.0;
                        activePage.PageSheet.CellsU["PageHeight"].FormulaU = $"{pageHeight} in";
                    }
                }
                catch { }
            }

            ApplyCoordinateOverrides(nodeList, ref sugiyamaCoords, pageWidth, pageHeight);

            // 2. 在画布上安置并绘制各节点
            int columns = (int)Math.Ceiling(Math.Sqrt(nodeList.Count));
            double startGridX = 1.0;
            double startGridY = pageHeight - 1.5;

            for (int i = 0; i < nodeList.Count; i++)
            {
                var node = nodeList[i];
                double initialX = 0;
                double initialY = 0;

                if (sugiyamaCoords != null && sugiyamaCoords.ContainsKey(node.Index))
                {
                    initialX = sugiyamaCoords[node.Index].Item1;
                    initialY = sugiyamaCoords[node.Index].Item2;
                }
                else
                {
                    int col = i % columns;
                    int row = i / columns;
                    initialX = startGridX + col * (1.8 + horSpacingInch);
                    initialY = startGridY - row * (1.2 + verSpacingInch);
                }

                // 根据类型绘制不同的学术单色风形状
                Visio.Shape shape = null;
                double w = 1.2;
                double h = 0.8;

                if (node.Type == "加工")
                {
                    string procStyle = cmbProcessStyle.SelectedItem?.ToString() ?? "正方形-带标";
                    if (procStyle == "圆角矩形")
                    {
                        shape = activePage.DrawRectangle(initialX - w / 2, initialY - h / 2, initialX + w / 2, initialY + h / 2);
                        shape.Text = node.Text;
                        ApplyAcademicShapeStyle(shape);
                        TrySetFormula(shape, "Roundness", "0.18");
                    }
                    else
                    {
                        // 正方形-带标
                        shape = DrawProcessWithHeader(activePage, node.Text, initialX, initialY, 1.0, 1.0);
                    }
                }
                else if (node.Type == "数据存储")
                {
                    string storeStyle = cmbStoreStyle.SelectedItem?.ToString() ?? "小正方形+两横线";
                    if (storeStyle == "三边矩形(右开口)")
                    {
                        shape = DrawOpenRectangle(activePage, node.Text, initialX, initialY, w, h);
                    }
                    else
                    {
                        // 小正方形+两横线
                        shape = DrawDoubleLineStore(activePage, node.Text, initialX, initialY, w, h);
                    }
                }
                else
                {
                    // 外部实体（普通矩形）
                    shape = activePage.DrawRectangle(initialX - w / 2, initialY - h / 2, initialX + w / 2, initialY + h / 2);
                    shape.Text = node.Text;
                    ApplyAcademicShapeStyle(shape);
                }

                node.VisioShape = shape;
            }

            // 3. 绘制连接线并静态胶连至 PinX
            object connectorTool = activePage.Application.ConnectorToolDataObject;
            foreach (var node in nodeList)
            {
                if (node.VisioShape == null) continue;

                for (int k = 0; k < node.NextIndices.Count; k++)
                {
                    int targetIdx = node.NextIndices[k];
                    if (!indexToNodeMap.ContainsKey(targetIdx)) continue;

                    var targetNode = indexToNodeMap[targetIdx];
                    if (targetNode.VisioShape == null) continue;

                    string lineText = k < node.LineTexts.Count ? node.LineTexts[k] : "";

                    // Drop 一个 Dynamic Connector
                    Visio.Shape connector = activePage.Drop(connectorTool, 0, 0);

                    // 应用连线线宽、字号和箭头样式
                    connector.CellsU["LineColor"].Formula = "RGB(0, 0, 0)";
                    connector.CellsU["LineWeight"].Formula = "0.75pt";
                    connector.CellsU["BeginArrow"].Formula = "0";
                    connector.CellsU["EndArrow"].Formula = "4"; // 带终点箭头
                    connector.Text = lineText;

                    // 设置连接线文字的字体样式
                    connector.CellsU["Char.Font"].FormulaU = $"\"{currentFontName}\"";
                    connector.CellsU["Char.Size"].FormulaU = "9pt";
                    connector.CellsU["Char.Color"].Formula = "RGB(0, 0, 0)";

                    string connStyle = cmbConnectorStyle.SelectedItem?.ToString() ?? "绑定形状";

                    if (connStyle == "绑定新增连接点")
                    {
                        // 动态为起终点形状新增中心连接点并胶连（Point-to-Point）
                        short rowBegin = node.VisioShape.AddRow((short)Visio.VisSectionIndices.visSectionConnectionPts, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagCnnctPt);
                        node.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowBegin, (short)Visio.VisCellIndices.visCnnctX].FormulaU = "Width*0.5";
                        node.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowBegin, (short)Visio.VisCellIndices.visCnnctY].FormulaU = "Height*0.5";
                        connector.CellsU["BeginX"].GlueTo(node.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowBegin, (short)Visio.VisCellIndices.visCnnctX]);

                        short rowEnd = targetNode.VisioShape.AddRow((short)Visio.VisSectionIndices.visSectionConnectionPts, (short)Visio.VisRowIndices.visRowLast, (short)Visio.VisRowTags.visTagCnnctPt);
                        targetNode.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowEnd, (short)Visio.VisCellIndices.visCnnctX].FormulaU = "Width*0.5";
                        targetNode.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowEnd, (short)Visio.VisCellIndices.visCnnctY].FormulaU = "Height*0.5";
                        connector.CellsU["EndX"].GlueTo(targetNode.VisioShape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, rowEnd, (short)Visio.VisCellIndices.visCnnctX]);
                    }
                    else if (connStyle == "普通动态连线")
                    {
                        // 不进行硬胶连，仅将连线端点放置在形状中心坐标附近（静态放置，不随形状移动而调整）
                        connector.CellsU["BeginX"].FormulaU = $"{node.VisioShape.CellsU["PinX"].ResultIU} in";
                        connector.CellsU["BeginY"].FormulaU = $"{node.VisioShape.CellsU["PinY"].ResultIU} in";
                        connector.CellsU["EndX"].FormulaU = $"{targetNode.VisioShape.CellsU["PinX"].ResultIU} in";
                        connector.CellsU["EndY"].FormulaU = $"{targetNode.VisioShape.CellsU["PinY"].ResultIU} in";
                    }
                    else
                    {
                        if (selectedScheme.Contains("方案 E"))
                        {
                            GlueConnectorToAcademicSides(connector, node.VisioShape, targetNode.VisioShape);
                        }
                        else
                        {
                            // 默认“绑定形状”：胶连到图形的 PinX (Shape-to-Shape)，Visio 会自动寻找边缘最短路径
                            connector.CellsU["BeginX"].GlueTo(node.VisioShape.CellsU["PinX"]);
                            connector.CellsU["EndX"].GlueTo(targetNode.VisioShape.CellsU["PinX"]);
                        }
                    }
                }
            }

            // 4. 调用 Visio 强大的流程图网状布局自动整理算法，瞬间让全图直角规范排列
            try
            {
                Visio.Shape pageSheet = activePage.PageSheet;
                pageSheet.CellsU["RouteStyle"].Formula = "5"; // 正交折线路由
                pageSheet.CellsU["LineToNodeGap"].FormulaU = selectedScheme.Contains("方案 E") ? "0.36 in" : "0.20 in";
                pageSheet.CellsU["LineToLineGap"].FormulaU = selectedScheme.Contains("方案 E") ? "0.24 in" : "0.15 in";

                if (selectedScheme.Contains("方案 A"))
                {
                    pageSheet.CellsU["PlaceStyle"].Formula = "1"; // 流程图自上而下排版
                    pageSheet.CellsU["AvenueSizeX"].FormulaU = "2.0 in";      // 节点水平间距，拉大以提供走线通道
                    pageSheet.CellsU["AvenueSizeY"].FormulaU = "2.0 in";      // 节点垂直间距
                }
                else
                {
                    // 方案 B / C / E 禁用 PlaceStyle 排版（防止 Visio 覆盖我们的精准坐标）
                    pageSheet.CellsU["PlaceStyle"].Formula = "0";
                    pageSheet.CellsU["AvenueSizeX"].FormulaU = selectedScheme.Contains("方案 E") ? "2.8 in" : "2.0 in";
                    pageSheet.CellsU["AvenueSizeY"].FormulaU = selectedScheme.Contains("方案 E") ? "2.4 in" : "2.0 in";
                }
            }
            catch { }

            // 只有方案 A 才需要整体重排，方案 B 和 C 的节点已经利用算法计算出坐标精准落位
            if (selectedScheme.Contains("方案 A"))
            {
                activePage.Layout();
                AppendLog($"生成{GetDiagramLevelName()}绘图成功！已应用方案 A 避让优化。");
            }
            else if (selectedScheme.Contains("方案 B"))
            {
                AppendLog($"生成{GetDiagramLevelName()}绘图成功！已应用方案 B (Sugiyama 拓扑分层) 布局算法。");
            }
            else if (selectedScheme.Contains("方案 E"))
            {
                AppendLog($"生成{GetDiagramLevelName()}绘图成功！已应用方案 E (智能语义拓扑布局)。");
            }
            else
            {
                AppendLog($"生成{GetDiagramLevelName()}绘图成功！已应用方案 C (DFD 语义分区) 布局算法。");
            }

            MessageBox.Show($"{GetDiagramLevelName()}生成成功！已完成避让与正交连线布局。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Close();
        }

        // --- 高级特殊 DFD 样式组装助手 ---

        private List<DfdOrthogonalChannelLayout.LayoutEdge> BuildDfdEdges(List<DFDNode> nodeList, Dictionary<int, DFDNode> indexToNodeMap)
        {
            var edges = new List<DfdOrthogonalChannelLayout.LayoutEdge>();
            foreach (var node in nodeList)
            {
                for (int i = 0; i < node.NextIndices.Count; i++)
                {
                    int targetIndex = node.NextIndices[i];
                    if (!indexToNodeMap.ContainsKey(targetIndex))
                    {
                        continue;
                    }

                    edges.Add(new DfdOrthogonalChannelLayout.LayoutEdge
                    {
                        From = node.Index,
                        To = targetIndex,
                        Text = i < node.LineTexts.Count ? node.LineTexts[i] : ""
                    });
                }
            }

            return edges;
        }

        private double? ParsePercentCell(object value)
        {
            if (value == null) return null;
            string text = Convert.ToString(value)?.Trim();
            if (string.IsNullOrWhiteSpace(text)) return null;
            text = text.Replace("%", "");

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
            {
                return ClampPercent(parsed);
            }

            return null;
        }

        private void ApplyCoordinateOverrides(
            List<DFDNode> nodeList,
            ref Dictionary<int, Tuple<double, double>> coordinates,
            double pageWidth,
            double pageHeight)
        {
            var coordinateNodes = nodeList
                .Where(n => n.XPercent.HasValue && n.YPercent.HasValue)
                .ToList();
            if (coordinateNodes.Count == 0)
            {
                return;
            }

            if (coordinates == null)
            {
                coordinates = new Dictionary<int, Tuple<double, double>>();
            }

            foreach (var node in coordinateNodes)
            {
                double x = pageWidth * node.XPercent.Value / 100.0;
                double y = pageHeight * (100.0 - node.YPercent.Value) / 100.0;
                coordinates[node.Index] = Tuple.Create(x, y);
            }

            AppendLog($"已应用 {coordinateNodes.Count} 个 AI/表格坐标节点。");
        }

        private List<DfdOrthogonalChannelLayout.LayoutEdge> BuildAcademicDfdEdges(
            List<DFDNode> nodeList,
            List<DfdOrthogonalChannelLayout.LayoutEdge> allEdges)
        {
            var nodeMap = nodeList.ToDictionary(n => n.Index);
            var kept = allEdges
                .Where(e => ShouldKeepAcademicEdge(e, nodeMap))
                .ToList();

            return kept.Count == 0 ? allEdges : kept;
        }

        private bool ShouldKeepAcademicEdge(DfdOrthogonalChannelLayout.LayoutEdge edge, Dictionary<int, DFDNode> nodeMap)
        {
            if (!nodeMap.ContainsKey(edge.From) || !nodeMap.ContainsKey(edge.To))
            {
                return false;
            }

            var from = nodeMap[edge.From];
            var to = nodeMap[edge.To];
            if (IsLogLikeEdge(edge.Text) && !IsLogLikeNode(from) && IsLogLikeNode(to))
            {
                return false;
            }

            if (from.Type == "加工" && to.Type == "加工" && IsLogLikeNode(to) && !IsLogLikeNode(from))
            {
                return false;
            }

            return true;
        }

        private bool IsLogLikeNode(DFDNode node)
        {
            string text = (node.Text ?? string.Empty).Replace("\n", "");
            return text.Contains("日志") || text.Contains("记录") || text.Contains("审计");
        }

        private bool IsLogLikeEdge(string text)
        {
            string value = text ?? string.Empty;
            return value.Contains("日志") || value.Contains("操作记录") || value.Contains("历史记录") || value.Contains("记录");
        }

        private void ApplyAcademicEdgesToNodes(List<DFDNode> nodeList, List<DfdOrthogonalChannelLayout.LayoutEdge> edges)
        {
            var outgoing = edges
                .GroupBy(e => e.From)
                .ToDictionary(g => g.Key, g => g.ToList());

            foreach (var node in nodeList)
            {
                if (!outgoing.ContainsKey(node.Index))
                {
                    node.NextIndices = new List<int>();
                    node.LineTexts = new List<string>();
                    continue;
                }

                node.NextIndices = outgoing[node.Index].Select(e => e.To).ToList();
                node.LineTexts = outgoing[node.Index].Select(e => e.Text ?? string.Empty).ToList();
            }
        }

        private void GlueConnectorToAcademicSides(Visio.Shape connector, Visio.Shape fromShape, Visio.Shape toShape)
        {
            double fromX = fromShape.CellsU["PinX"].ResultIU;
            double fromY = fromShape.CellsU["PinY"].ResultIU;
            double toX = toShape.CellsU["PinX"].ResultIU;
            double toY = toShape.CellsU["PinY"].ResultIU;

            string beginCell = Math.Abs(toX - fromX) >= Math.Abs(toY - fromY)
                ? (toX >= fromX ? "Connections.X3" : "Connections.X4")
                : (toY >= fromY ? "Connections.X2" : "Connections.X1");
            string endCell = Math.Abs(toX - fromX) >= Math.Abs(toY - fromY)
                ? (toX >= fromX ? "Connections.X4" : "Connections.X3")
                : (toY >= fromY ? "Connections.X1" : "Connections.X2");

            EnsureAcademicConnectionPoints(fromShape);
            EnsureAcademicConnectionPoints(toShape);
            connector.CellsU["BeginX"].GlueTo(fromShape.CellsU[beginCell]);
            connector.CellsU["EndX"].GlueTo(toShape.CellsU[endCell]);
        }

        private void EnsureAcademicConnectionPoints(Visio.Shape shape)
        {
            try
            {
                if (shape.SectionExists[(short)Visio.VisSectionIndices.visSectionConnectionPts, 0] == 0 ||
                    shape.RowCount[(short)Visio.VisSectionIndices.visSectionConnectionPts] < 4)
                {
                    SetConnectionPoint(shape, 0, "Width*0.5", "0");
                    SetConnectionPoint(shape, 1, "Width*0.5", "Height");
                    SetConnectionPoint(shape, 2, "Width", "Height*0.5");
                    SetConnectionPoint(shape, 3, "0", "Height*0.5");
                }
            }
            catch { }
        }

        private void SetConnectionPoint(Visio.Shape shape, short row, string xFormula, string yFormula)
        {
            while (shape.RowCount[(short)Visio.VisSectionIndices.visSectionConnectionPts] <= row)
            {
                shape.AddRow(
                    (short)Visio.VisSectionIndices.visSectionConnectionPts,
                    (short)Visio.VisRowIndices.visRowLast,
                    (short)Visio.VisRowTags.visTagCnnctPt);
            }

            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, row, (short)Visio.VisCellIndices.visCnnctX].FormulaU = xFormula;
            shape.CellsSRC[(short)Visio.VisSectionIndices.visSectionConnectionPts, row, (short)Visio.VisCellIndices.visCnnctY].FormulaU = yFormula;
        }

        private void RenderOrthogonalChannelDataFlow(
            Visio.Page page,
            List<DFDNode> nodeList,
            List<DfdOrthogonalChannelLayout.LayoutEdge> edgeList,
            double pageWidth,
            double pageHeight,
            double horSpacingInch,
            double verSpacingInch)
        {
            var layout = new DfdOrthogonalChannelLayout();
            var layoutNodes = nodeList.Select(n => new DfdOrthogonalChannelLayout.LayoutNode
            {
                Index = n.Index,
                Text = n.Text,
                Type = n.Type
            }).ToList();

            var result = layout.CalculateLayout(layoutNodes, edgeList, pageWidth, pageHeight, horSpacingInch, verSpacingInch);
            page.PageSheet.CellsU["PageWidth"].FormulaU = $"{result.PageWidth.ToString(CultureInfo.InvariantCulture)} in";
            page.PageSheet.CellsU["PageHeight"].FormulaU = $"{result.PageHeight.ToString(CultureInfo.InvariantCulture)} in";

            DrawDfdNodesAtLayout(page, nodeList, result.Nodes);
            DrawOrthogonalRoutedEdges(page, result.Edges);
            CreateOrthogonalSubDiagramPages(page.Application, nodeList, edgeList, result.SubDiagrams, horSpacingInch, verSpacingInch);
        }

        private void DrawDfdNodesAtLayout(Visio.Page page, List<DFDNode> nodeList, Dictionary<int, DfdOrthogonalChannelLayout.LayoutPoint> coordinates)
        {
            foreach (var node in nodeList)
            {
                if (!coordinates.ContainsKey(node.Index))
                {
                    continue;
                }

                var point = coordinates[node.Index];
                node.VisioShape = DrawDfdNodeShape(page, node, point.X, point.Y);
            }
        }

        private Visio.Shape DrawDfdNodeShape(Visio.Page page, DFDNode node, double x, double y)
        {
            double w = 1.2;
            double h = 0.8;

            if (node.Type == "加工")
            {
                string procStyle = cmbProcessStyle.SelectedItem?.ToString() ?? "正方形-带标";
                if (procStyle == "圆角矩形")
                {
                    Visio.Shape shape = page.DrawRectangle(x - w / 2, y - h / 2, x + w / 2, y + h / 2);
                    shape.Text = node.Text;
                    ApplyAcademicShapeStyle(shape);
                    TrySetFormula(shape, "Roundness", "0.18");
                    return shape;
                }

                return DrawProcessWithHeader(page, node.Text, x, y, 1.0, 1.0);
            }

            if (node.Type == "数据存储")
            {
                string storeStyle = cmbStoreStyle.SelectedItem?.ToString() ?? "小正方形+两横线";
                return storeStyle == "三边矩形(右开口)"
                    ? DrawOpenRectangle(page, node.Text, x, y, w, h)
                    : DrawDoubleLineStore(page, node.Text, x, y, w, h);
            }

            Visio.Shape entity = page.DrawRectangle(x - w / 2, y - h / 2, x + w / 2, y + h / 2);
            entity.Text = node.Text;
            ApplyAcademicShapeStyle(entity);
            return entity;
        }

        private void DrawOrthogonalRoutedEdges(Visio.Page page, List<DfdOrthogonalChannelLayout.RoutedEdge> edges)
        {
            foreach (var edge in edges)
            {
                var pieces = new List<Visio.Shape>();
                for (int i = 0; i < edge.Points.Count - 1; i++)
                {
                    var a = edge.Points[i];
                    var b = edge.Points[i + 1];
                    Visio.Shape segment = page.DrawLine(a.X, a.Y, b.X, b.Y);
                    segment.CellsU["LineColor"].Formula = "RGB(0, 0, 0)";
                    segment.CellsU["LineWeight"].Formula = "0.75pt";
                    segment.CellsU["BeginArrow"].Formula = "0";
                    segment.CellsU["EndArrow"].Formula = i == edge.Points.Count - 2 ? "4" : "0";
                    pieces.Add(segment);
                }

                if (!string.IsNullOrWhiteSpace(edge.Text))
                {
                    var labelPoint = edge.LabelPoint ?? edge.Points[Math.Max(0, edge.Points.Count / 2)];
                    Visio.Shape label = page.DrawRectangle(labelPoint.X - 0.55, labelPoint.Y - 0.12, labelPoint.X + 0.55, labelPoint.Y + 0.12);
                    label.Text = edge.Text;
                    label.CellsU["LinePattern"].Formula = "0";
                    label.CellsU["FillPattern"].Formula = "0";
                    label.CellsU["Char.Font"].FormulaU = $"\"{currentFontName}\"";
                    label.CellsU["Char.Size"].FormulaU = "8.5pt";
                    label.CellsU["Char.Color"].Formula = "RGB(0, 0, 0)";
                    label.CellsU["Para.HorzAlign"].Formula = "1";
                    label.CellsU["VerticalAlign"].Formula = "1";
                    pieces.Add(label);
                }

                if (pieces.Count > 1)
                {
                    try
                    {
                        GroupShapes(page, pieces);
                    }
                    catch
                    {
                        // 分组失败不影响图形内容，线段和文字仍然保留。
                    }
                }
            }
        }

        private void CreateOrthogonalSubDiagramPages(
            Visio.Application app,
            List<DFDNode> allNodes,
            List<DfdOrthogonalChannelLayout.LayoutEdge> allEdges,
            List<DfdOrthogonalChannelLayout.SubDiagramPlan> subDiagrams,
            double horSpacingInch,
            double verSpacingInch)
        {
            if (subDiagrams == null || subDiagrams.Count == 0)
            {
                return;
            }

            var nodeMap = allNodes.ToDictionary(n => n.Index);
            foreach (var plan in subDiagrams)
            {
                if (!nodeMap.ContainsKey(plan.FocusNodeIndex))
                {
                    continue;
                }

                var localEdges = allEdges.Where(e => e.From == plan.FocusNodeIndex || e.To == plan.FocusNodeIndex).ToList();
                var localIds = new HashSet<int>(localEdges.SelectMany(e => new[] { e.From, e.To }));
                var localNodes = allNodes
                    .Where(n => localIds.Contains(n.Index))
                    .Select(CloneDfdNode)
                    .ToList();

                if (localNodes.Count <= 1)
                {
                    continue;
                }

                Visio.Page page = app.ActiveDocument.Pages.Add();
                try
                {
                    page.Name = EnsureUniquePageName(app.ActiveDocument, plan.PageName);
                }
                catch { }

                var layout = new DfdOrthogonalChannelLayout();
                var layoutNodes = localNodes.Select(n => new DfdOrthogonalChannelLayout.LayoutNode
                {
                    Index = n.Index,
                    Text = n.Text,
                    Type = n.Type
                }).ToList();

                var result = layout.CalculateLayout(layoutNodes, localEdges, 11.0, 8.5, horSpacingInch, verSpacingInch);
                page.PageSheet.CellsU["PageWidth"].FormulaU = $"{result.PageWidth.ToString(CultureInfo.InvariantCulture)} in";
                page.PageSheet.CellsU["PageHeight"].FormulaU = $"{result.PageHeight.ToString(CultureInfo.InvariantCulture)} in";
                DrawDfdNodesAtLayout(page, localNodes, result.Nodes);
                DrawOrthogonalRoutedEdges(page, result.Edges);
            }
        }

        private DFDNode CloneDfdNode(DFDNode node)
        {
            return new DFDNode
            {
                Index = node.Index,
                Text = node.Text,
                Type = node.Type,
                NextIndices = new List<int>(node.NextIndices),
                LineTexts = new List<string>(node.LineTexts),
                XPercent = node.XPercent,
                YPercent = node.YPercent
            };
        }

        private string EnsureUniquePageName(Visio.Document document, string baseName)
        {
            string safeName = string.IsNullOrWhiteSpace(baseName) ? "2层数据流图-局部图" : baseName;
            if (safeName.Length > 31)
            {
                safeName = safeName.Substring(0, 31);
            }

            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 1; i <= document.Pages.Count; i++)
            {
                existing.Add(document.Pages[i].Name);
            }

            if (!existing.Contains(safeName))
            {
                return safeName;
            }

            for (int i = 2; i < 100; i++)
            {
                string suffix = "-" + i.ToString(CultureInfo.InvariantCulture);
                string candidate = safeName.Length + suffix.Length > 31
                    ? safeName.Substring(0, 31 - suffix.Length) + suffix
                    : safeName + suffix;
                if (!existing.Contains(candidate))
                {
                    return candidate;
                }
            }

            return safeName;
        }

        private Visio.Shape GroupShapes(Visio.Page page, List<Visio.Shape> shapes)
        {
            Visio.Selection selection = page.Application.ActiveWindow.Selection;
            selection.DeselectAll();
            foreach (Visio.Shape shape in shapes)
            {
                selection.Select(shape, 2); // 2 == visSelect
            }
            return selection.Group();
        }

        // 绘制带上边标头的正方形加工形状
        private Visio.Shape DrawProcessWithHeader(Visio.Page page, string text, double cx, double cy, double w, double h)
        {
            var shapes = new List<Visio.Shape>();
            // 外框
            Visio.Shape box = page.DrawRectangle(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2);
            ApplyAcademicShapeStyle(box);
            shapes.Add(box);

            // 分割线 Y 轴高度（距离顶部 25%）
            double lineY = cy + h / 2 - h * 0.25;
            Visio.Shape line = page.DrawLine(cx - w / 2, lineY, cx + w / 2, lineY);
            ApplyAcademicShapeStyle(line);
            shapes.Add(line);

            // 组装打组
            box.Text = text;
            Visio.Shape group = GroupShapes(page, shapes);
            ApplyAcademicShapeStyle(group);
            return group;
        }

        // 绘制右侧开口的数据存储形状 [
        private Visio.Shape DrawOpenRectangle(Visio.Page page, string text, double cx, double cy, double w, double h)
        {
            var shapes = new List<Visio.Shape>();

            // 绘制 3 条线：上、左、下
            Visio.Shape topLine = page.DrawLine(cx - w / 2, cy + h / 2, cx + w / 2, cy + h / 2);
            ApplyAcademicShapeStyle(topLine);
            shapes.Add(topLine);

            Visio.Shape leftLine = page.DrawLine(cx - w / 2, cy + h / 2, cx - w / 2, cy - h / 2);
            ApplyAcademicShapeStyle(leftLine);
            shapes.Add(leftLine);

            Visio.Shape bottomLine = page.DrawLine(cx - w / 2, cy - h / 2, cx + w / 2, cy - h / 2);
            ApplyAcademicShapeStyle(bottomLine);
            shapes.Add(bottomLine);

            // 用一个无边框的文本矩形承载文字
            Visio.Shape textRect = page.DrawRectangle(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2);
            textRect.Text = text;
            ApplyAcademicShapeStyle(textRect);
            textRect.CellsU["LinePattern"].Formula = "0";
            textRect.CellsU["FillPattern"].Formula = "0";
            shapes.Add(textRect);

            // 打组
            Visio.Shape group = GroupShapes(page, shapes);
            ApplyAcademicShapeStyle(group);
            return group;
        }

        // 绘制上下平行线的数据存储形状 =
        private Visio.Shape DrawDoubleLineStore(Visio.Page page, string text, double cx, double cy, double w, double h)
        {
            var shapes = new List<Visio.Shape>();

            // 顶部横线
            Visio.Shape topLine = page.DrawLine(cx - w / 2, cy + h / 2, cx + w / 2, cy + h / 2);
            ApplyAcademicShapeStyle(topLine);
            shapes.Add(topLine);

            // 底部横线
            Visio.Shape bottomLine = page.DrawLine(cx - w / 2, cy - h / 2, cx + w / 2, cy - h / 2);
            ApplyAcademicShapeStyle(bottomLine);
            shapes.Add(bottomLine);

            // 文字框
            Visio.Shape textRect = page.DrawRectangle(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2);
            textRect.Text = text;
            ApplyAcademicShapeStyle(textRect);
            textRect.CellsU["LinePattern"].Formula = "0";
            textRect.CellsU["FillPattern"].Formula = "0";
            shapes.Add(textRect);

            // 打组
            Visio.Shape group = GroupShapes(page, shapes);
            ApplyAcademicShapeStyle(group);
            return group;
        }

        // --- 通用画图底层辅助方法 ---

        private Visio.Page GetOrCreateActivePage(Visio.Application visioApp)
        {
            if (visioApp.ActivePage == null)
            {
                Visio.Document doc = visioApp.Documents.Add("");
                return doc.Pages[1];
            }
            return visioApp.ActivePage;
        }

        private void ClearPageShapes(Visio.Page page)
        {
            int shapeCount = page.Shapes.Count;
            for (int i = shapeCount; i >= 1; i--)
            {
                page.Shapes[i].Delete();
            }
        }

        private void ApplyAcademicShapeStyle(Visio.Shape shape)
        {
            try
            {
                shape.CellsU["FillPattern"].Formula = "0"; // 无填充
                shape.CellsU["LineColor"].Formula = "RGB(0, 0, 0)"; // 黑色边线
                shape.CellsU["LineWeight"].Formula = "0.75pt";
                shape.CellsU["Char.Color"].Formula = "RGB(0, 0, 0)";
                shape.CellsU["Char.Size"].FormulaU = $"{currentFontSizePt.ToString(CultureInfo.InvariantCulture)}pt";
                shape.CellsU["Char.Font"].FormulaU = $"\"{currentFontName}\"";
                shape.CellsU["Para.HorzAlign"].Formula = "1"; // 居中
                shape.CellsU["VerticalAlign"].Formula = "1";

                TrySetFormula(shape, "TxtPinX", "Width*0.5");
                TrySetFormula(shape, "TxtPinY", "Height*0.5");
                TrySetFormula(shape, "TxtWidth", "Width");
                TrySetFormula(shape, "TxtHeight", "Height");
                TrySetFormula(shape, "LeftMargin", "0");
                TrySetFormula(shape, "RightMargin", "0");
                TrySetFormula(shape, "TopMargin", "0");
                TrySetFormula(shape, "BottomMargin", "0");

                // 将形状声明为避让障碍物，连线必须从外侧绕行
                TrySetFormula(shape, "ObjType", "2");
            }
            catch { }
        }

        private void TrySetFormula(Visio.Shape shape, string cellName, string formula)
        {
            try
            {
                shape.CellsU[cellName].FormulaU = formula;
            }
            catch { }
        }

        private class DFDNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; } // 实体, 加工, 数据存储
            public List<int> NextIndices { get; set; } = new List<int>();
            public List<string> LineTexts { get; set; } = new List<string>();
            public double? XPercent { get; set; }
            public double? YPercent { get; set; }
            public Visio.Shape VisioShape { get; set; }
        }
    }
}
