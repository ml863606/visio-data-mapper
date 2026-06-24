using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace VisioDataMapper
{
    public class JsonVisualEditorControl : UserControl
    {
        private readonly TabControl tabControl;
        private readonly TabPage tabSource;
        private readonly TabPage tabEditor;
        private readonly TextBox txtJson;
        private readonly TreeView treeNodes;
        private readonly SplitContainer editorSplit;
        private TextBox txtPath;
        private TextBox txtType;
        private TextBox txtValue;
        private readonly ContextMenuStrip treeMenu;
        private readonly Button btnFormat;
        private readonly Button btnApply;
        private readonly Button btnExpandAll;
        private readonly Button btnCollapseAll;
        private readonly FlowLayoutPanel toolbar;
        private bool internalChange;

        public event EventHandler JsonTextChanged;

        public JsonVisualEditorControl()
        {
            Dock = DockStyle.Fill;
            BackColor = Color.White;

            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular)
            };

            tabSource = new TabPage("JSON源码") { BackColor = Color.White };
            tabEditor = new TabPage("可视化编辑") { BackColor = Color.White };

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
            txtJson.TextChanged += txtJson_TextChanged;

            toolbar = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 36,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Padding = new Padding(6, 5, 6, 3),
                BackColor = Color.FromArgb(248, 250, 252)
            };

            btnFormat = CreateToolbarButton("格式化", NodeIconKind.Format);
            btnFormat.Click += (s, e) => FormatJson();
            btnApply = CreateToolbarButton("应用当前值", NodeIconKind.Apply);
            btnApply.Click += (s, e) => ApplySelectedNodeToJson();
            btnExpandAll = CreateToolbarButton("全部展开", NodeIconKind.ExpandAll);
            btnExpandAll.Click += (s, e) => ExpandAllNodes();
            btnCollapseAll = CreateToolbarButton("全部收缩", NodeIconKind.CollapseAll);
            btnCollapseAll.Click += (s, e) => CollapseAllNodes();
            toolbar.Controls.Add(btnFormat);
            toolbar.Controls.Add(btnApply);
            toolbar.Controls.Add(btnExpandAll);
            toolbar.Controls.Add(btnCollapseAll);

            treeNodes = CreateTree();
            treeNodes.AfterSelect += treeNodes_AfterSelect;
            treeNodes.NodeMouseClick += treeNodes_NodeMouseClick;
            treeMenu = CreateTreeMenu();
            treeNodes.ContextMenuStrip = treeMenu;
            editorSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 360,
                BackColor = Color.White
            };
            editorSplit.Panel1.Controls.Add(treeNodes);
            editorSplit.Panel2.Controls.Add(CreateDetailPanel());

            tabSource.Controls.Add(txtJson);
            tabEditor.Controls.Add(editorSplit);
            tabEditor.Controls.Add(toolbar);
            tabControl.TabPages.Add(tabSource);
            tabControl.TabPages.Add(tabEditor);
            Controls.Add(tabControl);
        }

        public string JsonText
        {
            get { return txtJson.Text; }
            set { SetJsonText(value, true); }
        }

        public void SetJsonText(string json, bool format)
        {
            SetJsonText(json, format, null, null);
        }

        private void SetJsonText(string json, bool format, HashSet<string> expandedPaths, string selectedPath)
        {
            internalChange = true;
            try
            {
                txtJson.Text = format ? TryFormatJson(json) : (json ?? string.Empty);
                LoadTreeFromJson(expandedPaths, selectedPath);
            }
            finally
            {
                internalChange = false;
            }
            OnJsonTextChanged();
        }

        public void FormatJson()
        {
            SetJsonText(txtJson.Text, true);
        }

        private Button CreateToolbarButton(string text, NodeIconKind iconKind)
        {
            return new Button
            {
                Text = text,
                Image = CreateNodeIcon(iconKind),
                TextImageRelation = TextImageRelation.ImageBeforeText,
                ImageAlign = ContentAlignment.MiddleLeft,
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = true,
                Height = 26,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(31, 41, 55),
                Padding = new Padding(5, 0, 6, 0),
                Margin = new Padding(0, 0, 8, 0)
            };
        }

        private TreeView CreateTree()
        {
            return new TreeView
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                HideSelection = false,
                FullRowSelect = true,
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Regular)
            };
        }

        private ContextMenuStrip CreateTreeMenu()
        {
            var menu = new ContextMenuStrip();
            AddMenuItem(menu, "新增同级节点", NodeIconKind.AddSibling, (s, e) => AddSiblingNode());
            AddMenuItem(menu, "新增子节点", NodeIconKind.AddChild, (s, e) => AddChildNode());
            AddMenuItem(menu, "更改当前节点的值", NodeIconKind.Edit, (s, e) => ApplySelectedNodeToJson());
            AddMenuItem(menu, "删除当前节点", NodeIconKind.Delete, (s, e) => DeleteSelectedNode());
            AddMenuItem(menu, "隐藏当前节点", NodeIconKind.Hide, (s, e) => HideSelectedNode());
            return menu;
        }

        private void AddMenuItem(ContextMenuStrip menu, string text, NodeIconKind iconKind, EventHandler onClick)
        {
            var item = new ToolStripMenuItem(text, CreateNodeIcon(iconKind), onClick)
            {
                ImageScaling = ToolStripItemImageScaling.None
            };
            menu.Items.Add(item);
        }

        private Bitmap CreateNodeIcon(NodeIconKind kind)
        {
            var bitmap = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bitmap))
            using (Pen stroke = new Pen(Color.FromArgb(71, 85, 105), 1.6F))
            using (Pen greenPen = new Pen(Color.FromArgb(22, 163, 74), 1.8F))
            using (Pen bluePen = new Pen(Color.FromArgb(37, 99, 235), 1.6F))
            using (Pen redPen = new Pen(Color.FromArgb(220, 38, 38), 1.8F))
            using (Pen amberPen = new Pen(Color.FromArgb(217, 119, 6), 1.6F))
            using (Brush greenBrush = new SolidBrush(Color.FromArgb(22, 163, 74)))
            using (Brush blueBrush = new SolidBrush(Color.FromArgb(37, 99, 235)))
            using (Brush slateBrush = new SolidBrush(Color.FromArgb(71, 85, 105)))
            using (Brush redBrush = new SolidBrush(Color.FromArgb(220, 38, 38)))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                switch (kind)
                {
                    case NodeIconKind.Format:
                        DrawListLines(g, bluePen);
                        break;
                    case NodeIconKind.Apply:
                        g.DrawLines(greenPen, new[] { new PointF(3F, 8F), new PointF(6.5F, 11.5F), new PointF(13F, 4.5F) });
                        break;
                    case NodeIconKind.ExpandAll:
                        DrawNodeBox(g, stroke, 2, 2);
                        DrawPlus(g, greenPen, 5F, 5F);
                        DrawNodeBox(g, stroke, 8, 8);
                        DrawPlus(g, greenPen, 11F, 11F);
                        break;
                    case NodeIconKind.CollapseAll:
                        DrawNodeBox(g, stroke, 2, 2);
                        DrawMinus(g, amberPen, 5F, 5F);
                        DrawNodeBox(g, stroke, 8, 8);
                        DrawMinus(g, amberPen, 11F, 11F);
                        break;
                    case NodeIconKind.AddSibling:
                        DrawNodeBox(g, stroke, 2, 3);
                        DrawNodeBox(g, stroke, 9, 3);
                        DrawPlus(g, greenPen, 8F, 12F);
                        break;
                    case NodeIconKind.AddChild:
                        DrawNodeBox(g, stroke, 2, 2);
                        g.DrawLine(stroke, 5F, 8F, 5F, 11F);
                        g.DrawLine(stroke, 5F, 11F, 9F, 11F);
                        DrawNodeBox(g, stroke, 9, 8);
                        DrawPlus(g, greenPen, 12F, 4F);
                        break;
                    case NodeIconKind.Edit:
                        g.DrawLine(bluePen, 4F, 12F, 11.5F, 4.5F);
                        g.DrawLine(bluePen, 9.5F, 3.5F, 12.5F, 6.5F);
                        g.FillRectangle(blueBrush, 3F, 12F, 4F, 1.6F);
                        break;
                    case NodeIconKind.Delete:
                        g.DrawLine(redPen, 4F, 4F, 12F, 12F);
                        g.DrawLine(redPen, 12F, 4F, 4F, 12F);
                        break;
                    case NodeIconKind.Hide:
                        g.DrawArc(stroke, 2, 5, 12, 6, 0, 180);
                        g.DrawArc(stroke, 2, 5, 12, 6, 180, 180);
                        g.FillEllipse(slateBrush, 6, 7, 4, 4);
                        g.DrawLine(redPen, 3F, 13F, 13F, 3F);
                        break;
                }
            }
            return bitmap;
        }

        private void DrawListLines(Graphics g, Pen pen)
        {
            g.DrawLine(pen, 4F, 4F, 12F, 4F);
            g.DrawLine(pen, 4F, 8F, 12F, 8F);
            g.DrawLine(pen, 4F, 12F, 10F, 12F);
        }

        private void DrawNodeBox(Graphics g, Pen pen, int x, int y)
        {
            g.DrawRectangle(pen, x, y, 5, 5);
        }

        private void DrawPlus(Graphics g, Pen pen, float centerX, float centerY)
        {
            g.DrawLine(pen, centerX - 2F, centerY, centerX + 2F, centerY);
            g.DrawLine(pen, centerX, centerY - 2F, centerX, centerY + 2F);
        }

        private void DrawMinus(Graphics g, Pen pen, float centerX, float centerY)
        {
            g.DrawLine(pen, centerX - 2F, centerY, centerX + 2F, centerY);
        }

        private Control CreateDetailPanel()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(12),
                BackColor = Color.White
            };
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
            panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            txtPath = CreateDetailTextBox(true);
            txtType = CreateDetailTextBox(true);
            txtValue = CreateDetailTextBox(false);
            txtValue.Multiline = true;
            txtValue.ScrollBars = ScrollBars.Vertical;

            panel.Controls.Add(CreateDetailLabel("节点路径"), 0, 0);
            panel.Controls.Add(txtPath, 0, 1);
            panel.Controls.Add(CreateDetailLabel("类型"), 0, 2);
            panel.Controls.Add(txtType, 0, 3);
            panel.Controls.Add(CreateDetailLabel("值"), 0, 4);
            panel.Controls.Add(txtValue, 0, 5);
            return panel;
        }

        private Label CreateDetailLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = Color.FromArgb(74, 85, 104),
                Margin = new Padding(0, 8, 0, 4)
            };
        }

        private TextBox CreateDetailTextBox(bool readOnly)
        {
            return new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = readOnly,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = readOnly ? Color.FromArgb(248, 250, 252) : Color.White,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular)
            };
        }

        private DataGridView CreateGrid()
        {
            var grid = new DataGridView
            {
                Dock = DockStyle.Fill,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersWidth = 36,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                EditMode = DataGridViewEditMode.EditOnEnter,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular)
            };
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Path", HeaderText = "节点路径", Width = 260, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = "类型", Width = 80, ReadOnly = true });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Value", HeaderText = "值", Width = 420 });
            grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249);
            grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 41, 55);
            grid.ColumnHeadersDefaultCellStyle.Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold);
            grid.EnableHeadersVisualStyles = false;
            return grid;
        }

        private void txtJson_TextChanged(object sender, EventArgs e)
        {
            if (internalChange) return;
            LoadTreeFromJson(null, null);
            OnJsonTextChanged();
        }

        private void LoadTreeFromJson()
        {
            LoadTreeFromJson(null, null);
        }

        private void LoadTreeFromJson(HashSet<string> expandedPaths, string selectedPath)
        {
            treeNodes.Nodes.Clear();
            ClearDetails();
            object root;
            if (!TryDeserialize(txtJson.Text, out root))
            {
                return;
            }
            TreeNode rootNode = CreateTreeNode("$", "$", root);
            treeNodes.Nodes.Add(rootNode);
            if (expandedPaths == null || expandedPaths.Count == 0)
            {
                rootNode.Expand();
            }
            else
            {
                RestoreExpandedPaths(expandedPaths);
            }
            if (!string.IsNullOrWhiteSpace(selectedPath))
            {
                SelectNodeByPath(selectedPath);
            }
        }

        private void ExpandAllNodes()
        {
            treeNodes.BeginUpdate();
            try
            {
                treeNodes.ExpandAll();
            }
            finally
            {
                treeNodes.EndUpdate();
            }
        }

        private void CollapseAllNodes()
        {
            treeNodes.BeginUpdate();
            try
            {
                treeNodes.CollapseAll();
                if (treeNodes.Nodes.Count > 0)
                {
                    treeNodes.Nodes[0].Expand();
                    treeNodes.SelectedNode = treeNodes.Nodes[0];
                }
            }
            finally
            {
                treeNodes.EndUpdate();
            }
        }

        private void SetJsonTextAfterNodeOperation(string json, string selectedPath, params string[] forceExpandedPaths)
        {
            HashSet<string> expandedPaths = GetExpandedPaths();
            foreach (string path in forceExpandedPaths ?? new string[0])
            {
                if (!string.IsNullOrWhiteSpace(path))
                {
                    expandedPaths.Add(path);
                }
            }
            SetJsonText(json, false, expandedPaths, selectedPath);
        }

        private HashSet<string> GetExpandedPaths()
        {
            var paths = new HashSet<string>(StringComparer.Ordinal);
            foreach (TreeNode node in treeNodes.Nodes)
            {
                CollectExpandedPaths(node, paths);
            }
            return paths;
        }

        private void CollectExpandedPaths(TreeNode node, HashSet<string> paths)
        {
            var info = node.Tag as JsonTreeNodeInfo;
            if (info != null && node.IsExpanded)
            {
                paths.Add(info.Path);
            }
            foreach (TreeNode child in node.Nodes)
            {
                CollectExpandedPaths(child, paths);
            }
        }

        private void RestoreExpandedPaths(HashSet<string> expandedPaths)
        {
            treeNodes.BeginUpdate();
            try
            {
                foreach (TreeNode node in treeNodes.Nodes)
                {
                    RestoreExpandedPaths(node, expandedPaths);
                }
                if (treeNodes.Nodes.Count > 0)
                {
                    treeNodes.Nodes[0].Expand();
                }
            }
            finally
            {
                treeNodes.EndUpdate();
            }
        }

        private void RestoreExpandedPaths(TreeNode node, HashSet<string> expandedPaths)
        {
            var info = node.Tag as JsonTreeNodeInfo;
            if (info != null && expandedPaths.Contains(info.Path))
            {
                node.Expand();
            }
            foreach (TreeNode child in node.Nodes)
            {
                RestoreExpandedPaths(child, expandedPaths);
            }
        }

        private bool SelectNodeByPath(string path)
        {
            foreach (TreeNode node in treeNodes.Nodes)
            {
                TreeNode found = FindNodeByPath(node, path);
                if (found != null)
                {
                    treeNodes.SelectedNode = found;
                    found.EnsureVisible();
                    ShowNodeDetails(found);
                    return true;
                }
            }
            return false;
        }

        private TreeNode FindNodeByPath(TreeNode node, string path)
        {
            var info = node.Tag as JsonTreeNodeInfo;
            if (info != null && info.Path == path)
            {
                return node;
            }
            foreach (TreeNode child in node.Nodes)
            {
                TreeNode found = FindNodeByPath(child, path);
                if (found != null) return found;
            }
            return null;
        }

        private void AddRows(string path, object value)
        {
        }

        private TreeNode CreateTreeNode(string label, string path, object value)
        {
            string type = GetJsonType(value);
            string nodeText = GetNodeDisplayText(label, type, value);

            var node = new TreeNode(nodeText)
            {
                Tag = new JsonTreeNodeInfo
                {
                    Path = path,
                    Type = type,
                    Value = value
                }
            };

            var dictionary = value as IDictionary<string, object>;
            if (dictionary != null)
            {
                foreach (KeyValuePair<string, object> pair in dictionary)
                {
                    node.Nodes.Add(CreateTreeNode(pair.Key, path + "." + pair.Key, pair.Value));
                }
                return node;
            }

            foreach (IndexedValue item in EnumerateArray(value))
            {
                node.Nodes.Add(CreateTreeNode("[" + item.Index.ToString(CultureInfo.InvariantCulture) + "]", path + "[" + item.Index.ToString(CultureInfo.InvariantCulture) + "]", item.Value));
            }

            return node;
        }

        private IEnumerable<IndexedValue> EnumerateArray(object value)
        {
            var arrayList = value as ArrayList;
            if (arrayList != null)
            {
                for (int i = 0; i < arrayList.Count; i++)
                {
                    yield return new IndexedValue { Index = i, Value = arrayList[i] };
                }
                yield break;
            }

            var objectArray = value as object[];
            if (objectArray != null)
            {
                for (int i = 0; i < objectArray.Length; i++)
                {
                    yield return new IndexedValue { Index = i, Value = objectArray[i] };
                }
            }
        }

        private string GetNodeDisplayText(string label, string type, object value)
        {
            if (type == "object")
            {
                string summary = GetObjectSummary(value);
                return string.IsNullOrWhiteSpace(summary) ? label : label + "  " + summary;
            }

            if (type == "array")
            {
                return label + " (" + GetArrayCount(value).ToString(CultureInfo.InvariantCulture) + "项)";
            }

            return label + " = " + JsonValueToText(value);
        }

        private string GetObjectSummary(object value)
        {
            var dict = value as IDictionary<string, object>;
            if (dict == null) return string.Empty;
            foreach (string key in new[] { "name", "text", "title", "label", "id" })
            {
                if (dict.ContainsKey(key) && dict[key] != null)
                {
                    string text = JsonValueToText(dict[key]);
                    if (!string.IsNullOrWhiteSpace(text)) return text;
                }
            }
            return string.Empty;
        }

        private int GetArrayCount(object value)
        {
            var arrayList = value as ArrayList;
            if (arrayList != null) return arrayList.Count;
            var objectArray = value as object[];
            return objectArray == null ? 0 : objectArray.Length;
        }

        private void treeNodes_AfterSelect(object sender, TreeViewEventArgs e)
        {
            ShowNodeDetails(e.Node);
        }

        private void treeNodes_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                treeNodes.SelectedNode = e.Node;
            }
        }

        private void ShowNodeDetails(TreeNode node)
        {
            var info = node == null ? null : node.Tag as JsonTreeNodeInfo;
            if (info == null)
            {
                ClearDetails();
                return;
            }

            txtPath.Text = info.Path;
            txtType.Text = info.Type;
            txtValue.Text = JsonValueToText(info.Value);
            txtValue.ReadOnly = info.Type == "object" || info.Type == "array";
            txtValue.BackColor = txtValue.ReadOnly ? Color.FromArgb(248, 250, 252) : Color.White;
        }

        private void ClearDetails()
        {
            if (txtPath == null) return;
            txtPath.Text = string.Empty;
            txtType.Text = string.Empty;
            txtValue.Text = string.Empty;
            txtValue.ReadOnly = true;
            txtValue.BackColor = Color.FromArgb(248, 250, 252);
        }

        private void ApplySelectedNodeToJson()
        {
            var info = treeNodes.SelectedNode == null ? null : treeNodes.SelectedNode.Tag as JsonTreeNodeInfo;
            if (info == null || info.Type == "object" || info.Type == "array")
            {
                MessageBox.Show("请选择可编辑的叶子节点。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            object root;
            if (!TryDeserialize(txtJson.Text, out root))
            {
                MessageBox.Show("JSON格式不正确，无法应用节点修改。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SetValueByPath(root, info.Path, ParseJsonValue(info.Type, txtValue.Text));
            SetJsonTextAfterNodeOperation(SerializePretty(root), info.Path, GetParentPath(info.Path));
        }

        private void AddSiblingNode()
        {
            var info = GetSelectedInfo();
            if (info == null || info.Path == "$")
            {
                MessageBox.Show("根节点不能新增同级节点。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            object root;
            if (!TryDeserialize(txtJson.Text, out root)) return;
            object parent = GetParentByTokens(root, ParsePath(info.Path));
            bool needName = parent is IDictionary<string, object>;

            string name;
            string value;
            if (!PromptNodeInput("新增同级节点", needName, out name, out value)) return;

            if (AddSiblingValue(root, info.Path, name, value))
            {
                string parentPath = GetParentPath(info.Path);
                SetJsonTextAfterNodeOperation(SerializePretty(root), info.Path, parentPath);
            }
        }

        private void AddChildNode()
        {
            var info = GetSelectedInfo();
            if (info == null) return;
            if (info.Type != "object" && info.Type != "array")
            {
                MessageBox.Show("只有 object 或 array 节点可以新增子节点。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string name;
            string value;
            bool needName = info.Type == "object";
            if (!PromptNodeInput("新增子节点", needName, out name, out value)) return;
            object root;
            if (!TryDeserialize(txtJson.Text, out root)) return;

            object parent = GetValueByPath(root, info.Path);
            if (AddChildValue(parent, name, value))
            {
                SetJsonTextAfterNodeOperation(SerializePretty(root), info.Path, info.Path);
            }
            else if (parent is object[])
            {
                var newList = ((object[])parent).ToList();
                newList.Add(CreateAlignedValueForArray(newList, value));
                if (ReplaceValueByPath(root, info.Path, newList.ToArray()))
                {
                    SetJsonTextAfterNodeOperation(SerializePretty(root), info.Path, info.Path);
                }
            }
        }

        private void DeleteSelectedNode()
        {
            var info = GetSelectedInfo();
            if (info == null || info.Path == "$")
            {
                MessageBox.Show("根节点不能删除。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (MessageBox.Show("确定删除当前节点吗？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
            {
                return;
            }

            object root;
            if (!TryDeserialize(txtJson.Text, out root)) return;
            if (RemoveValueByPath(root, info.Path))
            {
                string parentPath = GetParentPath(info.Path);
                SetJsonTextAfterNodeOperation(SerializePretty(root), parentPath, parentPath);
            }
        }

        private void HideSelectedNode()
        {
            var info = GetSelectedInfo();
            if (info == null || info.Path == "$")
            {
                MessageBox.Show("根节点不能隐藏。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            object root;
            if (!TryDeserialize(txtJson.Text, out root)) return;
            object current = GetValueByPath(root, info.Path);
            var dict = current as IDictionary<string, object>;
            if (dict != null)
            {
                dict["hidden"] = true;
                SetJsonTextAfterNodeOperation(SerializePretty(root), info.Path, GetParentPath(info.Path));
            }
            else
            {
                RemoveValueByPath(root, info.Path);
                string parentPath = GetParentPath(info.Path);
                SetJsonTextAfterNodeOperation(SerializePretty(root), parentPath, parentPath);
            }
        }

        private JsonTreeNodeInfo GetSelectedInfo()
        {
            return treeNodes.SelectedNode == null ? null : treeNodes.SelectedNode.Tag as JsonTreeNodeInfo;
        }

        private bool PromptNodeInput(string title, bool requireName, out string name, out string value)
        {
            using (var dialog = new JsonNodePromptDialog(title, requireName))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    name = string.Empty;
                    value = string.Empty;
                    return false;
                }
                name = dialog.NodeName;
                value = dialog.NodeValue;
                return true;
            }
        }

        private void ApplyGridToJson()
        {
            object root;
            if (!TryDeserialize(txtJson.Text, out root))
            {
                MessageBox.Show("JSON格式不正确，无法应用表格修改。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (DataGridViewRow row in new DataGridViewRow[0])
            {
                if (row.IsNewRow) continue;
                string path = Convert.ToString(row.Cells["Path"].Value);
                string type = Convert.ToString(row.Cells["Type"].Value);
                string valueText = Convert.ToString(row.Cells["Value"].Value);
                SetValueByPath(root, path, ParseJsonValue(type, valueText));
            }

            SetJsonTextAfterNodeOperation(SerializePretty(root), "$", "$");
        }

        private bool AddSiblingValue(object root, string path, string name, string value)
        {
            List<PathToken> tokens = ParsePath(path);
            if (tokens.Count == 0) return false;
            PathToken last = tokens[tokens.Count - 1];
            object parent = GetParentByTokens(root, tokens);
            if (parent == null) return false;

            var dict = parent as IDictionary<string, object>;
            if (dict != null)
            {
                string key = string.IsNullOrWhiteSpace(name) ? MakeUniqueKey(dict, "newNode") : MakeUniqueKey(dict, name.Trim());
                dict[key] = value ?? string.Empty;
                return true;
            }

            var list = parent as ArrayList;
            if (list != null && last.Kind == PathTokenKind.Index)
            {
                list.Insert(Math.Min(last.Index + 1, list.Count), CreateAlignedValueForArray(list, value));
                return true;
            }

            var array = parent as object[];
            if (array != null && last.Kind == PathTokenKind.Index)
            {
                var newList = array.ToList();
                newList.Insert(Math.Min(last.Index + 1, newList.Count), CreateAlignedValueForArray(newList, value));
                return ReplaceArrayInParent(root, tokens, newList.ToArray());
            }

            return false;
        }

        private bool AddChildValue(object parent, string name, string value)
        {
            var dict = parent as IDictionary<string, object>;
            if (dict != null)
            {
                string key = string.IsNullOrWhiteSpace(name) ? MakeUniqueKey(dict, "newNode") : MakeUniqueKey(dict, name.Trim());
                dict[key] = value ?? string.Empty;
                return true;
            }

            var list = parent as ArrayList;
            if (list != null)
            {
                list.Add(CreateAlignedValueForArray(list, value));
                return true;
            }

            var array = parent as object[];
            if (array != null)
            {
                return false;
            }

            return false;
        }

        private object CreateAlignedValueForArray(IEnumerable items, string displayValue)
        {
            foreach (object item in items)
            {
                if (item is IDictionary<string, object>)
                {
                    return CreateValueFromTemplate(item, displayValue);
                }
            }
            return displayValue ?? string.Empty;
        }

        private object CreateValueFromTemplate(object template, string displayValue)
        {
            var dict = template as IDictionary<string, object>;
            if (dict != null)
            {
                var result = new Dictionary<string, object>();
                foreach (KeyValuePair<string, object> pair in dict)
                {
                    if (pair.Key == "hidden")
                    {
                        continue;
                    }
                    if (IsDisplayField(pair.Key))
                    {
                        result[pair.Key] = displayValue ?? string.Empty;
                    }
                    else
                    {
                        result[pair.Key] = CreateEmptyValueLike(pair.Value);
                    }
                }
                if (!result.Keys.Any(IsDisplayField))
                {
                    result["name"] = displayValue ?? string.Empty;
                }
                return result;
            }

            return displayValue ?? string.Empty;
        }

        private object CreateEmptyValueLike(object value)
        {
            if (value is IDictionary<string, object>)
            {
                return CreateValueFromTemplate(value, string.Empty);
            }
            if (value is ArrayList || value is object[])
            {
                return new object[0];
            }
            if (value is bool)
            {
                return false;
            }
            if (value is int || value is long || value is decimal || value is double || value is float)
            {
                return 0;
            }
            return string.Empty;
        }

        private bool IsDisplayField(string key)
        {
            return string.Equals(key, "name", StringComparison.OrdinalIgnoreCase)
                || string.Equals(key, "text", StringComparison.OrdinalIgnoreCase)
                || string.Equals(key, "title", StringComparison.OrdinalIgnoreCase)
                || string.Equals(key, "label", StringComparison.OrdinalIgnoreCase);
        }

        private bool RemoveValueByPath(object root, string path)
        {
            List<PathToken> tokens = ParsePath(path);
            if (tokens.Count == 0) return false;
            PathToken last = tokens[tokens.Count - 1];
            object parent = GetParentByTokens(root, tokens);
            if (parent == null) return false;

            var dict = parent as IDictionary<string, object>;
            if (dict != null && last.Kind == PathTokenKind.Property)
            {
                return dict.Remove(last.Name);
            }

            var list = parent as ArrayList;
            if (list != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < list.Count)
            {
                list.RemoveAt(last.Index);
                return true;
            }

            var array = parent as object[];
            if (array != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < array.Length)
            {
                var newList = array.ToList();
                newList.RemoveAt(last.Index);
                return ReplaceArrayInParent(root, tokens, newList.ToArray());
            }

            return false;
        }

        private bool ReplaceValueByPath(object root, string path, object value)
        {
            if (path == "$") return false;
            List<PathToken> tokens = ParsePath(path);
            if (tokens.Count == 0) return false;
            PathToken last = tokens[tokens.Count - 1];
            object parent = GetParentByTokens(root, tokens);
            if (parent == null) return false;

            var dict = parent as IDictionary<string, object>;
            if (dict != null && last.Kind == PathTokenKind.Property)
            {
                dict[last.Name] = value;
                return true;
            }

            var list = parent as ArrayList;
            if (list != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < list.Count)
            {
                list[last.Index] = value;
                return true;
            }

            var array = parent as object[];
            if (array != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < array.Length)
            {
                array[last.Index] = value;
                return true;
            }

            return false;
        }

        private object GetValueByPath(object root, string path)
        {
            object current = root;
            foreach (PathToken token in ParsePath(path))
            {
                current = GetChild(current, token);
                if (current == null) return null;
            }
            return current;
        }

        private object GetParentByTokens(object root, List<PathToken> tokens)
        {
            object current = root;
            for (int i = 0; i < tokens.Count - 1; i++)
            {
                current = GetChild(current, tokens[i]);
                if (current == null) return null;
            }
            return current;
        }

        private bool ReplaceArrayInParent(object root, List<PathToken> targetTokens, object[] newArray)
        {
            if (targetTokens.Count == 0) return false;
            List<PathToken> parentTokens = targetTokens.Take(targetTokens.Count - 1).ToList();
            object grandParent = GetParentByTokens(root, parentTokens);
            if (grandParent == null && parentTokens.Count == 0)
            {
                return false;
            }

            if (parentTokens.Count == 0)
            {
                return false;
            }

            PathToken arrayToken = parentTokens[parentTokens.Count - 1];
            var dict = grandParent as IDictionary<string, object>;
            if (dict != null && arrayToken.Kind == PathTokenKind.Property)
            {
                dict[arrayToken.Name] = newArray;
                return true;
            }

            var list = grandParent as ArrayList;
            if (list != null && arrayToken.Kind == PathTokenKind.Index && arrayToken.Index >= 0 && arrayToken.Index < list.Count)
            {
                list[arrayToken.Index] = newArray;
                return true;
            }

            var array = grandParent as object[];
            if (array != null && arrayToken.Kind == PathTokenKind.Index && arrayToken.Index >= 0 && arrayToken.Index < array.Length)
            {
                array[arrayToken.Index] = newArray;
                return true;
            }

            return false;
        }

        private string MakeUniqueKey(IDictionary<string, object> dict, string baseName)
        {
            string key = string.IsNullOrWhiteSpace(baseName) ? "newNode" : baseName;
            if (!dict.ContainsKey(key)) return key;
            int i = 2;
            while (dict.ContainsKey(key + i.ToString(CultureInfo.InvariantCulture)))
            {
                i++;
            }
            return key + i.ToString(CultureInfo.InvariantCulture);
        }

        private void SetValueByPath(object root, string path, object value)
        {
            List<PathToken> tokens = ParsePath(path);
            if (tokens.Count == 0) return;

            object current = root;
            for (int i = 0; i < tokens.Count - 1; i++)
            {
                current = GetChild(current, tokens[i]);
                if (current == null) return;
            }

            PathToken last = tokens[tokens.Count - 1];
            var dict = current as IDictionary<string, object>;
            if (dict != null && last.Kind == PathTokenKind.Property)
            {
                dict[last.Name] = value;
                return;
            }

            var list = current as ArrayList;
            if (list != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < list.Count)
            {
                list[last.Index] = value;
                return;
            }

            var array = current as object[];
            if (array != null && last.Kind == PathTokenKind.Index && last.Index >= 0 && last.Index < array.Length)
            {
                array[last.Index] = value;
            }
        }

        private object GetChild(object current, PathToken token)
        {
            var dict = current as IDictionary<string, object>;
            if (dict != null && token.Kind == PathTokenKind.Property && dict.ContainsKey(token.Name))
            {
                return dict[token.Name];
            }

            var list = current as ArrayList;
            if (list != null && token.Kind == PathTokenKind.Index && token.Index >= 0 && token.Index < list.Count)
            {
                return list[token.Index];
            }

            var array = current as object[];
            if (array != null && token.Kind == PathTokenKind.Index && token.Index >= 0 && token.Index < array.Length)
            {
                return array[token.Index];
            }
            return null;
        }

        private List<PathToken> ParsePath(string path)
        {
            var tokens = new List<PathToken>();
            if (string.IsNullOrWhiteSpace(path) || !path.StartsWith("$")) return tokens;

            int i = 1;
            while (i < path.Length)
            {
                if (path[i] == '.')
                {
                    i++;
                    int start = i;
                    while (i < path.Length && path[i] != '.' && path[i] != '[') i++;
                    tokens.Add(new PathToken { Kind = PathTokenKind.Property, Name = path.Substring(start, i - start) });
                }
                else if (path[i] == '[')
                {
                    int end = path.IndexOf(']', i);
                    if (end < 0) break;
                    int index;
                    if (int.TryParse(path.Substring(i + 1, end - i - 1), out index))
                    {
                        tokens.Add(new PathToken { Kind = PathTokenKind.Index, Index = index });
                    }
                    i = end + 1;
                }
                else
                {
                    i++;
                }
            }
            return tokens;
        }

        private string GetParentPath(string path)
        {
            List<PathToken> tokens = ParsePath(path);
            if (tokens.Count == 0) return "$";
            tokens.RemoveAt(tokens.Count - 1);
            return BuildPath(tokens);
        }

        private string BuildPath(IEnumerable<PathToken> tokens)
        {
            var path = new System.Text.StringBuilder("$");
            foreach (PathToken token in tokens)
            {
                if (token.Kind == PathTokenKind.Property)
                {
                    path.Append(".");
                    path.Append(token.Name);
                }
                else
                {
                    path.Append("[");
                    path.Append(token.Index.ToString(CultureInfo.InvariantCulture));
                    path.Append("]");
                }
            }
            return path.ToString();
        }

        private bool TryDeserialize(string json, out object root)
        {
            root = null;
            if (string.IsNullOrWhiteSpace(json)) return false;
            try
            {
                root = new JavaScriptSerializer().DeserializeObject(json);
                return root != null;
            }
            catch
            {
                return false;
            }
        }

        private string TryFormatJson(string json)
        {
            object root;
            if (!TryDeserialize(json, out root))
            {
                return json ?? string.Empty;
            }
            return SerializePretty(root);
        }

        private string SerializePretty(object root)
        {
            string compact = new JavaScriptSerializer().Serialize(root);
            return PrettyPrintJson(compact);
        }

        private string PrettyPrintJson(string json)
        {
            int indent = 0;
            bool inString = false;
            bool escape = false;
            var output = new System.Text.StringBuilder();
            foreach (char c in json)
            {
                if (escape)
                {
                    output.Append(c);
                    escape = false;
                    continue;
                }
                if (c == '\\' && inString)
                {
                    output.Append(c);
                    escape = true;
                    continue;
                }
                if (c == '"')
                {
                    inString = !inString;
                    output.Append(c);
                    continue;
                }
                if (inString)
                {
                    output.Append(c);
                    continue;
                }
                if (c == '{' || c == '[')
                {
                    output.Append(c);
                    output.AppendLine();
                    indent++;
                    output.Append(new string(' ', indent * 2));
                }
                else if (c == '}' || c == ']')
                {
                    output.AppendLine();
                    indent = Math.Max(0, indent - 1);
                    output.Append(new string(' ', indent * 2));
                    output.Append(c);
                }
                else if (c == ',')
                {
                    output.Append(c);
                    output.AppendLine();
                    output.Append(new string(' ', indent * 2));
                }
                else if (c == ':')
                {
                    output.Append(": ");
                }
                else if (!char.IsWhiteSpace(c))
                {
                    output.Append(c);
                }
            }
            return output.ToString();
        }

        private string GetJsonType(object value)
        {
            if (value == null) return "null";
            if (value is IDictionary<string, object>) return "object";
            if (value is ArrayList || value is object[]) return "array";
            if (value is bool) return "bool";
            if (value is int || value is long || value is decimal || value is double || value is float) return "number";
            return "string";
        }

        private string JsonValueToText(object value)
        {
            if (value == null) return string.Empty;
            if (value is bool) return ((bool)value) ? "true" : "false";
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private object ParseJsonValue(string type, string value)
        {
            if (type == "null") return null;
            if (type == "bool")
            {
                bool parsed;
                return bool.TryParse(value, out parsed) && parsed;
            }
            if (type == "number")
            {
                decimal parsed;
                if (decimal.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
                {
                    return parsed;
                }
                return 0;
            }
            return value ?? string.Empty;
        }

        private void OnJsonTextChanged()
        {
            JsonTextChanged?.Invoke(this, EventArgs.Empty);
        }

        private enum PathTokenKind
        {
            Property,
            Index
        }

        private enum NodeIconKind
        {
            Format,
            Apply,
            ExpandAll,
            CollapseAll,
            AddSibling,
            AddChild,
            Edit,
            Delete,
            Hide
        }

        private class PathToken
        {
            public PathTokenKind Kind { get; set; }
            public string Name { get; set; }
            public int Index { get; set; }
        }

        private class JsonTreeNodeInfo
        {
            public string Path { get; set; }
            public string Type { get; set; }
            public object Value { get; set; }
        }

        private class IndexedValue
        {
            public int Index { get; set; }
            public object Value { get; set; }
        }

        private class JsonNodePromptDialog : Form
        {
            private readonly TextBox txtName;
            private readonly TextBox txtValue;

            public JsonNodePromptDialog(string title, bool requireName)
            {
                Text = title;
                Size = new Size(360, requireName ? 205 : 165);
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular);

                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 1,
                    Padding = new Padding(14),
                    BackColor = Color.White
                };

                txtName = new TextBox { Dock = DockStyle.Top, Enabled = requireName };
                txtValue = new TextBox { Dock = DockStyle.Top };

                if (requireName)
                {
                    layout.Controls.Add(new Label { Text = "节点名称", AutoSize = true, Margin = new Padding(0, 0, 0, 4) });
                    layout.Controls.Add(txtName);
                }
                layout.Controls.Add(new Label { Text = "节点值", AutoSize = true, Margin = new Padding(0, 10, 0, 4) });
                layout.Controls.Add(txtValue);

                var buttons = new FlowLayoutPanel
                {
                    Dock = DockStyle.Bottom,
                    FlowDirection = FlowDirection.RightToLeft,
                    Height = 38,
                    Padding = new Padding(0, 8, 0, 0)
                };
                var btnOk = new Button { Text = "确定", DialogResult = DialogResult.OK, Width = 76 };
                var btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Width = 76 };
                buttons.Controls.Add(btnOk);
                buttons.Controls.Add(btnCancel);
                layout.Controls.Add(buttons);

                AcceptButton = btnOk;
                CancelButton = btnCancel;
                Controls.Add(layout);
            }

            public string NodeName
            {
                get { return txtName.Text.Trim(); }
            }

            public string NodeValue
            {
                get { return txtValue.Text; }
            }
        }
    }
}
