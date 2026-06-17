using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioDataMapper
{
    public class DfdSugiyamaLayout
    {
        public class LayoutNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
            public int Layer { get; set; } = -1;
            public double Barycenter { get; set; } = 0.0;
        }

        public class LayoutEdge
        {
            public int From { get; set; }
            public int To { get; set; }
        }

        /// <summary>
        /// 执行 Sugiyama 分层排版算法，并计算出每个节点的绝对坐标（以英寸为单位）
        /// </summary>
        /// <param name="nodes">节点列表</param>
        /// <param name="edges">连接边列表</param>
        /// <param name="pageWidth">画布宽度（英寸）</param>
        /// <param name="pageHeight">画布高度（英寸）</param>
        /// <param name="horGap">水平间距（英寸）</param>
        /// <param name="verGap">垂直间距（英寸）</param>
        /// <param name="direction">排版流向: "TB" (自上而下) 或 "LR" (自左而右)</param>
        /// <returns>节点 Index 对应的 X, Y 坐标字典</returns>
        public Dictionary<int, Tuple<double, double>> CalculateLayout(
            List<LayoutNode> nodes,
            List<LayoutEdge> edges,
            double pageWidth,
            double pageHeight,
            double horGap,
            double verGap,
            string direction = "LR")
        {
            if (nodes == null || nodes.Count == 0)
                return new Dictionary<int, Tuple<double, double>>();

            var nodeMap = nodes.ToDictionary(n => n.Index);

            // 1. 分层 (Layer Assignment) - 拓扑层级递推（处理可能存在的环，防死循环）
            AssignLayers(nodes, edges, nodeMap);

            // 2. 补齐未分层节点并按层级归类
            var layers = GroupNodesIntoLayers(nodes);

            // 3. 层内排序 (Node Ordering / Crossing Reduction) - 重心启发式
            OrderNodesInLayers(layers, edges);

            // 4. 计算并分配具体坐标 (Coordinate Assignment)
            var coordinates = AssignCoordinates(layers, pageWidth, pageHeight, horGap, verGap, direction);

            return coordinates;
        }

        private void AssignLayers(List<LayoutNode> nodes, List<LayoutEdge> edges, Dictionary<int, LayoutNode> nodeMap)
        {
            // 初始化入度计数
            var inDegree = nodes.ToDictionary(n => n.Index, n => 0);
            var adj = nodes.ToDictionary(n => n.Index, n => new List<int>());

            foreach (var edge in edges)
            {
                if (inDegree.ContainsKey(edge.To)) inDegree[edge.To]++;
                if (adj.ContainsKey(edge.From)) adj[edge.From].Add(edge.To);
            }

            // 设定初始层
            var queue = new Queue<int>();
            foreach (var node in nodes)
            {
                if (inDegree[node.Index] == 0)
                {
                    node.Layer = 0;
                    queue.Enqueue(node.Index);
                }
            }

            // 如果图中有强连通环导致没有任何入度为 0 的节点，则强行选第一个节点为 0 层
            if (queue.Count == 0 && nodes.Count > 0)
            {
                var firstNode = nodes[0];
                firstNode.Layer = 0;
                queue.Enqueue(firstNode.Index);
            }

            // 层级向前递推（Sugiyama 经典最长路径分层变种，支持轻度循环）
            int maxIterations = nodes.Count * 2;
            int iter = 0;
            while (queue.Count > 0 && iter < maxIterations)
            {
                iter++;
                int currIdx = queue.Dequeue();
                var currNode = nodeMap[currIdx];

                foreach (var neighborIdx in adj[currIdx])
                {
                    var neighbor = nodeMap[neighborIdx];
                    int newLayer = currNode.Layer + 1;
                    if (newLayer > neighbor.Layer)
                    {
                        neighbor.Layer = newLayer;
                        queue.Enqueue(neighborIdx);
                    }
                }
            }

            // 对于任何仍然未被分层的孤立/成环节点，强制给它们分配合理的层级，防止抛错
            foreach (var node in nodes)
            {
                if (node.Layer < 0)
                {
                    node.Layer = 0;
                }
            }
        }

        private List<List<LayoutNode>> GroupNodesIntoLayers(List<LayoutNode> nodes)
        {
            int maxLayer = nodes.Max(n => n.Layer);
            var layers = new List<List<LayoutNode>>();
            for (int i = 0; i <= maxLayer; i++)
            {
                layers.Add(new List<LayoutNode>());
            }

            foreach (var node in nodes)
            {
                layers[node.Layer].Add(node);
            }

            // 移除可能存在的空层
            return layers.Where(l => l.Count > 0).ToList();
        }

        private void OrderNodesInLayers(List<List<LayoutNode>> layers, List<LayoutEdge> edges)
        {
            // 为每层的节点赋予一个初始顺序值
            for (int l = 0; l < layers.Count; l++)
            {
                for (int i = 0; i < layers[l].Count; i++)
                {
                    layers[l][i].Barycenter = i;
                }
            }

            // 使用重心排序法进行两轮扫掠，减少连线交叉率
            for (int sweep = 0; sweep < 2; sweep++)
            {
                // 正向扫掠：从第一层到最后一层
                for (int l = 1; l < layers.Count; l++)
                {
                    var prevLayerMap = layers[l - 1].Select((n, idx) => new { n.Index, idx }).ToDictionary(x => x.Index, x => x.idx);

                    foreach (var node in layers[l])
                    {
                        // 寻找前驱节点（指向当前节点的父节点）
                        var parents = edges.Where(e => e.To == node.Index && prevLayerMap.ContainsKey(e.From)).ToList();
                        if (parents.Count > 0)
                        {
                            node.Barycenter = parents.Average(e => prevLayerMap[e.From]);
                        }
                    }

                    // 排序该层
                    layers[l] = layers[l].OrderBy(n => n.Barycenter).ToList();
                }
            }
        }

        private Dictionary<int, Tuple<double, double>> AssignCoordinates(
            List<List<LayoutNode>> layers,
            double pageWidth,
            double pageHeight,
            double horGap,
            double verGap,
            string direction)
        {
            var coords = new Dictionary<int, Tuple<double, double>>();

            // 假设默认单形状标准大小（以便做避让定位计算）
            double nodeW = 1.2;
            double nodeH = 0.8;

            if (direction == "TB") // 自上而下 (Top-to-Bottom)
            {
                double startY = pageHeight - 1.5;
                for (int l = 0; l < layers.Count; l++)
                {
                    var layerNodes = layers[l];
                    int count = layerNodes.Count;
                    double layerY = startY - l * (nodeH + verGap);

                    // 计算该层占用的总宽度以实现中线对齐居中
                    double totalWidth = count * nodeW + (count - 1) * horGap;
                    double startX = (pageWidth - totalWidth) / 2.0;
                    if (startX < 0.8) startX = 0.8; // 设定页边距安全保护下限

                    for (int i = 0; i < count; i++)
                    {
                        double nodeX = startX + i * (nodeW + horGap);
                        coords[layerNodes[i].Index] = new Tuple<double, double>(nodeX, layerY);
                    }
                }
            }
            else // 自左而右 (Left-to-Right) - 绝大多数复杂数据流图的经典方向
            {
                double startX = 1.2;
                for (int l = 0; l < layers.Count; l++)
                {
                    var layerNodes = layers[l];
                    int count = layerNodes.Count;
                    double layerX = startX + l * (nodeW + horGap);

                    // 计算该层占用的总高度以实现中线垂直居中对齐
                    double totalHeight = count * nodeH + (count - 1) * verGap;
                    double startY = pageHeight - (pageHeight - totalHeight) / 2.0;
                    if (startY > pageHeight - 0.8) startY = pageHeight - 0.8; // 设定顶部页边距保护上限

                    for (int i = 0; i < count; i++)
                    {
                        double nodeY = startY - i * (nodeH + verGap);
                        coords[layerNodes[i].Index] = new Tuple<double, double>(layerX, nodeY);
                    }
                }
            }

            return coords;
        }
    }
}
