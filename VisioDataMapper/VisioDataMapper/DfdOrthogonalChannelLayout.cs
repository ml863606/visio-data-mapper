using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioDataMapper
{
    public class DfdOrthogonalChannelLayout
    {
        public class LayoutNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
        }

        public class LayoutEdge
        {
            public int From { get; set; }
            public int To { get; set; }
            public string Text { get; set; }
        }

        public class LayoutPoint
        {
            public double X { get; set; }
            public double Y { get; set; }

            public LayoutPoint(double x, double y)
            {
                X = x;
                Y = y;
            }
        }

        public class RoutedEdge
        {
            public int From { get; set; }
            public int To { get; set; }
            public string Text { get; set; }
            public List<LayoutPoint> Points { get; set; } = new List<LayoutPoint>();
            public LayoutPoint LabelPoint { get; set; }
        }

        public class SubDiagramPlan
        {
            public int FocusNodeIndex { get; set; }
            public string PageName { get; set; }
        }

        public class LayoutResult
        {
            public double PageWidth { get; set; }
            public double PageHeight { get; set; }
            public Dictionary<int, LayoutPoint> Nodes { get; set; } = new Dictionary<int, LayoutPoint>();
            public List<RoutedEdge> Edges { get; set; } = new List<RoutedEdge>();
            public List<SubDiagramPlan> SubDiagrams { get; set; } = new List<SubDiagramPlan>();
        }

        private const double NodeWidth = 1.35;
        private const double NodeHeight = 0.9;
        private const double Margin = 1.2;
        private const double EntityLaneWidth = 2.2;
        private const double StoreLaneHeight = 2.0;
        private const double TopChannelPadding = 1.35;
        private const double BottomChannelPadding = 1.0;
        private const double ChannelGap = 0.22;
        private const int MaxEdgesPerOverviewProcess = 8;

        public LayoutResult CalculateLayout(
            List<LayoutNode> nodes,
            List<LayoutEdge> edges,
            double pageWidth,
            double pageHeight,
            double horGap,
            double verGap)
        {
            var result = new LayoutResult();
            if (nodes == null || nodes.Count == 0)
            {
                result.PageWidth = pageWidth;
                result.PageHeight = pageHeight;
                return result;
            }

            edges = edges ?? new List<LayoutEdge>();

            var validIds = new HashSet<int>(nodes.Select(n => n.Index));
            edges = edges.Where(e => validIds.Contains(e.From) && validIds.Contains(e.To) && e.From != e.To).ToList();

            var entities = nodes.Where(IsEntity).OrderBy(n => n.Index).ToList();
            var stores = nodes.Where(IsStore).OrderBy(n => n.Index).ToList();
            var processes = nodes.Where(n => !IsEntity(n) && !IsStore(n)).OrderBy(n => n.Index).ToList();
            if (processes.Count == 0)
            {
                processes = nodes.Where(n => !IsEntity(n)).OrderBy(n => n.Index).ToList();
            }

            var processLayers = BuildProcessLayers(processes, edges);
            int layerCount = Math.Max(1, processLayers.Count);
            int maxProcessRows = processLayers.Count == 0 ? 1 : Math.Max(1, processLayers.Max(l => l.Count));
            int leftEntityCount = (entities.Count + 1) / 2;
            int rightEntityCount = entities.Count - leftEntityCount;
            int maxEntityRows = Math.Max(leftEntityCount, rightEntityCount);
            int routeCount = Math.Max(1, edges.Count);

            double processLayerGap = Math.Max(2.4, horGap + 1.2);
            double rowGap = Math.Max(1.45, verGap + 0.55);
            double storeGap = Math.Max(2.0, horGap + 0.9);

            double processWidth = Math.Max(1, layerCount - 1) * processLayerGap + NodeWidth;
            double storeWidth = stores.Count <= 1 ? NodeWidth : (stores.Count - 1) * storeGap + NodeWidth;
            double processHeight = Math.Max(1, maxProcessRows - 1) * rowGap + NodeHeight;
            double entityHeight = maxEntityRows <= 1 ? NodeHeight : (maxEntityRows - 1) * rowGap + NodeHeight;
            double channelHeight = TopChannelPadding + BottomChannelPadding + routeCount * ChannelGap;

            double requiredWidth = Margin * 2 + EntityLaneWidth * 2 + Math.Max(processWidth, storeWidth);
            double requiredHeight = Margin * 2 + StoreLaneHeight + Math.Max(processHeight, entityHeight) + channelHeight;
            result.PageWidth = Math.Max(pageWidth, requiredWidth);
            result.PageHeight = Math.Max(pageHeight, requiredHeight);

            double processLeft = Margin + EntityLaneWidth;
            double processRight = result.PageWidth - Margin - EntityLaneWidth;
            double storeY = Margin + 0.75;
            double processCenterY = StoreLaneHeight + Margin + BottomChannelPadding + routeCount * ChannelGap + processHeight / 2.0;
            processCenterY = Math.Min(result.PageHeight - Margin - TopChannelPadding - routeCount * ChannelGap, processCenterY);

            PlaceProcesses(processLayers, result.Nodes, processLeft, processRight, processCenterY, rowGap);
            PlaceEntities(entities, edges, result.Nodes, leftEntityCount, result.PageWidth, processCenterY, rowGap);
            PlaceStores(stores, edges, result.Nodes, processLeft, processRight, storeY, storeGap);

            foreach (var node in nodes)
            {
                if (!result.Nodes.ContainsKey(node.Index))
                {
                    result.Nodes[node.Index] = new LayoutPoint((processLeft + processRight) / 2.0, processCenterY);
                }
            }

            var routedEdges = new List<RoutedEdge>();
            for (int i = 0; i < edges.Count; i++)
            {
                routedEdges.Add(RouteEdge(edges[i], i, result.Nodes, result.PageHeight));
            }

            result.Edges = routedEdges;
            result.SubDiagrams = BuildSubDiagramPlans(processes, edges);
            return result;
        }

        private static bool IsEntity(LayoutNode node)
        {
            return string.Equals(node.Type, "实体", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsStore(LayoutNode node)
        {
            return string.Equals(node.Type, "数据存储", StringComparison.OrdinalIgnoreCase);
        }

        private static List<List<LayoutNode>> BuildProcessLayers(List<LayoutNode> processes, List<LayoutEdge> edges)
        {
            if (processes.Count == 0)
            {
                return new List<List<LayoutNode>>();
            }

            var ids = new HashSet<int>(processes.Select(p => p.Index));
            var processMap = processes.ToDictionary(p => p.Index);
            var processEdges = edges.Where(e => ids.Contains(e.From) && ids.Contains(e.To)).ToList();
            var layer = processes.ToDictionary(p => p.Index, p => 0);

            bool changed;
            int guard = Math.Max(1, processes.Count * processes.Count);
            do
            {
                changed = false;
                foreach (var edge in processEdges.OrderBy(e => e.From).ThenBy(e => e.To))
                {
                    int nextLayer = layer[edge.From] + 1;
                    if (nextLayer > layer[edge.To] && nextLayer < processes.Count)
                    {
                        layer[edge.To] = nextLayer;
                        changed = true;
                    }
                }
            }
            while (changed && guard-- > 0);

            var layers = processes
                .GroupBy(p => layer[p.Index])
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(p => NeighborScore(p.Index, processEdges, layer)).ThenBy(p => p.Index).ToList())
                .ToList();

            return layers.Count == 0 ? new List<List<LayoutNode>> { processes } : layers;
        }

        private static double NeighborScore(int nodeIndex, List<LayoutEdge> edges, Dictionary<int, int> layers)
        {
            var scores = edges
                .Where(e => e.From == nodeIndex || e.To == nodeIndex)
                .Select(e => e.From == nodeIndex ? e.To : e.From)
                .Where(layers.ContainsKey)
                .Select(id => (double)layers[id])
                .ToList();
            return scores.Count == 0 ? nodeIndex : scores.Average();
        }

        private static void PlaceProcesses(List<List<LayoutNode>> layers, Dictionary<int, LayoutPoint> result, double left, double right, double centerY, double rowGap)
        {
            for (int layerIndex = 0; layerIndex < layers.Count; layerIndex++)
            {
                var layer = layers[layerIndex];
                double x = layers.Count == 1 ? (left + right) / 2.0 : left + layerIndex * ((right - left) / (layers.Count - 1));
                double totalHeight = Math.Max(0, layer.Count - 1) * rowGap;
                double startY = centerY + totalHeight / 2.0;
                for (int i = 0; i < layer.Count; i++)
                {
                    result[layer[i].Index] = new LayoutPoint(x, startY - i * rowGap);
                }
            }
        }

        private static void PlaceEntities(List<LayoutNode> entities, List<LayoutEdge> edges, Dictionary<int, LayoutPoint> result, int leftCount, double width, double centerY, double rowGap)
        {
            var ordered = entities
                .Select(e => new
                {
                    Node = e,
                    Score = edges.Count(edge => edge.From == e.Index) - edges.Count(edge => edge.To == e.Index)
                })
                .OrderByDescending(x => x.Score)
                .ThenBy(x => x.Node.Index)
                .Select(x => x.Node)
                .ToList();

            PlaceColumn(ordered.Take(leftCount).ToList(), Margin + NodeWidth / 2.0, centerY, rowGap, result);
            PlaceColumn(ordered.Skip(leftCount).ToList(), width - Margin - NodeWidth / 2.0, centerY, rowGap, result);
        }

        private static void PlaceStores(List<LayoutNode> stores, List<LayoutEdge> edges, Dictionary<int, LayoutPoint> result, double left, double right, double y, double gap)
        {
            if (stores.Count == 0)
            {
                return;
            }

            var ordered = stores
                .OrderBy(s => NeighborAverageX(s.Index, edges, result))
                .ThenBy(s => s.Index)
                .ToList();

            double spacing = stores.Count == 1 ? 0 : Math.Max(gap, (right - left) / (stores.Count - 1));
            double totalWidth = Math.Max(0, stores.Count - 1) * spacing;
            double startX = stores.Count == 1 ? (left + right) / 2.0 : (left + right - totalWidth) / 2.0;
            for (int i = 0; i < ordered.Count; i++)
            {
                result[ordered[i].Index] = new LayoutPoint(startX + i * spacing, y);
            }
        }

        private static void PlaceColumn(List<LayoutNode> nodes, double x, double centerY, double rowGap, Dictionary<int, LayoutPoint> result)
        {
            if (nodes.Count == 0)
            {
                return;
            }

            double totalHeight = Math.Max(0, nodes.Count - 1) * rowGap;
            double startY = centerY + totalHeight / 2.0;
            for (int i = 0; i < nodes.Count; i++)
            {
                result[nodes[i].Index] = new LayoutPoint(x, startY - i * rowGap);
            }
        }

        private static double NeighborAverageX(int index, List<LayoutEdge> edges, Dictionary<int, LayoutPoint> result)
        {
            var xs = edges
                .Where(e => e.From == index || e.To == index)
                .Select(e => e.From == index ? e.To : e.From)
                .Where(result.ContainsKey)
                .Select(id => result[id].X)
                .ToList();
            return xs.Count == 0 ? double.MaxValue : xs.Average();
        }

        private static RoutedEdge RouteEdge(LayoutEdge edge, int slotIndex, Dictionary<int, LayoutPoint> nodes, double pageHeight)
        {
            var from = nodes[edge.From];
            var to = nodes[edge.To];
            double slotY = ChooseSlotY(from, to, slotIndex, pageHeight);
            double startX = from.X + Math.Sign(to.X - from.X == 0 ? 1 : to.X - from.X) * NodeWidth / 2.0;
            double endX = to.X - Math.Sign(to.X - from.X == 0 ? 1 : to.X - from.X) * NodeWidth / 2.0;
            double startY = from.Y;
            double endY = to.Y;

            if (Math.Abs(from.X - to.X) < 0.05)
            {
                double sideX = from.X + NodeWidth + 0.45 + slotIndex * 0.08;
                return BuildEdge(edge, new[]
                {
                    new LayoutPoint(from.X + NodeWidth / 2.0, startY),
                    new LayoutPoint(sideX, startY),
                    new LayoutPoint(sideX, endY),
                    new LayoutPoint(to.X + NodeWidth / 2.0, endY)
                });
            }

            return BuildEdge(edge, new[]
            {
                new LayoutPoint(startX, startY),
                new LayoutPoint(startX, slotY),
                new LayoutPoint(endX, slotY),
                new LayoutPoint(endX, endY)
            });
        }

        private static double ChooseSlotY(LayoutPoint from, LayoutPoint to, int slotIndex, double pageHeight)
        {
            bool useTop = !IsStoreLikeY(from.Y) && !IsStoreLikeY(to.Y);
            double baseY = useTop
                ? Math.Min(pageHeight - Margin - 0.35, Math.Max(from.Y, to.Y) + TopChannelPadding)
                : Math.Max(Margin + 1.35, Math.Min(from.Y, to.Y) - BottomChannelPadding);
            double offset = slotIndex * ChannelGap;
            return useTop ? Math.Min(pageHeight - Margin * 0.5, baseY + offset) : Math.Max(Margin * 0.5, baseY - offset);
        }

        private static bool IsStoreLikeY(double y)
        {
            return y < StoreLaneHeight + Margin + 0.3;
        }

        private static RoutedEdge BuildEdge(LayoutEdge edge, IEnumerable<LayoutPoint> points)
        {
            var distinct = new List<LayoutPoint>();
            foreach (var point in points)
            {
                if (distinct.Count == 0 || Math.Abs(distinct.Last().X - point.X) > 0.001 || Math.Abs(distinct.Last().Y - point.Y) > 0.001)
                {
                    distinct.Add(point);
                }
            }

            var routed = new RoutedEdge
            {
                From = edge.From,
                To = edge.To,
                Text = edge.Text ?? "",
                Points = distinct
            };
            routed.LabelPoint = GetLabelPoint(distinct);
            return routed;
        }

        private static LayoutPoint GetLabelPoint(List<LayoutPoint> points)
        {
            if (points.Count == 0)
            {
                return new LayoutPoint(0, 0);
            }

            if (points.Count >= 3)
            {
                var a = points[1];
                var b = points[2];
                return new LayoutPoint((a.X + b.X) / 2.0, (a.Y + b.Y) / 2.0 + 0.12);
            }

            var first = points.First();
            var last = points.Last();
            return new LayoutPoint((first.X + last.X) / 2.0, (first.Y + last.Y) / 2.0 + 0.12);
        }

        private static List<SubDiagramPlan> BuildSubDiagramPlans(List<LayoutNode> processes, List<LayoutEdge> edges)
        {
            return processes
                .Where(p => edges.Count(e => e.From == p.Index || e.To == p.Index) > MaxEdgesPerOverviewProcess)
                .Select(p => new SubDiagramPlan
                {
                    FocusNodeIndex = p.Index,
                    PageName = "2层数据流图-" + SanitizePageName(p.Text)
                })
                .ToList();
        }

        private static string SanitizePageName(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return "局部图";
            }

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                text = text.Replace(c, '-');
            }

            return text.Length > 24 ? text.Substring(0, 24) : text;
        }
    }
}
