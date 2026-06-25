using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioDataMapper
{
    public class DfdSmartSemanticTopologyLayout
    {
        public class LayoutNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
            public int Layer { get; set; }
            public double SortWeight { get; set; }
        }

        public class LayoutEdge
        {
            public int From { get; set; }
            public int To { get; set; }
        }

        public class LayoutResult
        {
            public double PageWidth { get; set; }
            public double PageHeight { get; set; }
            public Dictionary<int, Tuple<double, double>> Coordinates { get; set; } = new Dictionary<int, Tuple<double, double>>();
            public List<string> Warnings { get; set; } = new List<string>();
        }

        private const double NodeWidth = 1.55;
        private const double NodeHeight = 1.0;
        private const double Margin = 1.4;
        private const double EntityLaneWidth = 2.4;
        private const double StoreLaneHeight = 2.2;
        private const int HeavyEdgeThreshold = 9;

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
            var ids = new HashSet<int>(nodes.Select(n => n.Index));
            edges = edges.Where(e => ids.Contains(e.From) && ids.Contains(e.To) && e.From != e.To).ToList();

            var entities = nodes.Where(IsEntity).OrderBy(n => n.Index).ToList();
            var stores = nodes.Where(IsStore).OrderBy(n => n.Index).ToList();
            var processes = nodes.Where(n => !IsEntity(n) && !IsStore(n)).OrderBy(n => n.Index).ToList();
            if (processes.Count == 0)
            {
                processes = nodes.Where(n => !IsEntity(n)).OrderBy(n => n.Index).ToList();
            }

            AssignProcessLayers(processes, edges);
            var processLayers = BuildOrderedProcessLayers(processes, edges);

            int layerCount = Math.Max(1, processLayers.Count);
            int maxProcessRows = processLayers.Count == 0 ? 1 : Math.Max(1, processLayers.Max(l => l.Count));
            int leftEntityCount = CountLeftEntities(entities, edges);
            int rightEntityCount = entities.Count - leftEntityCount;
            int maxEntityRows = Math.Max(leftEntityCount, rightEntityCount);

            double layerGap = Math.Max(2.35, horGap + 1.35);
            double rowGap = Math.Max(1.65, verGap + 0.75);
            double storeGap = Math.Max(2.25, horGap + 1.0);
            double processWidth = Math.Max(1, layerCount - 1) * layerGap + NodeWidth;
            double storeWidth = stores.Count <= 1 ? NodeWidth : (stores.Count - 1) * storeGap + NodeWidth;
            double processHeight = Math.Max(1, maxProcessRows - 1) * rowGap + NodeHeight;
            double entityHeight = maxEntityRows <= 1 ? NodeHeight : (maxEntityRows - 1) * rowGap + NodeHeight;
            double routeReserve = Math.Max(2.1, Math.Min(4.5, edges.Count * 0.14 + 1.8));

            result.PageWidth = Math.Max(pageWidth, Margin * 2 + EntityLaneWidth * 2 + Math.Max(processWidth, storeWidth));
            result.PageHeight = Math.Max(pageHeight, Margin * 2 + StoreLaneHeight + routeReserve + Math.Max(processHeight, entityHeight));

            double processLeft = Margin + EntityLaneWidth;
            double processRight = result.PageWidth - Margin - EntityLaneWidth;
            double processCenterY = StoreLaneHeight + Margin + routeReserve + processHeight / 2.0;
            processCenterY = Math.Min(result.PageHeight - Margin - processHeight / 2.0, processCenterY);

            PlaceProcesses(processLayers, result.Coordinates, processLeft, processRight, processCenterY, rowGap);
            PlaceEntities(entities, edges, result.Coordinates, leftEntityCount, result.PageWidth, processCenterY, rowGap);
            PlaceStores(stores, edges, result.Coordinates, processLeft, processRight, Margin + 0.85, storeGap);

            foreach (var node in nodes)
            {
                if (!result.Coordinates.ContainsKey(node.Index))
                {
                    result.Coordinates[node.Index] = Tuple.Create((processLeft + processRight) / 2.0, processCenterY);
                }
            }

            AddComplexityWarnings(processes, edges, result.Warnings);
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

        private static void AssignProcessLayers(List<LayoutNode> processes, List<LayoutEdge> edges)
        {
            if (processes.Count == 0)
            {
                return;
            }

            var processIds = new HashSet<int>(processes.Select(p => p.Index));
            var processMap = processes.ToDictionary(p => p.Index);
            var processEdges = edges.Where(e => processIds.Contains(e.From) && processIds.Contains(e.To)).ToList();

            foreach (var process in processes)
            {
                process.Layer = 0;
            }

            bool changed;
            int guard = Math.Max(1, processes.Count * processes.Count);
            do
            {
                changed = false;
                foreach (var edge in processEdges.OrderBy(e => e.From).ThenBy(e => e.To))
                {
                    int candidate = processMap[edge.From].Layer + 1;
                    if (candidate > processMap[edge.To].Layer && candidate < processes.Count)
                    {
                        processMap[edge.To].Layer = candidate;
                        changed = true;
                    }
                }
            }
            while (changed && guard-- > 0);
        }

        private static List<List<LayoutNode>> BuildOrderedProcessLayers(List<LayoutNode> processes, List<LayoutEdge> edges)
        {
            if (processes.Count == 0)
            {
                return new List<List<LayoutNode>>();
            }

            var layers = processes
                .GroupBy(p => p.Layer)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(p => p.Index).ToList())
                .ToList();

            if (layers.Count == 1 && layers[0].Count > 3)
            {
                return SplitFlatProcessLayer(layers[0]);
            }

            for (int sweep = 0; sweep < 4; sweep++)
            {
                for (int i = 0; i < layers.Count; i++)
                {
                    var neighborX = BuildNeighborPositionMap(layers, i, edges);
                    foreach (var node in layers[i])
                    {
                        node.SortWeight = neighborX.ContainsKey(node.Index) ? neighborX[node.Index] : node.Index;
                    }
                    layers[i] = layers[i].OrderBy(n => n.SortWeight).ThenBy(n => n.Index).ToList();
                }
            }

            return layers;
        }

        private static List<List<LayoutNode>> SplitFlatProcessLayer(List<LayoutNode> nodes)
        {
            int columnCount = Math.Min(3, Math.Max(2, (int)Math.Ceiling(Math.Sqrt(nodes.Count))));
            int columnSize = (int)Math.Ceiling(nodes.Count / (double)columnCount);
            var result = new List<List<LayoutNode>>();

            for (int column = 0; column < columnCount; column++)
            {
                var chunk = nodes
                    .Skip(column * columnSize)
                    .Take(columnSize)
                    .OrderBy(n => n.Index)
                    .ToList();
                if (chunk.Count > 0)
                {
                    result.Add(chunk);
                }
            }

            return result;
        }

        private static Dictionary<int, double> BuildNeighborPositionMap(List<List<LayoutNode>> layers, int layerIndex, List<LayoutEdge> edges)
        {
            var neighborPositions = new Dictionary<int, int>();
            if (layerIndex > 0)
            {
                foreach (var item in layers[layerIndex - 1].Select((n, i) => new { n.Index, i }))
                {
                    neighborPositions[item.Index] = item.i;
                }
            }
            if (layerIndex < layers.Count - 1)
            {
                foreach (var item in layers[layerIndex + 1].Select((n, i) => new { n.Index, i }))
                {
                    neighborPositions[item.Index] = item.i;
                }
            }

            return layers[layerIndex]
                .Select(n => new
                {
                    n.Index,
                    Positions = edges
                        .Where(e => e.From == n.Index || e.To == n.Index)
                        .Select(e => e.From == n.Index ? e.To : e.From)
                        .Where(neighborPositions.ContainsKey)
                        .Select(id => neighborPositions[id])
                        .ToList()
                })
                .Where(x => x.Positions.Count > 0)
                .ToDictionary(x => x.Index, x => x.Positions.Average());
        }

        private static int CountLeftEntities(List<LayoutNode> entities, List<LayoutEdge> edges)
        {
            int inputHeavyCount = entities.Count(e => edges.Count(edge => edge.From == e.Index) >= edges.Count(edge => edge.To == e.Index));
            if (entities.Count == 0)
            {
                return 0;
            }
            return Math.Min(entities.Count, Math.Max(1, inputHeavyCount));
        }

        private static void PlaceProcesses(List<List<LayoutNode>> layers, Dictionary<int, Tuple<double, double>> result, double left, double right, double centerY, double rowGap)
        {
            for (int layerIndex = 0; layerIndex < layers.Count; layerIndex++)
            {
                var layer = layers[layerIndex];
                double x = layers.Count == 1 ? (left + right) / 2.0 : left + layerIndex * ((right - left) / (layers.Count - 1));
                double totalHeight = Math.Max(0, layer.Count - 1) * rowGap;
                double startY = centerY + totalHeight / 2.0;
                for (int i = 0; i < layer.Count; i++)
                {
                    result[layer[i].Index] = Tuple.Create(x, startY - i * rowGap);
                }
            }
        }

        private static void PlaceEntities(List<LayoutNode> entities, List<LayoutEdge> edges, Dictionary<int, Tuple<double, double>> result, int leftCount, double pageWidth, double centerY, double rowGap)
        {
            var ordered = entities
                .OrderByDescending(e => edges.Count(edge => edge.From == e.Index) - edges.Count(edge => edge.To == e.Index))
                .ThenBy(e => e.Index)
                .ToList();

            PlaceColumn(ordered.Take(leftCount).ToList(), Margin + NodeWidth / 2.0, centerY, rowGap, result);
            PlaceColumn(ordered.Skip(leftCount).ToList(), pageWidth - Margin - NodeWidth / 2.0, centerY, rowGap, result);
        }

        private static void PlaceStores(List<LayoutNode> stores, List<LayoutEdge> edges, Dictionary<int, Tuple<double, double>> result, double left, double right, double y, double gap)
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
                result[ordered[i].Index] = Tuple.Create(startX + i * spacing, y);
            }
        }

        private static void PlaceColumn(List<LayoutNode> nodes, double x, double centerY, double rowGap, Dictionary<int, Tuple<double, double>> result)
        {
            if (nodes.Count == 0)
            {
                return;
            }

            double totalHeight = Math.Max(0, nodes.Count - 1) * rowGap;
            double startY = centerY + totalHeight / 2.0;
            for (int i = 0; i < nodes.Count; i++)
            {
                result[nodes[i].Index] = Tuple.Create(x, startY - i * rowGap);
            }
        }

        private static double NeighborAverageX(int index, List<LayoutEdge> edges, Dictionary<int, Tuple<double, double>> result)
        {
            var xs = edges
                .Where(e => e.From == index || e.To == index)
                .Select(e => e.From == index ? e.To : e.From)
                .Where(result.ContainsKey)
                .Select(id => result[id].Item1)
                .ToList();
            return xs.Count == 0 ? double.MaxValue : xs.Average();
        }

        private static void AddComplexityWarnings(List<LayoutNode> processes, List<LayoutEdge> edges, List<string> warnings)
        {
            foreach (var process in processes)
            {
                int degree = edges.Count(e => e.From == process.Index || e.To == process.Index);
                if (degree > HeavyEdgeThreshold)
                {
                    warnings.Add($"{process.Text} 连接数较多，建议必要时拆成局部 2层数据流图。");
                }
            }
        }
    }
}
