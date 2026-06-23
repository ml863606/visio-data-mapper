using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioDataMapper
{
    /// <summary>
    /// DFD hybrid layout: semantic zones for DFD roles plus topological layering for process nodes.
    /// </summary>
    public class DfdHybridLayout
    {
        public class LayoutNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; }
            public int Layer { get; set; } = -1;
            public double Barycenter { get; set; }
        }

        public class LayoutEdge
        {
            public int From { get; set; }
            public int To { get; set; }
        }

        public Dictionary<int, Tuple<double, double>> CalculateLayout(
            List<LayoutNode> nodes,
            List<LayoutEdge> edges,
            double pageWidth,
            double pageHeight,
            double horGap,
            double verGap)
        {
            var result = new Dictionary<int, Tuple<double, double>>();
            if (nodes == null || nodes.Count == 0)
            {
                return result;
            }

            edges = edges ?? new List<LayoutEdge>();

            double nodeW = 1.4;
            double nodeH = 1.0;
            double margin = 1.3;
            double entityLaneWidth = 2.1;
            double storeLaneHeight = 1.9;

            var entities = nodes.Where(IsEntity).ToList();
            var stores = nodes.Where(IsStore).ToList();
            var processes = nodes.Where(n => !IsEntity(n) && !IsStore(n)).ToList();
            if (processes.Count == 0)
            {
                processes = nodes.Where(n => !IsEntity(n)).ToList();
            }

            AssignProcessLayers(processes, edges);
            var processLayers = GroupAndOrderProcessLayers(processes, edges);

            int layerCount = Math.Max(1, processLayers.Count);
            int maxLayerCount = processLayers.Count == 0 ? 1 : Math.Max(1, processLayers.Max(l => l.Count));
            int leftEntityCount = (entities.Count + 1) / 2;
            int rightEntityCount = entities.Count / 2;
            int maxEntityCount = Math.Max(leftEntityCount, rightEntityCount);

            double layerGap = Math.Max(horGap + 1.0, 2.6);
            double rowGap = Math.Max(verGap + 0.65, 1.6);
            double processAreaWidth = Math.Max(1, layerCount - 1) * layerGap + nodeW;
            double processAreaHeight = Math.Max(1, maxLayerCount - 1) * rowGap + nodeH;
            double entityHeight = maxEntityCount > 0 ? (maxEntityCount - 1) * rowGap + nodeH : 0;
            double storeWidth = stores.Count > 0 ? (stores.Count - 1) * Math.Max(horGap + 1.0, 2.2) + nodeW : 0;

            double requiredWidth = Math.Max(pageWidth, margin * 2 + entityLaneWidth * 2 + Math.Max(processAreaWidth, storeWidth));
            double requiredHeight = Math.Max(pageHeight, margin * 2 + storeLaneHeight + Math.Max(processAreaHeight, entityHeight) + 1.2);

            double processLeft = margin + entityLaneWidth;
            double processRight = requiredWidth - margin - entityLaneWidth;
            double processTop = requiredHeight - margin - 0.9;
            double processCenterY = (storeLaneHeight + margin + processTop) / 2.0;

            for (int layerIndex = 0; layerIndex < processLayers.Count; layerIndex++)
            {
                var layer = processLayers[layerIndex];
                double x = layerCount == 1
                    ? (processLeft + processRight) / 2.0
                    : processLeft + layerIndex * ((processRight - processLeft) / (layerCount - 1));

                double layerHeight = (layer.Count - 1) * rowGap;
                double startY = processCenterY + layerHeight / 2.0;
                for (int i = 0; i < layer.Count; i++)
                {
                    result[layer[i].Index] = Tuple.Create(x, startY - i * rowGap);
                }
            }

            PlaceEntities(entities, edges, result, leftEntityCount, requiredWidth, processCenterY, rowGap, margin, nodeW);
            PlaceStores(stores, edges, result, processLeft, processRight, storeLaneHeight, horGap, nodeW);

            foreach (var node in nodes)
            {
                if (!result.ContainsKey(node.Index))
                {
                    result[node.Index] = Tuple.Create(requiredWidth / 2.0, processCenterY);
                }
            }

            result[-1] = Tuple.Create(requiredWidth, requiredHeight);
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
            var processEdges = edges
                .Where(e => processIds.Contains(e.From) && processIds.Contains(e.To) && e.From != e.To)
                .ToList();

            foreach (var process in processes)
            {
                process.Layer = -1;
            }

            var inDegree = processes.ToDictionary(p => p.Index, p => 0);
            var outgoing = processes.ToDictionary(p => p.Index, p => new List<int>());
            foreach (var edge in processEdges)
            {
                inDegree[edge.To]++;
                outgoing[edge.From].Add(edge.To);
            }

            var queue = new Queue<int>(processes.Where(p => inDegree[p.Index] == 0).Select(p => p.Index));
            foreach (int id in queue)
            {
                processMap[id].Layer = 0;
            }

            if (queue.Count == 0)
            {
                queue.Enqueue(processes[0].Index);
                processes[0].Layer = 0;
            }

            int guard = Math.Max(1, processes.Count * processes.Count);
            while (queue.Count > 0 && guard-- > 0)
            {
                int current = queue.Dequeue();
                int nextLayer = processMap[current].Layer + 1;
                foreach (int target in outgoing[current])
                {
                    if (nextLayer > processMap[target].Layer)
                    {
                        processMap[target].Layer = nextLayer;
                        queue.Enqueue(target);
                    }
                }
            }

            foreach (var process in processes)
            {
                if (process.Layer < 0)
                {
                    int incomingLayer = processEdges
                        .Where(e => e.To == process.Index && processMap.ContainsKey(e.From) && processMap[e.From].Layer >= 0)
                        .Select(e => processMap[e.From].Layer + 1)
                        .DefaultIfEmpty(0)
                        .Max();
                    process.Layer = incomingLayer;
                }
            }
        }

        private static List<List<LayoutNode>> GroupAndOrderProcessLayers(List<LayoutNode> processes, List<LayoutEdge> edges)
        {
            if (processes.Count == 0)
            {
                return new List<List<LayoutNode>>();
            }

            var layers = processes
                .GroupBy(p => p.Layer)
                .OrderBy(g => g.Key)
                .Select(g => g.OrderBy(n => n.Index).ToList())
                .ToList();

            for (int sweep = 0; sweep < 4; sweep++)
            {
                for (int i = 1; i < layers.Count; i++)
                {
                    OrderLayerByNeighbors(layers, edges, i, i - 1, true);
                }

                for (int i = layers.Count - 2; i >= 0; i--)
                {
                    OrderLayerByNeighbors(layers, edges, i, i + 1, false);
                }
            }

            return layers;
        }

        private static void OrderLayerByNeighbors(List<List<LayoutNode>> layers, List<LayoutEdge> edges, int layerIndex, int neighborLayerIndex, bool useParents)
        {
            var neighborPositions = layers[neighborLayerIndex]
                .Select((n, index) => new { n.Index, Position = index })
                .ToDictionary(x => x.Index, x => x.Position);

            foreach (var node in layers[layerIndex])
            {
                var positions = useParents
                    ? edges.Where(e => e.To == node.Index && neighborPositions.ContainsKey(e.From)).Select(e => neighborPositions[e.From])
                    : edges.Where(e => e.From == node.Index && neighborPositions.ContainsKey(e.To)).Select(e => neighborPositions[e.To]);

                var list = positions.ToList();
                node.Barycenter = list.Count == 0 ? node.Barycenter : list.Average();
            }

            layers[layerIndex] = layers[layerIndex]
                .OrderBy(n => n.Barycenter)
                .ThenBy(n => n.Index)
                .ToList();
        }

        private static void PlaceEntities(
            List<LayoutNode> entities,
            List<LayoutEdge> edges,
            Dictionary<int, Tuple<double, double>> result,
            int leftCount,
            double width,
            double processCenterY,
            double rowGap,
            double margin,
            double nodeW)
        {
            var orderedEntities = entities
                .Select(e => new
                {
                    Node = e,
                    Score = edges.Count(edge => edge.From == e.Index) - edges.Count(edge => edge.To == e.Index)
                })
                .OrderByDescending(x => x.Score)
                .ThenBy(x => x.Node.Index)
                .Select(x => x.Node)
                .ToList();

            var left = orderedEntities.Take(leftCount).ToList();
            var right = orderedEntities.Skip(leftCount).ToList();
            PlaceVerticalColumn(left, margin + nodeW / 2.0, processCenterY, rowGap, result);
            PlaceVerticalColumn(right, width - margin - nodeW / 2.0, processCenterY, rowGap, result);
        }

        private static void PlaceStores(
            List<LayoutNode> stores,
            List<LayoutEdge> edges,
            Dictionary<int, Tuple<double, double>> result,
            double left,
            double right,
            double y,
            double horGap,
            double nodeW)
        {
            if (stores.Count == 0)
            {
                return;
            }

            var orderedStores = stores
                .OrderBy(s => GetNeighborAverageX(s.Index, edges, result))
                .ThenBy(s => s.Index)
                .ToList();

            double spacing = stores.Count == 1 ? 0 : Math.Max(horGap + 1.0, (right - left) / (stores.Count - 1));
            double totalWidth = (stores.Count - 1) * spacing;
            double startX = stores.Count == 1 ? (left + right) / 2.0 : (left + right - totalWidth) / 2.0;
            for (int i = 0; i < orderedStores.Count; i++)
            {
                result[orderedStores[i].Index] = Tuple.Create(startX + i * spacing, y);
            }
        }

        private static double GetNeighborAverageX(int storeIndex, List<LayoutEdge> edges, Dictionary<int, Tuple<double, double>> result)
        {
            var xs = edges
                .Where(e => e.From == storeIndex || e.To == storeIndex)
                .Select(e => e.From == storeIndex ? e.To : e.From)
                .Where(result.ContainsKey)
                .Select(id => result[id].Item1)
                .ToList();

            return xs.Count == 0 ? double.MaxValue : xs.Average();
        }

        private static void PlaceVerticalColumn(List<LayoutNode> nodes, double x, double centerY, double rowGap, Dictionary<int, Tuple<double, double>> result)
        {
            if (nodes.Count == 0)
            {
                return;
            }

            double totalHeight = (nodes.Count - 1) * rowGap;
            double startY = centerY + totalHeight / 2.0;
            for (int i = 0; i < nodes.Count; i++)
            {
                result[nodes[i].Index] = Tuple.Create(x, startY - i * rowGap);
            }
        }
    }
}
