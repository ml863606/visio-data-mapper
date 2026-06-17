using System;
using System.Collections.Generic;
using System.Linq;

namespace VisioDataMapper
{
    /// <summary>
    /// DFD 语义感知分区布局算法。
    /// 核心思路：不走通用图论路线，而是利用 DFD 节点的类型语义（实体/加工/数据存储）
    /// 来直接确定空间位置，产出类似教科书级别的标准 DFD 布局。
    /// 
    /// 布局结构（自上而下）：
    ///   ┌────────────────────────────────────────────┐
    ///   │  上区：左侧实体        右侧实体             │
    ///   ├────────────────────────────────────────────┤
    ///   │  中区：加工过程行（水平均匀排列）            │
    ///   ├────────────────────────────────────────────┤
    ///   │  下区：数据存储行（水平均匀排列）            │
    ///   └────────────────────────────────────────────┘
    /// </summary>
    public class DfdSemanticLayout
    {
        public class LayoutNode
        {
            public int Index { get; set; }
            public string Text { get; set; }
            public string Type { get; set; } // "实体" / "加工" / "数据存储"
        }

        public class LayoutEdge
        {
            public int From { get; set; }
            public int To { get; set; }
        }

        /// <summary>
        /// 基于 DFD 语义执行三区分层布局计算，返回每个节点的精确 Visio 英寸坐标。
        /// </summary>
        public Dictionary<int, Tuple<double, double>> CalculateLayout(
            List<LayoutNode> nodes,
            List<LayoutEdge> edges,
            double pageWidth,
            double pageHeight,
            double horGap,
            double verGap)
        {
            if (nodes == null || nodes.Count == 0)
                return new Dictionary<int, Tuple<double, double>>();

            // ─────── Phase 1: 按 DFD 语义分类 ───────
            var entities = nodes.Where(n => n.Type == "实体").ToList();
            var processes = nodes.Where(n => n.Type == "加工").ToList();
            var stores = nodes.Where(n => n.Type == "数据存储").ToList();

            // 未归类的节点视为加工
            var unclassified = nodes.Where(n => n.Type != "实体" && n.Type != "加工" && n.Type != "数据存储").ToList();
            processes.AddRange(unclassified);

            // ─────── Phase 2: 动态计算画布所需尺寸 ───────
            double nodeW = 1.6;   // 单个节点的参考宽度（英寸）
            double nodeH = 1.0;   // 单个节点的参考高度（英寸）
            double margin = 1.5;  // 页面边距

            // 计算加工行所需宽度
            double processRowWidth = processes.Count * nodeW + (Math.Max(0, processes.Count - 1)) * horGap;
            // 计算存储行所需宽度
            double storeRowWidth = stores.Count * nodeW + (Math.Max(0, stores.Count - 1)) * horGap;
            // 取两行中较宽的，加上两侧实体的空间
            double entityColumnWidth = entities.Count > 0 ? (nodeW + horGap) : 0;
            double requiredWidth = Math.Max(processRowWidth, storeRowWidth) + entityColumnWidth * 2 + margin * 2;

            // 计算实体列所需高度
            int leftCount = (entities.Count + 1) / 2;  // 左侧实体数量（奇数时多一个放左边）
            int rightCount = entities.Count / 2;        // 右侧实体数量
            int maxEntityCol = Math.Max(leftCount, rightCount);
            double entityColumnHeight = maxEntityCol * nodeH + (Math.Max(0, maxEntityCol - 1)) * verGap;

            // 三区结构的最低高度：实体区 + 通道 + 过程区 + 通道 + 存储区 + 边距
            double zoneGap = Math.Max(2.0, verGap * 1.5); // 每两个区之间的垂直走线通道
            double requiredHeight = margin + entityColumnHeight + zoneGap + nodeH + zoneGap + nodeH + margin;

            // 动态调整页面大小
            double effectiveWidth = Math.Max(pageWidth, requiredWidth);
            double effectiveHeight = Math.Max(pageHeight, requiredHeight);

            // ─────── Phase 3: 确定三个区域的 Y 坐标锚点 ───────
            // Visio 坐标系：Y=0 在底部，Y=pageHeight 在顶部
            //
            //   顶部 ──── 实体区 (entityRowY)      ← 页面上部 ~75%
            //   中间 ──── 过程行 (processRowY)      ← 页面正中 ~50%
            //   底部 ──── 存储行 (storeRowY)        ← 页面下部 ~20%

            // 实体区的 Y 位置：页面上部（约 75% 高度处）
            double entityRowY = effectiveHeight * 0.75;

            // 过程行的 Y 位置：页面正中间（约 50% 高度处）
            double processRowY = effectiveHeight * 0.48;

            // 数据存储行的 Y 位置：页面底部（约 20% 高度处）
            double storeRowY = effectiveHeight * 0.18;

            // ─────── Phase 4: 排布加工过程行（中间区域，水平均匀分布） ───────
            var result = new Dictionary<int, Tuple<double, double>>();

            // 加工区域的 X 范围：为两侧实体列留出空间
            double processAreaLeft = margin + entityColumnWidth;
            double processAreaRight = effectiveWidth - margin - entityColumnWidth;
            double processAreaWidth = processAreaRight - processAreaLeft;

            if (processes.Count == 1)
            {
                double cx = (processAreaLeft + processAreaRight) / 2.0;
                result[processes[0].Index] = Tuple.Create(cx, processRowY);
            }
            else if (processes.Count > 1)
            {
                double spacing = processAreaWidth / (processes.Count - 1);
                if (spacing < nodeW + horGap * 0.5)
                {
                    spacing = nodeW + horGap;
                    processAreaWidth = spacing * (processes.Count - 1);
                    processAreaLeft = (effectiveWidth - processAreaWidth) / 2.0;
                }

                for (int i = 0; i < processes.Count; i++)
                {
                    double cx = processAreaLeft + i * spacing;
                    result[processes[i].Index] = Tuple.Create(cx, processRowY);
                }
            }

            // ─────── Phase 5: 排布外部实体（上部左右两侧，垂直分布） ───────
            var leftEntities = entities.Take(leftCount).ToList();
            var rightEntities = entities.Skip(leftCount).ToList();

            // 左侧实体列
            if (leftEntities.Count > 0)
            {
                double leftX = margin + nodeW / 2;

                if (leftEntities.Count == 1)
                {
                    result[leftEntities[0].Index] = Tuple.Create(leftX, entityRowY);
                }
                else
                {
                    double totalH = (leftEntities.Count - 1) * (nodeH + verGap);
                    double startY = entityRowY + totalH / 2;
                    for (int i = 0; i < leftEntities.Count; i++)
                    {
                        double cy = startY - i * (nodeH + verGap);
                        result[leftEntities[i].Index] = Tuple.Create(leftX, cy);
                    }
                }
            }

            // 右侧实体列
            if (rightEntities.Count > 0)
            {
                double rightX = effectiveWidth - margin - nodeW / 2;

                if (rightEntities.Count == 1)
                {
                    result[rightEntities[0].Index] = Tuple.Create(rightX, entityRowY);
                }
                else
                {
                    double totalH = (rightEntities.Count - 1) * (nodeH + verGap);
                    double startY = entityRowY + totalH / 2;
                    for (int i = 0; i < rightEntities.Count; i++)
                    {
                        double cy = startY - i * (nodeH + verGap);
                        result[rightEntities[i].Index] = Tuple.Create(rightX, cy);
                    }
                }
            }

            // ─────── Phase 6: 排布数据存储行（底部区域，水平均匀分布） ───────
            if (stores.Count == 1)
            {
                double cx = effectiveWidth / 2.0;
                result[stores[0].Index] = Tuple.Create(cx, storeRowY);
            }
            else if (stores.Count > 1)
            {
                double storeAreaLeft = processAreaLeft;
                double storeAreaRight = processAreaRight;
                double storeSpacing = (storeAreaRight - storeAreaLeft) / (stores.Count - 1);

                if (storeSpacing < nodeW + horGap * 0.5)
                {
                    storeSpacing = nodeW + horGap;
                    double storeW = storeSpacing * (stores.Count - 1);
                    storeAreaLeft = (effectiveWidth - storeW) / 2.0;
                }

                for (int i = 0; i < stores.Count; i++)
                {
                    double cx = storeAreaLeft + i * storeSpacing;
                    result[stores[i].Index] = Tuple.Create(cx, storeRowY);
                }
            }

            // ─────── Phase 7: 记录所需的有效画布尺寸（通过特殊键传递） ───────
            result[-1] = Tuple.Create(effectiveWidth, effectiveHeight);

            return result;
        }
    }
}
