using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Msagl.Core.Geometry;
using Microsoft.Msagl.Core.Geometry.Curves;
using Microsoft.Msagl.Core.Layout;
using Microsoft.Msagl.Layout.Layered;
using Microsoft.Msagl.Miscellaneous;

namespace VisioDataMapper
{
    public class DfdMsaglLayout
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
        }

        /// <summary>
        /// 调用微软开源 MSAGL 自动构图排版引擎，计算图的节点在 Visio 中的精确英寸坐标
        /// </summary>
        /// <param name="nodes">节点列表</param>
        /// <param name="edges">连接边列表</param>
        /// <param name="pageWidth">画布宽度（英寸）</param>
        /// <param name="pageHeight">画布高度（英寸）</param>
        /// <param name="horGap">水平间距参数</param>
        /// <param name="verGap">垂直间距参数</param>
        /// <returns>节点 Index 对应的 X, Y 坐标字典</returns>
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

            // 1. 创建 MSAGL 几何图 (GeometryGraph)
            var graph = new GeometryGraph();

            // DFD 节点的默认尺寸设计为 130x90（像素级单位，对应大约 2.16x1.5 英寸，比实际画的 1.2x0.8 稍大，能够预留更足的包络安全空隙）
            double nodeW = 130.0;
            double nodeH = 90.0;

            var msaglNodes = new Dictionary<int, Node>();

            // 2. 批量注册节点到几何图
            foreach (var node in nodes)
            {
                var boundary = CurveFactory.CreateRectangle(nodeW, nodeH, new Point(0, 0));
                var msaglNode = new Node(boundary, node.Index);
                graph.Nodes.Add(msaglNode);
                msaglNodes[node.Index] = msaglNode;
            }

            // 3. 批量注册边到几何图
            foreach (var edge in edges)
            {
                if (msaglNodes.ContainsKey(edge.From) && msaglNodes.ContainsKey(edge.To))
                {
                    var msaglEdge = new Edge(msaglNodes[edge.From], msaglNodes[edge.To]);
                    graph.Edges.Add(msaglEdge);
                }
            }

            // 4. 配置 SugiyamaLayoutSettings (层次流向布局)
            var settings = new SugiyamaLayoutSettings();

            // 旋转 90 度以形成从左到右 (LR) 的网状层级流动
            settings.Transformation = PlaneTransformation.Rotation(Math.PI / 2);

            // 间距转换（将 Visio 的间距参数映射为 MSAGL 的像素分离度。大大调大安全间距）
            settings.NodeSeparation = Math.Max(120.0, horGap * 60.0);
            settings.LayerSeparation = Math.Max(140.0, verGap * 60.0);

            // 5. 执行 MSAGL 高级几何计算
            LayoutHelpers.CalculateLayout(graph, settings, null);

            // 6. 进行物理投影与坐标映射
            double minX = graph.Nodes.Min(n => n.Center.X);
            double maxX = graph.Nodes.Max(n => n.Center.X);
            double minY = graph.Nodes.Min(n => n.Center.Y);
            double maxY = graph.Nodes.Max(n => n.Center.Y);

            // 平移并按 GDI 到 Visio 英寸比例（固定使用 60 像素 = 1 英寸的比例，绝对不压缩）
            double scale = 1.0 / 60.0;

            var result = new Dictionary<int, Tuple<double, double>>();
            double marginX = 1.2;
            double marginY = 1.2;

            foreach (var node in nodes)
            {
                var msaglNode = msaglNodes[node.Index];
                double rawX = msaglNode.Center.X;
                double rawY = msaglNode.Center.Y;

                // 转换坐标
                double visioX = marginX + (rawX - minX) * scale;
                // 反转 Y 轴，因为 MSAGL 和 Visio 的纵向朝向相反，用 maxY - rawY 实现精准垂直反转
                double visioY = marginY + (maxY - rawY) * scale;

                result[node.Index] = new Tuple<double, double>(visioX, visioY);
            }

            return result;
        }
    }
}
