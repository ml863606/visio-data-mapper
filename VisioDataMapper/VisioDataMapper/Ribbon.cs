using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace VisioDataMapper
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("VisioDataMapper.Ribbon.xml");
        }

        #endregion

        #region 功能区回调

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAction(Office.IRibbonControl control)
        {
            string featureName = "未知功能";
            if (control.Id == "btnBasicFlowchart")
            {
                using (var form = new FormBasicFlowchart())
                {
                    form.ShowDialog();
                }
                return;
            }

            switch (control.Id)
            {
                // 会员信息
                case "btnMemberManage": featureName = "会员管理"; break;
                case "btnTutorial1": featureName = "智能画图教程 - 基础教程"; break;
                case "btnTutorial2": featureName = "智能画图教程 - 进阶教程"; break;
                case "btnTemplate1": featureName = "自定义套模板 - 模板管理"; break;
                case "btnTemplate2": featureName = "自定义套模板 - 新建模板"; break;
                case "btnBindDocData": featureName = "文档绑定数据"; break;
                case "btnCustomShapes": featureName = "自定义形状"; break;
                case "btnExportChartToTable": featureName = "绘图导出表格"; break;

                // 流程图
                case "btnBasicFlowchart": featureName = "基本流程图"; break;
                case "btnSwimlane": featureName = "泳道图"; break;
                case "btnSwimlaneAdd": featureName = "泳道图编辑 - 添加泳道"; break;
                case "btnSwimlaneDel": featureName = "泳道图编辑 - 删除泳道"; break;
                case "btnTechRoadmap": featureName = "技术路线图"; break;
                case "btnChemFlowchart": featureName = "化工流程图"; break;
                case "btnNSChart": featureName = "NS盒图"; break;

                // 架构图
                case "btnOrgChart":
                    using (var form = new FormOrgChart())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnModuleDiagram":
                    using (var form = new FormModuleDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnLayeredArch": featureName = "层级架构图"; break;
                case "btnEquityArch": featureName = "股权架构图"; break;
                case "btnSysArch": featureName = "系统架构图"; break;
                case "btnMatrixArch": featureName = "矩阵架构图"; break;
                case "btnFamilyTree": featureName = "家谱图"; break;

                // 软件与数据库
                case "btnFuncStructure": featureName = "功能结构图"; break;
                case "btnClassDiagram": featureName = "类图"; break;
                case "btnEntityAttribute":
                    using (var form = new FormEntityAttributeDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnERLogical": featureName = "ER图 - 逻辑模型"; break;
                case "btnERPhysical": featureName = "ER图 - 物理模型"; break;
                case "btnUseCaseDiagram":
                    using (var form = new FormUseCaseDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnSeqNormal":
                case "btnSeqLoop":
                    using (var form = new FormSequenceDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnActivityDiagram": featureName = "活动图"; break;
                case "btnStateDiagram": featureName = "状态图"; break;
                case "btnDataFlow":
                    using (var form = new FormDataFlowDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnUMLDeployment": featureName = "更多UML绘图 - 部署图"; break;
                case "btnUMLComponent": featureName = "更多UML绘图 - 组件图"; break;

                // 项目管理
                case "btnMindmap": featureName = "思维导图"; break;
                case "btnDecisionTree": featureName = "决策树图"; break;
                case "btnFishbone": featureName = "鱼骨图"; break;
                case "btnGantt": featureName = "甘特图"; break;
                case "btnTimeline": featureName = "时间轴"; break;
                case "btnDrawValueStream": featureName = "更多绘图 - 价值流图"; break;
                case "btnDrawBPMN": featureName = "更多绘图 - BPMN图"; break;

                // 其他绘图
                case "btnNetworkTopology": featureName = "网络拓扑图"; break;

                default:
                    MessageBox.Show($"未知的操作: {control.Id}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
            }
            MessageBox.Show($"正在启动/使用: {featureName}...", "Visio画图小助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] names = asm.GetManifestResourceNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (string.Compare(resourceName, names[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(names[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
