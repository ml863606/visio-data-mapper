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
            if (control.Id == "btnBasicFlowchart")
            {
                using (var form = new FormBasicFlowchart())
                {
                    form.ShowDialog();
                }
                return;
            }
            if (control.Id == "btnSysArch")
            {
                using (var form = new FormSystemArchitectureDiagram())
                {
                    form.ShowDialog();
                }
                return;
            }

            switch (control.Id)
            {
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
                case "btnEntityAttribute":
                    using (var form = new FormEntityAttributeDiagram())
                    {
                        form.ShowDialog();
                    }
                    return;
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
                case "btnDataFlowTop":
                    using (var form = new FormDataFlowDiagram(FormDataFlowDiagram.DataFlowDiagramLevel.Top))
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnDataFlowLevel1":
                    using (var form = new FormDataFlowDiagram(FormDataFlowDiagram.DataFlowDiagramLevel.Level1))
                    {
                        form.ShowDialog();
                    }
                    return;
                case "btnDataFlow":
                case "btnDataFlowLevel2":
                    using (var form = new FormDataFlowDiagram(FormDataFlowDiagram.DataFlowDiagramLevel.Level2))
                    {
                        form.ShowDialog();
                    }
                    return;

                default:
                    MessageBox.Show($"未知的操作: {control.Id}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
            }
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
