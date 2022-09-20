using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using ExcelIP = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

// TODO:   按照以下步骤启用功能区(XML)项: 

// 1. 将以下代码块复制到 ThisAddin、ThisWorkbook 或 ThisDocument 类中。

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new RosRibbon();
//  }

// 2. 在此类的“功能区回调”区域中创建回调方法，以处理用户
//    操作(如单击某个按钮)。注意: 如果已经从功能区设计器中导出此功能区，
//    则将事件处理程序中的代码移动到回调方法并修改该代码以用于
//    功能区扩展性(RibbonX)编程模型。

// 3. 向功能区 XML 文件中的控制标记分配特性，以标识代码中的相应回调方法。  

// 有关详细信息，请参见 Visual Studio Tools for Office 帮助中的功能区 XML 文档。


namespace KK.ROSExcelAddin
{
    [ComVisible(true)]
    public class RosRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RosRibbon()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("KK.ROSExcelAddin.RosRibbon.xml");
        }

        #endregion

        #region 功能区回调
        //在此处创建回叫方法。有关添加回叫方法的详细信息，请访问 https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }


        public void OnClick_DHCPAbout(Office.IRibbonControl ctrl)
        {
            if (ctrl == null) return;

            switch (ctrl.Id)
            {
                case "BTN_ROS_TODHCPSCRIPT":
                    GenerateDHCPAddScript();
                    break;
            }
        }

        #endregion


        #region DHCP模块

        public void GenerateDHCPAddScript()
        {
            try
            {
                StringBuilder _sb = new StringBuilder();
                ExcelIP.Worksheet _xlSheet = (ExcelIP.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                ExcelIP.Application _xlApp = Globals.ThisAddIn.Application;

                if (_xlSheet == null)
                {
                    ShowMessage("未知的工作表对象", MessageBoxIcon.Error);
                    return;
                }

                // 获取选择的行
                ///ShowMessage(_xlApp.Selection.Address);
                //ShowMessage(_xlApp.Selection.Count);
                //return;

                ExcelIP.Range _selRng = null;

                // 验证选择的区域是不是单元格区域
                try
                {
                    _selRng =(ExcelIP.Range)_xlApp.Selection;
                }
                catch (Exception ex)
                {
                    String errMsg = "【选择的对象不是工作表单元格区域。】";
                    if (Common.ShowDebugInfo)
                    {
                        errMsg += ex.Message + ex.StackTrace;
                    }
                    ShowMessage(errMsg);
                    return;
                }

                // 通过Range获取所有行号
                List<Int32> listRowNumber = new List<int>();
                if (_selRng.Count > 0)
                {
                    foreach (ExcelIP.Range rng in _selRng)
                    {
                        if (!listRowNumber.Contains(rng.Row))
                        {
                            listRowNumber.Add(rng.Row);
                        }
                    }
                }

                //遍历行
                foreach (Int32 rowNo in listRowNumber)
                {
                    //ShowMessage("第｛" + rowNo + "｝行");
                    ZS.RouterOS.DHCP.Lease model = new ZS.RouterOS.DHCP.Lease();
                    model.MacAddress = new ZS.RouterOS._MACAddress(_xlSheet.Range["A" + rowNo].Value);
                    model.Address = new ZS.RouterOS._IPAddress(_xlSheet.Range["B" + rowNo].Value);
                    model.Server = _xlSheet.Range["C" + rowNo].Value;
                    model.Comment = _xlSheet.Range["E" + rowNo].Value;

                    _sb.Append(model.ToAddScriptText());
                }

               DHD.WinFormControls.TextPreview tv = new DHD.WinFormControls.TextPreview();
                tv.PreviewText = _sb.ToString();
                tv.Show();

                return;
            }
            catch (Exception ex)
            {
                String errMsg = "【选择的对象不是工作表单元格区域。】";
                if (Common.ShowDebugInfo)
                {
                    errMsg += ex.Message + ex.StackTrace;
                }
                ShowMessage(errMsg);
            }
        }


        #endregion


        private void ShowMessage(String msg, MessageBoxIcon ico = MessageBoxIcon.Information)
        {
            MessageBox.Show(msg, "", MessageBoxButtons.OK, ico);
        }

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
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
