using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ingr.SP3D.Common.Client;
using Ingr.SP3D.Common.Client.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace MyNameSpace
{
    public class SetAssemblyBoardSide : BaseModalCommand
    {
        public override void OnStart(int commandID, object argument)
        {
            List<string[]> listProperty = new List<string[]>();
            Stopwatch sw = new Stopwatch();
            sw.Start();
            base.OnStart(commandID, argument);
            SelectSet ssBlock = ClientServiceProvider.SelectSet;
            try
            {
                if (ssBlock.SelectedObjects.Count == 1)
                {
                    if ((ssBlock.SelectedObjects[0]) is IAssembly)
                    {
                        IAssembly assBlock = (ssBlock.SelectedObjects[0]) as IAssembly;

                        if (assBlock.AssemblyChildren.Count > 0 && ssBlock.SelectedObjects[0].MyGetProperty("Type") == "B")
                        {
                            base.WriteStatusBarMsg("In Porgress for  Check and Update Board side infor, please wait about 1~3 min.");
                            listProperty = CommonTools.MyGetBlockAssemblyInfor(assBlock, assBlock.ToString());//递归迭代
                            Excel.Application xlapp = new Excel.Application();
                            Excel.Workbook wb = xlapp.Workbooks.Add();
                            Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                            ws.Range["A1:e1"].Value = new string[] { "Assembly Name", "Assembly Type", "Old Side", "New  Side ","Remark" };
                            sw.Stop();
                            base.WriteStatusBarMsg("Changed BoardInfor Complete !  total cost time --->" + sw.Elapsed );
                            for (int i = 0; i <= listProperty.Count-1; i++)
                            {
                                var item = listProperty[i];
                                int k = i + 2;
                                switch (item[1])
                                {
                                    case "S":
                                        ws.Range["B" + k].Value = "Sub Block"; break;
                                    case "A":
                                        ws.Range["B" + k].Value = "Assembly"; break;
                                    case "PA":
                                        ws.Range["B" + k].Value = "Pre-Assembly"; break;
                                    case "SA":
                                        ws.Range["B" + k].Value = "Seat-Assembly"; break;
                                }
                                if (item[0].LastIndexOf(")")!=0)
                                {
                                    ws.Range["A" + k].Value = item[0];
                                    ws.Range["C" + k].Value = item[2];
                                    ws.Range["D" + k].Value = item[3];
                                    if (item[2] == item[3])
                                    {
                                        ws.Range["E" + k].Value = "Correct Board Side Infor";
                                        ws.Range["E" + k].Interior.ColorIndex = 4;
                                    }
                                    else
                                    {
                                        ws.Range["E" + k].Value = "Update by This Plug";
                                        ws.Range["E" + k].Interior.ColorIndex =8;
                                    }
  
                                }
                                else
                                {
                                    ws.Range["A" + k].Value = item[0];
                                    ws.Range["A" + k].Interior.ColorIndex = 6;
                                    ws.Range["C" + k].Value = item[2];
                                    ws.Range["D" + k].Value = item[3];
                                    ws.Range["E" + k].Value = "Need Manually Setting";
                                }
                               
                            }
                            ws.Range["A1:E1"].EntireColumn.AutoFit();
                            xlapp.Visible = true;
                            xlapp.WindowState = Excel.XlWindowState.xlMaximized;
                        }
                        else
                        {
                            base.WriteStatusBarMsg("Please selected the Block folder!");
                        }
                    }
                    else
                    {
                        base.WriteStatusBarMsg("Please selected the Block folder!");
                    }
                }
                else
                {
                    base.WriteStatusBarMsg("Please selected the Block folder!");
                }
            }
            catch (Exception ex)
            {
                base.WriteStatusBarMsg(ex.Message);
            }
        }
    }
}



