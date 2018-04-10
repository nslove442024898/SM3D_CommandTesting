using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
// 
using Ingr.SP3D.Manufacturing.Middle.Services;
using Ingr.SP3D.Structure.Middle.Services;
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Structure.Middle;
//
using Ingr.SP3D.Common.Client;
using Ingr.SP3D.Common.Client.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
//
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace MyNameSpace
{
    public class PGCheck : BaseModalCommand
    {
        public override void OnStart(int commandID, object argument)
        {
            Excel.Application xlapp = new Excel.Application();
            try
            {
                List<string[]> pgInfor = new List<string[]>();
                SelectSet ss = ClientServiceProvider.SelectSet;
                WorkingSet currws = ClientServiceProvider.WorkingSet;
                PermissionGroup currPG = currws.ActiveConnection.ActivePermissionGroup;
                base.WriteStatusBarMsg($"Check All Object whether PG was --->{currPG.Name},please wait a moment");
                if (ss.SelectedObjects.Count == 1)
                    foreach (BusinessObject item in ss.SelectedObjects)
                    {
                        if (item is IAssembly && item is BusinessObject)
                        {
                            var assBlock = item as BusinessObject;
                            pgInfor.Add(new string[] { assBlock.ClassInfo.DisplayName,
                                assBlock.ToString(),
                            assBlock.PermissionGroup.ToString(),
                            assBlock.PermissionGroup.PermissionGroupID.ToString() });
                            var temp = GetAllObjectPg(assBlock);
                            if (temp != null)
                            {
                                pgInfor.AddRange(GetAllObjectPg(assBlock));
                            }
                            else
                            {
                                base.WriteStatusBarMsg("please selected a Block Folder");
                            }
                            //单独做一个地递归
                        }
                    }
                base.WriteStatusBarMsg("Complete Check PG, now Export Check Report !");
                Excel.Workbook wb = xlapp.Workbooks.Add();
                Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                ws.Range["A1:D1"].Value = new string[] { "Object Type", "Object Name", "Pg Name", "Pg NameID" };
                for (int i = 0; i <= pgInfor.Count - 1; i++)
                {
                    var item = pgInfor[i];
                    int k = i + 2;
                    ws.Range["A" + k + ":D" + k].Value = item;
                    if (ws.Range["C" + k].Value != currPG.Name)
                    {
                        ws.Range["C" + k].Interior.ColorIndex = 3;
                        ws.Range["D" + k].Interior.ColorIndex = 3;
                    }
                    else
                    {
                        ws.Range["C" + k].Interior.ColorIndex = 4;
                        ws.Range["D" + k].Interior.ColorIndex = 4;
                    }
                }
                ws.Range["A1:D1"].EntireColumn.AutoFit();
                xlapp.Visible = true;
                xlapp.WindowState = Excel.XlWindowState.xlMaximized;
            }
            catch (Exception ex)
            {
                base.WriteStatusBarMsg(ex.Message);
                xlapp = null;
            }
        }
        private List<string[]> GetAllObjectPg(BusinessObject bojBlock)
        {
            List<string[]> str = new List<string[]>();
            if (bojBlock.ClassInfo.DisplayName == "Assembly" && bojBlock.MyGetProperty("Type") == "B")
            {
                var assBlock = bojBlock as IAssembly;
                if (assBlock.AssemblyChildren.Count > 0)
                {
                    foreach (var item in assBlock.AssemblyChildren)
                    {
                        if (item is IAssembly)
                        {
                            var bojPanel = item as BusinessObject;
                            if (bojPanel.MyGetProperty("Type") == "S")//限定了只能是sub block的panel文件夹
                            {
                                //     str.Add(new string[] { ((BusinessObject)item).ClassInfo.DisplayName,
                                //    ((BusinessObject)item).ToString(),
                                //((BusinessObject)item).PermissionGroup.ToString(),
                                // ((BusinessObject)item).PermissionGroup.PermissionGroupID.ToString() });
                                str.AddRange(GetAllAssObjectPg(item as IAssembly));
                            }
                        }
                        else
                        {
                            str.Add(new string[] { ((BusinessObject)item).ClassInfo.DisplayName,
                               ((BusinessObject)item).ToString(),
                           ((BusinessObject)item).PermissionGroup.ToString(),
                            ((BusinessObject)item).PermissionGroup.PermissionGroupID.ToString() });
                        }
                    }
                }
            }
            else
            {
                str = null;
            }
            return str;
        }
        private List<string[]> GetAllAssObjectPg(IAssembly assPanel)//获取每个panel assembly
        {
            List<string[]> str = new List<string[]>();
            if (assPanel.AssemblyChildren.Count > 0)
            {
                str.Add(new string[] { ((BusinessObject)assPanel).ClassInfo.DisplayName,
                               ((BusinessObject)assPanel).ToString(),
                           ((BusinessObject)assPanel).PermissionGroup.ToString(),
                            ((BusinessObject)assPanel).PermissionGroup.PermissionGroupID.ToString() });
                foreach (var item in assPanel.AssemblyChildren)
                {
                    if (item is IAssembly)
                    {
                        var strTemp = GetAllAssObjectPg(item as IAssembly);
                        str.AddRange(strTemp);
                    }
                    else
                    {
                        str.Add(new string[] { ((BusinessObject)item).ClassInfo.DisplayName,
                               ((BusinessObject)item).ToString(),
                           ((BusinessObject)item).PermissionGroup.ToString(),
                            ((BusinessObject)item).PermissionGroup.PermissionGroupID.ToString() });
                    }
                }
            }
            else
            {
                str.Add(new string[] { ((BusinessObject)assPanel).ClassInfo.DisplayName,
                            ((BusinessObject)assPanel).ToString(),
                            ((BusinessObject)assPanel).PermissionGroup.ToString(),
                            ((BusinessObject)assPanel).PermissionGroup.PermissionGroupID.ToString()});
            }
            return str;
        }
    }
}

