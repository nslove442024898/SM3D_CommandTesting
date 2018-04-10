
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
    public class ShowObjectAllPro : BaseModalCommand
    {
        public override void OnStart(int commandID, object argument)
        {
            Excel.Application xlapp = new Excel.Application();
            try
            {
                SelectSet ss = ClientServiceProvider.SelectSet;
                if (ss.SelectedObjects.Count == 1)
                {
                    var boj = ss.SelectedObjects.ToList()[0];
                    Excel.Workbook wb = xlapp.Workbooks.Add();
                    Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                    ws.Range["A1"].Value = boj.ClassInfo.DisplayName;
                    ws.Range["A2:B2"].Value = new string[] { "Property Name", "PropertyValue"};
                    var temp = (boj.GetAllProperties().Where(c => c.ToString()!="")).ToList();
                    for (int i = 0; i <= temp.Count-1; i++)
                    {
                        int k = i + 3;
                        var item = temp[i];
                        ws.Range["A" + k].Value = item.PropertyInfo.Name;
                        ws.Range["B" + k].Value = item.ToString();
                    }
                    ws.Range["A1:B1"].EntireColumn.AutoFit();
                    xlapp.Visible = true;
                    xlapp.WindowState = Excel.XlWindowState.xlMaximized;
                }
                
            }
            catch (Exception ex)
            {
                base.WriteStatusBarMsg(ex.Message);
                xlapp = null;
            }
        }
    }
}
