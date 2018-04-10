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

namespace MyNameSpace
{
    public class MyPanelManufacutringCheck : BaseModalCommand
    {
        Panel_Manufacturing_Check frm;
        public string ProjectName { get; set; }
        public string BlockName { get; set; }
        public Dictionary<string, int> panelMfg { get; set; }
        public override void OnStart(int commandID, object argument)
        {
            SelectSet ss = ClientServiceProvider.SelectSet;
            WorkingSet currWS = ClientServiceProvider.WorkingSet;
            SP3DConnection sp3dCon = currWS.ActiveConnection;
            this.ProjectName = sp3dCon.Name.Substring(0, sp3dCon.Name.IndexOf('_'));
            if (ss.SelectedObjects.Count == 1 && ss.SelectedObjects[0].ClassInfo.DisplayName == "Assembly")
            {
                var bojBlock = ss.SelectedObjects[0];
                if (bojBlock.MyGetProperty("Type") == "B")
                {
                    this.BlockName = bojBlock.ToString();
                    base.WriteStatusBarMsg("Start to get all panel's Panel Mfg. Status. please wait a moment!");
                    if (bojBlock is IAssembly)
                    {
                        this.panelMfg = GetAssemblyPanelMfg(bojBlock as IAssembly);
                    }
                }
                else
                {
                    base.WriteStatusBarMsg("Please Select the Block Folder!");
                }
            }
            else
            {
                base.WriteStatusBarMsg("Please Select the Block Folder!");
            }
            if (frm == null)
            {
                frm = new Panel_Manufacturing_Check();
                frm.FormClosed += Frm_FormClosed;
                frm.Hull = this.ProjectName;
                frm.BLK = this.BlockName;
                frm.panelMfg = this.panelMfg;
                frm.ShowDialog();
            }
            else frm.Activate();
        }
        private Dictionary<string, int> GetAssemblyPanelMfg(IAssembly ass)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();
            var panels = ass.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly" && (c as BusinessObject).MyGetProperty("Type") == "S");
            foreach (var item in panels.ToList())
            {
                var temp = (item as IAssembly).AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Manufacturing Plate Part");
                dic[item.ToString()] = temp.ToArray().Length;
            }
            return dic;
        }

        private void Frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            frm = null;
        }
    }
}

