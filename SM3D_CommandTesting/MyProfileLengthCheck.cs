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
using Ingr.SP3D.Planning.Client;
using Ingr.SP3D.Planning.Client.Services;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Planning.Middle.Services;

//

namespace MyNameSpace
{
    public class MyProfileLengthCheck : BaseModalCommand
    {
        public List<StiffenerPart> Stiff=new List<StiffenerPart>();
        FrmPanelProfileLengthCheck frm;
        public override void OnStart(int commandID, object argument)
        {
            base.OnStart(commandID, argument);
            //CommonPartsGroup cpg =
            SelectSet ss = ClientServiceProvider.SelectSet;
            if (ss.SelectedObjects.Count == 1 && ss.SelectedObjects[0].ClassInfo.DisplayName == "Assembly")
            {
                var bojPanel = ss.SelectedObjects[0];
                if (bojPanel.MyGetProperty("Type") == "S")
                {
                    var temp = GetStiffParts(bojPanel);
                    if (temp.Count>0)
                    {
                        Stiff.AddRange(temp);
                        //
                        if (frm == null)
                        {
                            frm = new FrmPanelProfileLengthCheck();
                            frm.stiff = this.Stiff;
                            frm.Panel = bojPanel.ToString();
                            frm.FormClosed += Frm_FormClosed; ;
                            frm.ShowDialog();
                        }
                        else
                        {
                            frm.Activate();
                        }
                        //
                    }
                    else
                    {
                        base.WriteStatusBarMsg("Can't find any Bulb Flat or Angle In currctly active Panels ");
                    }
                }
                else
                {
                    base.WriteStatusBarMsg("Please Select the Panel Folder!");
                }
            }
            else
            {
                base.WriteStatusBarMsg("Please Select the Panel Folder!");
            }
        }

        private void Frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            frm = null;
        }

        private List<StiffenerPart> GetStiffParts(BusinessObject bojPanel)
        {
            List<StiffenerPart> temp = new List<StiffenerPart>();
            var profileSets = (bojPanel as IAssembly).AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "StiffenerPart");
            var ass = (bojPanel as IAssembly).AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly");
            if (ass.ToList().Count == 0 && profileSets.ToList().Count>0)
            {
                foreach (var item in profileSets.ToArray())
                {
                    if ((item as StiffenerPart).CrossSectionTypeName == "BulbFlat" || (item as StiffenerPart).CrossSectionTypeName == "UnequalAngle" || (item as StiffenerPart).CrossSectionTypeName == "EqualAngle")
                    {
                        temp.Add(item as StiffenerPart);
                    }
                }
            }
            else if (ass.ToList().Count > 0 && profileSets.ToList().Count > 0)
            {
                foreach (var item in profileSets.ToArray())
                {
                    if ((item as StiffenerPart).CrossSectionTypeName == "BulbFlat" || (item as StiffenerPart).CrossSectionTypeName == "UnequalAngle" || (item as StiffenerPart).CrossSectionTypeName == "EqualAngle")
                    {
                        temp.Add(item as StiffenerPart);
                    }
                }
                foreach (var item in ass.ToList())
                {
                    temp.AddRange(GetStiffParts(item as BusinessObject));
                }
            }
            else
            {
                foreach (var item in ass.ToList())
                {
                    temp.AddRange(GetStiffParts(item as BusinessObject));
                }
            }
            return temp;
        }
    }
}
