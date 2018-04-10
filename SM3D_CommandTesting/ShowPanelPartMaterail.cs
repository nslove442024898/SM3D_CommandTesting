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
    public class ShowPanelPartMaterail : BaseModalCommand
    {
        PanelMaterialList frm;
        string panelName;
        List<string[]> listPlateParts = new List<string[]>();
        List<string[]> listStiffParts = new List<string[]>();
        public override void OnStart(int commandID, object argument)
        {
            SelectSet ss = ClientServiceProvider.SelectSet;
            WorkingSet currWS = ClientServiceProvider.WorkingSet;
            SP3DConnection sp3dCon = currWS.ActiveConnection;
            var projectName = sp3dCon.Name.Substring(0, sp3dCon.Name.IndexOf('_'));
            if (ss.SelectedObjects.Count == 1 && ss.SelectedObjects[0].ClassInfo.DisplayName == "Assembly")
            {
                var bojPanel = ss.SelectedObjects[0];

                if (bojPanel.MyGetProperty("Type") == "S")
                {
                    panelName = bojPanel.ToString();
                    base.WriteStatusBarMsg("Start to get all parts material for select panel. please wait a moment!");
                    if (bojPanel is IAssembly)
                    {
                        var temp1 = GetAssemblyMaterialList(bojPanel as IAssembly);
                        this.listPlateParts = temp1[0];
                        this.listStiffParts = temp1[1];
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
            if (frm == null)
            {
                frm = new PanelMaterialList();
                frm.FormClosed += Frm_FormClosed;
                frm.listPlateParts = this.listPlateParts;
                frm.listStiffParts = this.listStiffParts;
                frm.Block = this.panelName;
                frm.HullNumber = projectName;
                frm.ShowDialog();
            }
            else frm.Activate();

        }

        private List<string[]>[] GetAssemblyMaterialList(IAssembly assPanel)
        {
            List<string[]>[] str1= { new List<string[]>(), new List<string[]>()};
            var parts = assPanel.AssemblyChildren.Where(c =>
            (c as BusinessObject).ClassInfo.DisplayName.EndsWith("ePart") || (c as BusinessObject).ClassInfo.DisplayName.EndsWith("rPart"));
            var assFolders = assPanel.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly");
            if (assFolders.ToList().Count != 0)
            {
                foreach (var item in assFolders)
                {
                    var temp = GetAssemblyMaterialList(item as IAssembly);
                    str1[0].AddRange(temp[0]);
                    str1[1].AddRange(temp[1]);
                }
                foreach (var item in parts)
                {
                    switch ((item as BusinessObject).ClassInfo.DisplayName)
                    {
                        //CollarPart,ProfilePart,PlatePart,StiffenerPart,ERProfilePart
                        case "PlatePart":
                            var platePart = item as PlatePart;
                            str1[0].Add(new string[] { platePart.Name,platePart.MaterialGrade,(1000*platePart.Thickness).ToString(),
                                        platePart.MyGetProperty("PlateLength"), platePart.MyGetProperty("PlateWidth")
                                        ,platePart.MyGetProperty("Area"),platePart.MyGetProperty("DryWeight"),platePart.MyGetProperty("Curved")});
                            break;
                        case "CollarPart":
                            var collarPart = item as CollarPart;
                            str1[0].Add(new string[] { collarPart.Name,collarPart.MaterialGrade,(1000*collarPart.Thickness).ToString(),
                                        collarPart.MyGetProperty("PlateLength"), collarPart.MyGetProperty("PlateWidth")
                                        ,collarPart.MyGetProperty("Area"),collarPart.MyGetProperty("DryWeight"),collarPart.MyGetProperty("Curved")});
                            break;
                        case "StiffenerPart":
                            var stiffPart = item as StiffenerPart;
                            str1[1].Add(new string[] { stiffPart.Name,stiffPart.MaterialGrade,(stiffPart.SectionName).ToString(),
                                        stiffPart.MyGetProperty("ProfileLength"),stiffPart.MyGetProperty("DryWeight"),stiffPart.MyGetProperty("Curved") });
                            break;
                        case "ProfilePart":
                        case "ERProfilePart":
                            var proPart = item as ProfilePart;
                            str1[1].Add(new string[] { proPart.Name,proPart.MaterialGrade,(proPart.SectionName).ToString(),
                                        proPart.MyGetProperty("ProfileLength"),proPart.MyGetProperty("DryWeight"),proPart.MyGetProperty("Curved") });
                            break;
                    }

                }
            }
            else
            {
                foreach (var item in parts)
                {
                    switch ((item as BusinessObject).ClassInfo.DisplayName)
                    {
                        //CollarPart,ProfilePart,PlatePart,StiffenerPart
                        case "PlatePart":
                            var platePart = item as PlatePart;
                            str1[0].Add(new string[] { platePart.Name,platePart.MaterialGrade,(1000*platePart.Thickness).ToString(),
                                        platePart.MyGetProperty("PlateLength"), platePart.MyGetProperty("PlateWidth")
                                        ,platePart.MyGetProperty("Area"),platePart.MyGetProperty("DryWeight"),platePart.MyGetProperty("Curved")});
                            break;
                        case "CollarPart":
                            var collarPart = item as CollarPart;
                            str1[0].Add(new string[] { collarPart.Name,collarPart.MaterialGrade,(1000*collarPart.Thickness).ToString(),
                                        collarPart.MyGetProperty("PlateLength"), collarPart.MyGetProperty("PlateWidth")
                                        ,collarPart.MyGetProperty("Area"),collarPart.MyGetProperty("DryWeight"),collarPart.MyGetProperty("Curved")});
                            break;
                        case "StiffenerPart":
                            var stiffPart = item as StiffenerPart;
                            str1[1].Add(new string[] { stiffPart.Name,stiffPart.MaterialGrade,(stiffPart.SectionName).ToString(),
                                        stiffPart.MyGetProperty("ProfileLength"),stiffPart.MyGetProperty("DryWeight"),stiffPart.MyGetProperty("Curved") });
                            break;
                        case "ProfilePart":
                        case "ERProfilePart":
                            var proPart = item as ProfilePart;
                            str1[1].Add(new string[] { proPart.Name,proPart.MaterialGrade,(proPart.SectionName).ToString(),
                                        proPart.MyGetProperty("ProfileLength"),proPart.MyGetProperty("DryWeight"),proPart.MyGetProperty("Curved") });
                            break;
                    }

                }
            }
            return str1;
        }

        private void Frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            frm = null;
        }
    }
}
