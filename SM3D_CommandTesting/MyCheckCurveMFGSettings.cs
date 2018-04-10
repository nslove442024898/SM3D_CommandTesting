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
    public class MyCheckCurveMFGSettings : BaseModalCommand
    {
        public override void OnStart(int commandID, object argument)
        {
            SelectSet ss = ClientServiceProvider.SelectSet;
            //SP3DConnection sp3dc = ClientServiceProvider.WorkingSet.ActiveConnection;
            List<string[]> listCurvePlates = new List<string[]>();
            List<string[]> listCurveProfiles = new List<string[]>();
            try
            {
                if (ss.SelectedObjects.Count > 0)
                {
                    foreach (var item1 in ss.SelectedObjects)
                    {
                        if (item1.MyGetProperty("Type") == "S" && item1.ClassInfo.DisplayName == "Assembly")
                        {
                            var panelAss = item1 as IAssembly;
                            //弯曲板材
                            base.WriteStatusBarMsg($"get panel {item1.ToString()} curve plate and profile information.");
                            var panelParts = panelAss.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "PlatePart");
                            if (panelParts.ToList().Count > 0)
                            {
                                var curvePlates = panelParts.Where(c => ((!((c as PlatePart).Curvature == Curvature.Flat || 
                                (c as PlatePart).Curvature == Curvature.Knuckled)) && (c as PlatePart).Type == PlateType.Hull));
                                //
                                foreach (var item in curvePlates.ToList())
                                {
                                    listCurvePlates.Add(MyGetCurvePlateInfor(item));
                                }
                            }
                            //弯曲型材
                            var panelProfile = panelAss.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "StiffenerPart");
                            if (panelProfile.ToList().Count > 0)
                            {
                                var curveProfiles = panelProfile.Where(c => (c as BusinessObject).MyGetProperty("Curved") != "Straight" && (c as BusinessObject).MyGetProperty("NamingGroup") == "HP");
                                foreach (var item in curveProfiles.ToList())
                                {
                                    listCurveProfiles.Add(MyGetCurveProfileInfor(item));
                                }
                            }
                            base.WriteStatusBarMsg($"complete get panel {item1.ToString()} curve plate and profile information.");
                        }
                        else WriteStatusBarMsg("please select a panel folder!");
                    }
                }
            }

            catch (Exception ex)
            {
                base.WriteStatusBarMsg(ex.Message);
            }

            if (listCurvePlates.Count > 0 && listCurveProfiles.Count > 0)
            {
                base.WriteStatusBarMsg($"complete get all select panels curve plates and profiles information, now export the result to excel !");
                Excel.Application xlapp = new Excel.Application();
                Excel.Workbook wb = xlapp.Workbooks.Add();
                Excel.Worksheet ws = (wb.Sheets[1]) as Excel.Worksheet;
                ws.Range["A1:I1"].Value = new string[] { "Mfg Name", "Frame System", "PlateUpside", "PlPlateLocation ", "PlProfileLocation", "Template Set", "Frame System", "Side", "Orientation" };
                int flag = 0;
                for (int i = 0; i < listCurvePlates.Count - 1; i++)
                {
                    int K = i + 2;
                    ws.Range["A" + K + ":I" + K].Value = listCurvePlates[i];
                    //
                    if (ws.Range["B" + K].Value == "未设置")
                    {
                        ws.Range["B" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["B" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["C" + K].Value != "TemplateSide")
                    {
                        ws.Range["C" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["C" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["d" + K].Value != "Triangle Thickness Direction")
                    {
                        ws.Range["d" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["d" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["e" + K].Value != "Triangle Thickness Direction")
                    {
                        ws.Range["e" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["e" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["g" + K].Value == "未设置")
                    {
                        ws.Range["g" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["g" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["h" + K].Value != "MoldedSide")
                    {
                        ws.Range["h" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["h" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["I" + K].Value != "Perpendicular")
                    {
                        ws.Range["I" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["I" + K].Interior.ColorIndex = 4;

                    flag = K;
                }
                flag++;
                ws.Range["A" + flag + ":C" + flag].Value = new string[] { "MfgProfile Name", "Frame System","Marking Line Label" };
                for (int i = 0; i < listCurveProfiles.Count - 1; i++)
                {
                    int K = flag + i + 1;
                    ws.Range["A" + K + ":C" + K].Value = listCurveProfiles[i];
                    if (ws.Range["B" + K].Value == "未设置")
                    {
                        ws.Range["B" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["B" + K].Interior.ColorIndex = 4;
                    //
                    if (ws.Range["C" + K].Value == "未做")
                    {
                        ws.Range["C" + K].Interior.ColorIndex = 3;
                    }
                    else ws.Range["C" + K].Interior.ColorIndex = 4;
                }
                ws.Range["A1:I1"].EntireColumn.AutoFit();
                xlapp.Visible = true;
                xlapp.WindowState = Excel.XlWindowState.xlMaximized;
            }
            else
            {
                base.WriteStatusBarMsg("Can't find any curve plates in select panels   !");
            }
        }

        private static string[] MyGetCurvePlateInfor(IAssemblyChild item)
        {
            string[] str1 = new string[9];
            if (((BusinessObject)item) is PlatePart)
            {
                PlatePart PlatePart = ((BusinessObject)item) as PlatePart;
                var plateMFG = PlatePart.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Manufacturing Plate Part");
                if (plateMFG.ToList().Count>0)
                {
                    ManufacturingPlate mfg = (plateMFG.ToList()[0]) as ManufacturingPlate;
                    str1[0] = mfg.Name;
                    if (mfg.CoordinateSystem == null)
                    {
                        str1[1] = "未设置";
                    }
                    else
                    {
                        str1[1] = mfg.CoordinateSystem.Name;
                    }

                    str1[2] = (mfg.GetSettingValue(SettingType.Process, "IJUAMfgPlateProdProcess", "PlateUpside").Name);
                    str1[3] = (mfg.GetSettingValue(SettingType.Marking, "IJUAMfgPlateMarkingFaceLoc", "PlPlateLocation").Name);
                    str1[4] = (mfg.GetSettingValue(SettingType.Marking, "IJUAMfgPlateMarkingFaceLoc", "PlProfileLocation").Name);
                    var plateTemplateSets = PlatePart.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Template Set");
                    if (plateTemplateSets.ToList().Count > 0 && plateTemplateSets != null)
                    {
                        TemplateSet ts = (plateTemplateSets.ToList()[0]) as TemplateSet;
                        str1[5] = (ts.Name);
                        if (ts.CoordinateSystem == null)
                        {
                            str1[6] = ("未设置");
                        }
                        else
                        {
                            str1[6] = (ts.CoordinateSystem.Name);
                        }
                        str1[7] = (ts.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Side").Name);
                        str1[8] = (ts.GetSettingValue(SettingType.Process, "IJUAMfgTemplateProcessPlate", "Orientation").Name);
                    }
                    else
                    {
                        str1[5] = "Template 未做";
                        str1[6] = "Template 未做";
                        str1[7] = "Template 未做";
                        str1[8] = "Template 未做";
                    }
                }
                else
                {
                    str1[0] = PlatePart.Name;
                    str1[1] = "MFG 未做";
                    str1[2] = "MFG 未做";
                    str1[3] = "MFG 未做";
                    str1[4] = "MFG 未做";
                    str1[5] = "MFG 未做";
                    str1[6] = "MFG 未做";
                    str1[7] = "MFG 未做";
                    str1[8] = "MFG 未做";
                }
            }
            return str1;
        }

        private static string[] MyGetCurveProfileInfor(IAssemblyChild item)
        {
            string[] str1 = new string[3];
            if (((BusinessObject)item) is ProfilePart)
            {
                ProfilePart profilePart = ((BusinessObject)item) as ProfilePart;
                var profileMFG = profilePart.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Manufacturing Profile Part");
                var profileMarking= profilePart.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Mfg Marking Folder"); 
                if (profileMFG.ToList().Count>0)
                {
                    ManufacturingProfile mfg = (profileMFG.ToList()[0]) as ManufacturingProfile;
                    str1[0] = (mfg.Name);
                    if (mfg.CoordinateSystem == null)
                    {
                        str1[1] = "未设置";
                    }
                    else
                    {
                        str1[1] = (mfg.CoordinateSystem.Name);
                    }
                }
                if (profileMarking.ToList().Count > 0)
                {
                   str1[2] = "已做";
                }else str1[2] = "未做";
            }
            return str1;
        }
    }
}
