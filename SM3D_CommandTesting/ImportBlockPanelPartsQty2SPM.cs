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
    public class ImportBlockPanelPartsQty2SPM : BaseModalCommand
    {
        FormForPartsQtyImport frm;
        public string projectName;
        public string blockName;
        Dictionary<string, int> panelPartQtyDic = new Dictionary<string, int>();
        public override void OnStart(int commandID, object argument)
        {
            SelectSet ss = ClientServiceProvider.SelectSet;

            // GetProject Name
            WorkingSet currWS = ClientServiceProvider.WorkingSet;
            SP3DConnection sp3dCon = currWS.ActiveConnection;
            projectName = sp3dCon.Name.Substring(0, sp3dCon.Name.IndexOf('_'));
            base.WriteStatusBarMsg("please wait a moment, now program was export parts qty from sm3d modeling");
            if (ss.SelectedObjects.Count == 1)
            {
                BusinessObject boj = ss.SelectedObjects[0] as BusinessObject;
                if (boj.ClassInfo.DisplayName == "Assembly")
                {
                    if (boj is IAssembly)
                    {
                        if (boj.MyGetProperty("Type") == "B")
                        {
                            blockName = boj.ToString();
                            panelPartQtyDic = MyGetBlockPanelPartQty(boj as IAssembly);
                            if (panelPartQtyDic.Count > 0)
                            {
                                if (frm == null)
                                {
                                    base.WriteStatusBarMsg("click the green  button to import parts qty to Spm !");
                                    frm = new FormForPartsQtyImport();
                                    frm.MyPartsQty = this.panelPartQtyDic;
                                    frm.ProjectName = this.projectName;
                                    frm.BlockName = this.blockName;
                                    frm.FormClosed += Frm_FormClosed;
                                    frm.ShowDialog();
                                }
                                else frm.Activate();
                            }
                            else base.WriteStatusBarMsg("No Parts In select Block!");
                        }
                        else base.WriteStatusBarMsg("Please select the Block Assembly Folder!");
                    }
                    else base.WriteStatusBarMsg("Please select the Block Assembly Folder!");
                }
                else base.WriteStatusBarMsg("Please select the Block Assembly Folder!");
            }
            else base.WriteStatusBarMsg("Please select the Block Assembly Folder!");
        }

        private void Frm_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            frm = null;
        }

        public Dictionary<string, int> MyGetBlockPanelPartQty(IAssembly assBlock)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();
            //panel文件夹---->20180220符合条件的
            var assPanels = assBlock.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly" && (c as BusinessObject).MyGetProperty("Type") == "S");
            //panel文件夹---->20180220
            foreach (var assPanel in assPanels.ToList())
            {
                //if (assPanel.ToString()=="TB82A(S)")
                //遍历每个panel
                int qty = 0;
                foreach (var item in (assPanel as IAssembly).AssemblyChildren)
                {
                    if ((item as BusinessObject).ClassInfo.DisplayName == "Assembly")
                    {
                        qty += GetAssemblyFolderPartQty(item as IAssembly);
                        //直接在panel的子文件加下的零件的集合
                    }
                    //CollarPart,ProfilePart,PlatePart,StiffenerPart,"CollarPart
                    else if ((item as BusinessObject).ClassInfo.DisplayName.EndsWith("ePart") ||
                        (item as BusinessObject).ClassInfo.DisplayName.EndsWith("rPart"))
                    {
                        qty++;
                    }
                    else
                    {
                        DebugLogMessage((item as BusinessObject).ToString()+"===>"+ (item as BusinessObject).ClassInfo.DisplayName);
                    }
                }
                dic[assBlock.ToString() + "/" + assPanel.ToString()] = qty;
            }
            return dic;
        }
        public int GetAssemblyFolderPartQty(IAssembly assFolder)
        {
            int res = 0;
            var ass1 = assFolder.AssemblyChildren.Where(c => ((c as BusinessObject).ClassInfo.DisplayName == "Assembly"));
            if (ass1.ToList().Count == 0)
            {
                var aloneParts = assFolder.AssemblyChildren.ToList().Where(c => ((c as BusinessObject).ClassInfo.DisplayName.EndsWith("ePart")) ||
                ((c as BusinessObject).ClassInfo.DisplayName.EndsWith("rPart")));
                res += aloneParts.ToList().Count;
            }
            else
            {
                foreach (var item in ass1.ToList())
                {
                    res += GetAssemblyFolderPartQty(item as IAssembly);
                }
            }
            return res;
        }
    }
}
