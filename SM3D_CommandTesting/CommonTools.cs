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

namespace MyNameSpace
{
    public static class CommonTools
    {
        public static string MyGetProperty(this BusinessObject boj, string proValue)
        {
            string str = "";
            try
            {

                var assPanelType = boj.GetAllProperties().Where(c => c.PropertyInfo.Name == proValue);
                if (assPanelType.ToList().Count == 1)
                {
                    str = boj.GetPropertyValue(assPanelType.ToList()[0].PropertyInfo).ToString();
                }
                else
                {
                    str = "";
                }
            }
            catch (Exception)
            {

                str = "";
            }

            return str;
        }
        public static PropertyInformation MyGetPropertyInfor(this BusinessObject boj, string proValue)
        {
            PropertyInformation str = null;
            var assPanelType = boj.GetAllProperties().Where(c => c.PropertyInfo.Name == proValue);
            if (assPanelType.ToList().Count == 1)
            {
                str = assPanelType.ToList()[0].PropertyInfo;
            }
            else
            {
                str = null;
            }
            return str;
        }
        public static int MyBoardSideinfor2Int(string boardSideString)
        {
            int i = 0;
            switch (boardSideString)
            {
                case "P":
                    i = 3;
                    break;
                case "C":
                    i = 1;
                    break;
                case "S":
                    i = 2;
                    break;
                default:
                    i = 1;
                    break;
            }
            return i;
        }
        public static string MyIntBoardSideinfor(int i)
        {
            string str = "";
            switch (i)
            {
                case 1:
                    str = "C";
                    break;
                case 2:
                    str = "S";
                    break;
                case 3:
                    str = "P";
                    break;
                default:
                    str = "";
                    break;
            }
            return str;
        }
        public static List<string[]> MyGetBlockAssemblyInfor(IAssembly assObj, string ParentName)
        {
            List<string[]> str = new List<string[]>();
            BusinessObject boj = assObj as BusinessObject;//将block的assembly对象转化成bobj

            var panels = assObj.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly" && (c as BusinessObject).MyGetProperty("Type")== "S");

            foreach (var item in panels.ToList())//遍历整个panel
            {
                BusinessObject bojPanel = item as BusinessObject;
                //if (bojPanel.UserClassInfo.DisplayName=="Assembly")//这层判断有问题啊-20180205
                //{
                string correctSide = "C", sideInSm3d = MyIntBoardSideinfor(int.Parse(bojPanel.MyGetProperty("BoardSide")));
                //check不是assembly的对象
                if (bojPanel.ToString().EndsWith(")"))
                {
                    correctSide = bojPanel.ToString().Substring(
                        bojPanel.ToString().LastIndexOf("(") + 1, 1);
                }
                //panel 文件夹层次的board side infor检查与更正
                if (sideInSm3d != correctSide)
                {
                    bojPanel.SetPropertyValue(MyBoardSideinfor2Int(correctSide), bojPanel.MyGetPropertyInfor("BoardSide"));
                    ClientServiceProvider.TransactionMgr.Commit("FilterSieve");
                }
                var strTemp = new string[] {ParentName+"/"+bojPanel.ToString(),
                    bojPanel.MyGetProperty("Type"),sideInSm3d,correctSide};
                str.Add(strTemp);
                //panel下面文件夹层次的board side infor检查与更正
                if (bojPanel.MyGetProperty("Type") == "S" && (bojPanel as IAssembly).AssemblyChildren.Count > 0)
                {
                    foreach (var item2 in (bojPanel as IAssembly).AssemblyChildren)//每个panel的子文件夹
                    {
                        if (item2 is IAssembly && (item2 as BusinessObject).ClassInfo.DisplayName == "Assembly")
                        {
                            var strTemp2 = MyGetPanelAssemblyInfor(item2 as IAssembly,
                                ParentName + "/" + bojPanel.ToString(),
                                correctSide);//获取子文件下的对象
                            if (strTemp2 != null && strTemp2.Count > 0)
                            {
                                foreach (var item3 in strTemp2) str.Add(item3);
                            }
                            else continue;
                        }
                        else continue;
                    }
                }
                //}
                //else continue;
            }
            return str;
        }
        public static List<string[]> MyGetPanelAssemblyInfor(IAssembly assObj, string PanelName, string correctSide)
        {
            List<string[]> str = new List<string[]>();
            BusinessObject subAssFolder = assObj as BusinessObject;
            string sideInfor = MyIntBoardSideinfor(int.Parse(subAssFolder.MyGetProperty("BoardSide")));
            var flag = (assObj as IAssembly).AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly");
            if (flag.ToList().Count > 0)
            {
                foreach (var item in flag.ToList())
                {
                    BusinessObject subAssFolder1 = item as BusinessObject;
                    if (subAssFolder1.ClassInfo.DisplayName == "Assembly")
                    {
                        var strTemp = MyGetPanelAssemblyInfor((item as IAssembly), PanelName + "-" + subAssFolder1.ToString(), correctSide);
                        if (strTemp != null)
                        {
                            foreach (var item1 in strTemp) str.Add(item1);
                        }
                    }
                }
            }
            else
            {
                if (sideInfor != correctSide)
                {
                    subAssFolder.SetPropertyValue(MyBoardSideinfor2Int(correctSide),
                        subAssFolder.MyGetPropertyInfor("BoardSide"));
                    ClientServiceProvider.TransactionMgr.Commit("FilterSieve");
                }
                var strTemp = new string[] {PanelName+"-"+subAssFolder.ToString(),
                    subAssFolder.MyGetProperty("Type"),sideInfor,correctSide};
                str.Add(strTemp);
            }
            return str;
        }
        public static double[] MyGetShopDrawingKpi(int Kpin)
        {
            double[] kpi = new double[4];
            // "Done", "Check", "koike Done", "koike check")
            if (Kpin >= 0 && Kpin <= 350)
            {
                if (Kpin > 0 && Kpin <= 10) { kpi[0] = 3.9; kpi[1] = 1.4; kpi[2] = 1.6; kpi[3] = 0.5; }
                if (Kpin >= 11 && Kpin <= 20) { kpi[0] = 4.8; kpi[1] = 1.7; kpi[2] = 1.7; kpi[3] = 0.5; }
                if (Kpin >= 21 && Kpin <= 31) { kpi[0] = 5.9; kpi[1] = 2.0; kpi[2] = 1.9; kpi[3] = 0.6; }
                if (Kpin >= 32 && Kpin <= 40) { kpi[0] = 7.0; kpi[1] = 2.4; kpi[2] = 2.1; kpi[3] = 0.7; }
                if (Kpin >= 41 && Kpin <= 50) { kpi[0] = 8.3; kpi[1] = 2.7; kpi[2] = 2.3; kpi[3] = 0.8; }
                if (Kpin >= 51 && Kpin <= 60) { kpi[0] = 10.0; kpi[1] = 2.9; kpi[2] = 2.4; kpi[3] = 0.8; }
                if (Kpin >= 61 && Kpin <= 70) { kpi[0] = 11.6; kpi[1] = 3.5; kpi[2] = 2.6; kpi[3] = 0.9; }
                if (Kpin >= 71 && Kpin <= 80) { kpi[0] = 13.2; kpi[1] = 3.8; kpi[2] = 3.0; kpi[3] = 1.0; }
                if (Kpin >= 81 && Kpin <= 90) { kpi[0] = 15.1; kpi[1] = 4.2; kpi[2] = 3.3; kpi[3] = 1.1; }
                if (Kpin >= 91 && Kpin <= 110) { kpi[0] = 17.5; kpi[1] = 4.8; kpi[2] = 3.5; kpi[3] = 1.2; }
                if (Kpin >= 111 && Kpin <= 130) { kpi[0] = 19.9; kpi[1] = 5.3; kpi[2] = 3.8; kpi[3] = 1.3; }
                if (Kpin >= 131 && Kpin <= 150) { kpi[0] = 22.4; kpi[1] = 5.9; kpi[2] = 4.0; kpi[3] = 1.3; }
                if (Kpin >= 151 && Kpin <= 170) { kpi[0] = 25.2; kpi[1] = 6.4; kpi[2] = 4.1; kpi[3] = 1.4; }
                if (Kpin >= 171 && Kpin <= 190) { kpi[0] = 27.4; kpi[1] = 7.0; kpi[2] = 4.4; kpi[3] = 1.6; }
                if (Kpin >= 191 && Kpin <= 220) { kpi[0] = 29.8; kpi[1] = 7.6; kpi[2] = 4.5; kpi[3] = 1.6; }
                if (Kpin >= 221 && Kpin <= 250) { kpi[0] = 32.3; kpi[1] = 8.1; kpi[2] = 4.7; kpi[3] = 1.7; }
                if (Kpin >= 251 && Kpin <= 275) { kpi[0] = 35; kpi[1] = 8.7; kpi[2] = 5.0; kpi[3] = 1.8; }
                if (Kpin >= 276 && Kpin <= 300) { kpi[0] = 37.2; kpi[1] = 9.4; kpi[2] = 5.2; kpi[3] = 1.9; }
                if (Kpin >= 301 && Kpin <= 325) { kpi[0] = 39.6; kpi[1] = 10.8; kpi[2] = 5.5; kpi[3] = 2.1; }
                if (Kpin >= 326 && Kpin <= 350) { kpi[0] = 42.5; kpi[1] = 12.3; kpi[2] = 5.6; kpi[3] = 2.1; }
            }
            else { kpi[0] = 42.5+((Kpin-350)/50)*10; kpi[1] = 12.3 + ((Kpin - 350) / 50) * 3; kpi[2] = 5.6+ ((Kpin - 350) / 50) * 3; kpi[3] = 3; }
            return kpi;
        }
    }
}
