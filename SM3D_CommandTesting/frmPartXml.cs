using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//
using Ingr.SP3D.Common.Client;
using Ingr.SP3D.Common.Client.Services;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
//
using Ingr.SP3D.Manufacturing.Middle;
using Ingr.SP3D.Manufacturing.Middle.Services;
//
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Structure.Middle.Services;

namespace MyNameSpace
{
    public partial class frmPartXml : Form
    {
        private SelectSet _mSlectObj;
        public string BlockName { get; set; }
        public List<ManufacturingPlate> mfgPlates = new List<ManufacturingPlate>();
        public List<ManufacturingProfile> mfgPros = new List<ManufacturingProfile>();
        public frmPartXml()
        {
            InitializeComponent();
            this.button1.Enabled = false;
            this.textBox1.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\";
            _mSlectObj = ClientServiceProvider.SelectSet;
            if (_mSlectObj.SelectedObjects.Count > 0)
            {
                foreach (var item in _mSlectObj.SelectedObjects)
                {
                    if (item.ClassInfo.DisplayName == "Assembly")
                    {
                        this.comboBox1.Items.Add(item.ToString());
                    }
                }
            }
            BlockName = (_mSlectObj.SelectedObjects[0] as IAssemblyChild).AssemblyParent.ToString();
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.checkedListBox1.Items.Clear();
            var panel = _mSlectObj.SelectedObjects.Where(C => C.ToString() == this.comboBox1.Text).ToList()[0];
            if (panel is IAssembly)
            {
                var assPanel = panel as IAssembly;
                var listMfgPlates = GetAssemblyPlateParts(assPanel);
                mfgPlates = listMfgPlates;
                var listMfgProfiles = GetAssemblyProfileParts(assPanel);
                mfgPros = listMfgProfiles;
                //
                for (int i = 0; i <= listMfgPlates.Count - 1; i++)
                {
                    this.checkedListBox1.Items.Add(listMfgPlates[i].ToString());
                }
                //for (int i = 0; i <= listMfgProfiles.Count - 1; i++)
                //{
                //    this.checkedListBox1.Items.Add(listMfgProfiles[i].ToString());
                //}
            }
            //
            this.button1.Enabled = true;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ManufacturingOutputBase mfgPart = null;
            for (int i = 0; i <= this.checkedListBox1.CheckedItems.Count - 1; i++)
            {
                string str = this.checkedListBox1.CheckedItems[i].ToString();
                //if (this.checkedListBox1.Items.IndexOf(str) > mfgPlates.Count - 1)//是型材
                //{
                //    mfgPart = mfgPros[this.checkedListBox1.Items.IndexOf(str) - mfgPlates.Count];
                //}
                //else//是板材
                //{
                    mfgPart = mfgPlates[this.checkedListBox1.Items.IndexOf(str)];
                //}
                if (!mfgPart.IsUpToDate)
                {
                    mfgPart.Update();
                    ClientServiceProvider.TransactionMgr.Commit("FilterSieve");
                }
                    mfgPart.OutputAsString("Plates & Profiles with Annotation", this.textBox1.Text + mfgPart.Name.Replace('/', '-') + ".xml");
            }
            MessageBox.Show($"Save selected parts xml file to { this.textBox1.Text} already, Please review ! ");
        }
        private List<ManufacturingPlate> GetAssemblyPlateParts(IAssembly assFolder)
        {
            List<ManufacturingPlate> resultList = new List<ManufacturingPlate>();
            var plateParts = assFolder.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "PlatePart");//单个的parts
            var platePartsTemp = assFolder.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly");//组合的肘板或者t-girder等
            //单个的parts
            if (plateParts.ToList().Count > 0 && platePartsTemp.ToList().Count == 0)
            {
                foreach (var item in plateParts.ToList())
                {
                    foreach (var item1 in (item as IAssembly).AssemblyChildren)
                    {
                        if ((item1 as BusinessObject).ClassInfo.DisplayName == "Manufacturing Plate Part")
                        {
                            ManufacturingPlate mfgPlate = item1 as ManufacturingPlate;
                            resultList.Add(mfgPlate);
                        }
                    }
                }
            }
            //组合的肘板或者t-girder等
            else if (plateParts.ToList().Count > 0 && platePartsTemp.ToList().Count > 0)
            {
                foreach (var item in platePartsTemp.ToList())
                {
                    resultList.AddRange(GetAssemblyPlateParts(item as IAssembly));
                }
                foreach (var item in plateParts.ToList())
                {
                    foreach (var item1 in (item as IAssembly).AssemblyChildren)
                    {
                        if ((item1 as BusinessObject).ClassInfo.DisplayName == "Manufacturing Plate Part")
                        {
                            ManufacturingPlate mfgPlate = item1 as ManufacturingPlate;
                            resultList.Add(mfgPlate);
                        }
                    }
                }
            }
            else
            {
                foreach (var item in platePartsTemp.ToList())
                {
                    resultList.AddRange(GetAssemblyPlateParts(item as IAssembly));
                }
            }
            return resultList;
        }

        private List<ManufacturingProfile> GetAssemblyProfileParts(IAssembly assFolder)//骨材
        {
            List<ManufacturingProfile> resultList = new List<ManufacturingProfile>();
            var stiffParts = assFolder.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName.EndsWith("ProfilePart") ||
            (c as BusinessObject).ClassInfo.DisplayName.EndsWith("StiffenerPart"));//单个的骨材
            var platePartsTemp = assFolder.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Assembly");//组合的肘板或者t-girder等
            //单个的parts
            if (stiffParts.ToList().Count > 0 && platePartsTemp.ToList().Count == 0)
            {
                foreach (var item in stiffParts.ToList())
                {
                    foreach (var item1 in (item as IAssembly).AssemblyChildren)
                    {
                        if ((item1 as BusinessObject).ClassInfo.DisplayName == "Manufacturing Profile Part")
                        {
                            ManufacturingProfile mfgStiff = item1 as ManufacturingProfile;
                            resultList.Add(mfgStiff);
                        }
                    }
                }
            }
            //组合的肘板或者t-girder等
            else if (stiffParts.ToList().Count > 0 && platePartsTemp.ToList().Count > 0)
            {
                foreach (var item in platePartsTemp.ToList())
                {
                    resultList.AddRange(GetAssemblyProfileParts(item as IAssembly));
                }
                foreach (var item in stiffParts.ToList())
                {
                    foreach (var item1 in (item as IAssembly).AssemblyChildren)
                    {
                        if ((item1 as BusinessObject).ClassInfo.DisplayName == "Manufacturing Profile Part")
                        {
                            ManufacturingProfile mfgStiff = item1 as ManufacturingProfile;
                            resultList.Add(mfgStiff);
                        }
                    }
                }
            }
            else
            {
                foreach (var item in platePartsTemp.ToList())
                {
                    resultList.AddRange(GetAssemblyProfileParts(item as IAssembly));
                }
            }
            return resultList;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = fbd.SelectedPath + @"\";
            }
        }

        private void frmPartXml_Deactivate(object sender, EventArgs e)
        {
            this.Activate();
        }

    }
}
