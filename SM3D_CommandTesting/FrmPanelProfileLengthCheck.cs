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
//

namespace MyNameSpace
{
    public partial class FrmPanelProfileLengthCheck : Form
    {
        public List<StiffenerPart> stiff;
        public string Panel { get; set; }
        public FrmPanelProfileLengthCheck()
        {
            InitializeComponent();
        }

        private void FrmPanelProfileLengthCheck_Load(object sender, EventArgs e)
        {
            //public List<StiffenerPart> stiff;
            int i = 0;
            this.dataGridView1.Columns.Add("Part Name", "Part Name");
            this.dataGridView1.Columns.Add("Grade", "Grade");
            this.dataGridView1.Columns.Add("Size", "Size");
            this.dataGridView1.Columns.Add("Part Status", "Part Status");
            this.dataGridView1.Columns.Add("Detailing LTH", "Detailing LTH");
            this.dataGridView1.Columns.Add("Margin1", "Margin1");
            this.dataGridView1.Columns.Add("Margin2", "Margin2");
            this.dataGridView1.Columns.Add("Margin3", "Margin3");
            this.dataGridView1.Columns.Add("Margin4", "Margin4");
            this.dataGridView1.Columns.Add("MFG LTH", "MFG LTH");
            //CurvatureType
            this.dataGridView1.Columns.Add("CurvatureType", "CurvatureType");
            this.dataGridView1.Columns.Add("In TDL", "In TDL");
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            foreach (var item in this.stiff)
            {
                this.dataGridView1.Rows.Add(1);
                this.dataGridView1.Rows[i].Cells[0].Value = item.ToString();
                i++;
            }
            this.textBox1.Text = $"Will Check {this.Panel} all Profile Status, total {stiff.Count} profiles need be check";

        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string[]> str = new List<string[]>();
            foreach (var item in stiff.ToList())
            {
                string[] temp = new string[12];
                temp[0] = item.ToString();
                temp[1] = item.MaterialGrade;
                temp[2] = item.SectionName;
                temp[4] = item.MyGetProperty("ProfileLength");
                //
                var tdl=item.ToDoRecords.Where(c => c.ToDoItems.Count > 0);
                //
                //(item as BusinessObject)
                var margin = item.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Constant Margin");
                var mfg = item.AssemblyChildren.Where(c => (c as BusinessObject).ClassInfo.DisplayName == "Manufacturing Profile Part");
                #region//MFG information
                if (mfg.ToList().Count>0)
                {
                    var profile = mfg.ToList()[0] as ManufacturingProfile;
                    if (profile.IsUpToDate)
                    {
                        temp[3] = "Up to Data";
                    }
                    else temp[3] = "Out Of Data";
                    temp[9] =string.Format("{0:0.00}",profile.MyGetProperty("AfterFeaturesTotal"));
                    temp[10] = profile.MyGetProperty("CurvatureType");
                    if (tdl.Count()>0)
                    {
                        temp[11] = "Detailing in TDL";
                    }
                    else
                    {
                        tdl = profile.ToDoRecords.Where(c => c.ToDoItems.Count > 0);
                        if (tdl.Count() > 0)
                        {
                            temp[11] = "MFG in TDL";
                        }
                        else
                        {
                            temp[11] = "Not in TDL";
                        }
                    }
                }
                else
                {
                    temp[3] = "See Master";
                    temp[9] = "See Master";
                    temp[10] = "See Master";
                    if (tdl.Count() > 0)
                    {
                        temp[11] = "Detailing in TDL";
                    }
                    else
                    {
                        temp[11] = "Not in TDL";
                    }
                }
#endregion
                #region  //Margin information
                if (margin.ToList().Count > 0)
                {
                    List<string> str1 = new List<string>();
                    foreach (var item2 in margin.ToList())
                    {
                        str1.Add((item2 as BusinessObject).MyGetProperty("Type") + ":" + (item2 as BusinessObject).MyGetProperty("Value"));
                    }
                    switch (str1.Count)
                    {
                        case 1:
                            temp[5] = str1[0];
                            temp[6] = "N.A.";
                            temp[7] = "N.A.";
                            temp[8] = "N.A.";
                            break;
                        case 2:
                            temp[5] = str1[0];
                            temp[6] = str1[1];
                            temp[7] = "N.A.";
                            temp[8] = "N.A.";
                            break;

                        case 3:
                            temp[5] = str1[0];
                            temp[6] = str1[1];
                            temp[7] = str1[2];
                            temp[8] = "N.A.";
                            break;

                        case 4:
                            temp[5] = str1[0];
                            temp[6] = str1[1];
                            temp[7] = str1[2];
                            temp[8] = str1[3];
                            break;
                    }
                }
                else
                {
                    temp[5] = "N.A.";
                    temp[6] = "N.A.";
                    temp[7] = "N.A.";
                    temp[8] = "N.A.";
                }
                #endregion
                str.Add(temp);
            }
            int i = 0;
            this.dataGridView1.Rows.Clear();
            foreach (var item in str)
            {
                this.dataGridView1.Rows.Add(1);
                this.dataGridView1.Rows[i].Cells[0].Value = item[0];
                this.dataGridView1.Rows[i].Cells[1].Value = item[1];
                this.dataGridView1.Rows[i].Cells[2].Value = item[2];
                this.dataGridView1.Rows[i].Cells[3].Value = item[3];
                this.dataGridView1.Rows[i].Cells[4].Value = item[4];
                this.dataGridView1.Rows[i].Cells[5].Value = item[5];
                this.dataGridView1.Rows[i].Cells[6].Value = item[6];
                this.dataGridView1.Rows[i].Cells[7].Value = item[7];
                this.dataGridView1.Rows[i].Cells[8].Value = item[8];
                this.dataGridView1.Rows[i].Cells[9].Value = item[9];
                this.dataGridView1.Rows[i].Cells[10].Value = item[10];
                this.dataGridView1.Rows[i].Cells[11].Value = item[11];
                if (item[11]!= "Not in TDL")
                {
                    this.dataGridView1.Rows[i].Cells[11].Style.BackColor = Color.Red;
                }else this.dataGridView1.Rows[i].Cells[11].Style.BackColor = Color.LightGreen;

                if (item[3] != "Up to Data")
                {
                    if (item[3] != "See Master")
                    {
                        this.dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.Red;
                    }else this.dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.Gray;
                }
                else this.dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.LightGreen;

                if (item[9] != "See Master")
                {
                    this.dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.LightGreen;
                }
                else this.dataGridView1.Rows[i].Cells[9].Style.BackColor = Color.Gray;
                i++;
            }
            this.dataGridView1.Sort(this.dataGridView1.Columns[0], ListSortDirection.Ascending);
        }
    }
}
