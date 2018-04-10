using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
using System.Data.SqlClient;

namespace MyNameSpace
{
    public partial class PanelMaterialList : Form
    {
        public List<string[]> listPlateParts { get; set; }
        public List<string[]> listStiffParts { get; set; }
        public string Block { get; set; }
        public string HullNumber { get; set; }
        public PanelMaterialList()
        {
            InitializeComponent();
        }

        private void PanelMaterialList_Load(object sender, EventArgs e)
        {
            ///板材
            ///
            this.groupBox1.Text += this.Block;
            this.dataGridView2.Columns.Add("Name", "Name");
            this.dataGridView2.Columns.Add("Grade", "Grade");
            this.dataGridView2.Columns.Add("Thk.", "Thk.");
            this.dataGridView2.Columns.Add("Length", "Length");
            this.dataGridView2.Columns.Add("Width", "Width");
            this.dataGridView2.Columns.Add("Area", "Area");
            this.dataGridView2.Columns.Add("Weight", "Weight");
            this.dataGridView2.Columns.Add("Curved", "Curved");
            this.dataGridView2.Columns.Add("Material Status", "Material Status");
            if (this.listPlateParts.Count>0)
            {
                for (int i = 0; i <= listPlateParts.Count - 1; i++)
                {
                    int index = this.dataGridView2.Rows.Add();
                    this.dataGridView2.Rows[index].Cells[0].Value = listPlateParts[i][0];
                    this.dataGridView2.Rows[index].Cells[1].Value = listPlateParts[i][1];
                    this.dataGridView2.Rows[index].Cells[2].Value = listPlateParts[i][2];
                    this.dataGridView2.Rows[index].Cells[3].Value = listPlateParts[i][3];
                    this.dataGridView2.Rows[index].Cells[4].Value = listPlateParts[i][4];
                    this.dataGridView2.Rows[index].Cells[5].Value = listPlateParts[i][5];
                    this.dataGridView2.Rows[index].Cells[6].Value = listPlateParts[i][6];
                    this.dataGridView2.Rows[index].Cells[7].Value = listPlateParts[i][7];
                }
            }
            this.dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView2.AdvancedCellBorderStyle.All = DataGridViewAdvancedCellBorderStyle.Single;
            ///型材
            this.dataGridView1.Columns.Add("Name", "Name");
            this.dataGridView1.Columns.Add("Grade", "Grade");
            this.dataGridView1.Columns.Add("Size", "Size");
            this.dataGridView1.Columns.Add("Length", "Length");
            this.dataGridView1.Columns.Add("Weight", "Weight");
            this.dataGridView1.Columns.Add("Curved", "Curved");
            this.dataGridView1.Columns.Add("Material Status", "Material Status");
            if (this.listStiffParts.Count > 0)
            {
                for (int i = 0; i <= listStiffParts.Count - 1; i++)
                {
                    int index = this.dataGridView1.Rows.Add();
                    this.dataGridView1.Rows[index].Cells[0].Value = listStiffParts[i][0];
                    this.dataGridView1.Rows[index].Cells[1].Value = listStiffParts[i][1];
                    this.dataGridView1.Rows[index].Cells[2].Value = listStiffParts[i][2];
                    this.dataGridView1.Rows[index].Cells[3].Value = listStiffParts[i][3];
                    this.dataGridView1.Rows[index].Cells[4].Value = listStiffParts[i][4];
                    this.dataGridView1.Rows[index].Cells[5].Value = listStiffParts[i][5];
                }
            }
            this.dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.AdvancedCellBorderStyle.All = DataGridViewAdvancedCellBorderStyle.Single;
        }

        private void button1_Click(object sender, EventArgs e)//check material status
        {
            #region//将mto清单存放到list中
            List<tb_MTO> listMto = new List<tb_MTO>();
            string sqlcommand = $"select * from tb_MTO where Project='{this.HullNumber}'and BLK='{this.Block}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    listMto.Add(new tb_MTO() { Project = sdr["Project"].ToString(),
                    BLK=sdr["BLK"].ToString(),
                    Grade=sdr["Grade"].ToString(),
                    ItemSize=sdr["ItemSize"].ToString()
                    });
                }
            }
            else
            {
                MessageBox.Show("Please check whether the Mto is import to SPM dataBase ?");
                this.Close();
            }
            #endregion//
            #region//检查板材的材料
            for (int i = 0; i <= this.dataGridView2.Rows.Count-1; i++)
            {
                var gr=this.dataGridView1.Rows[i].Cells[2].Value.ToString();
                var thk = this.dataGridView1.Rows[i].Cells[3].Value.ToString();
                //20180224
            }
            #endregion

        }
    }
}
