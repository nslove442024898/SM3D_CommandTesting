using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyNameSpace
{
    public partial class FormForPartsQtyImport : Form
    {
        private Dictionary<string, int> _myPartsQty;
        public string ProjectName { get; set; }
        public string BlockName { get; set; }
        public FormForPartsQtyImport()
        {
            InitializeComponent();
        }
        List<tb_Schedule> tsSet = new List<tb_Schedule>();
        public Dictionary<string, int> MyPartsQty
        {
            get
            {
                return this._myPartsQty;
            }

            set
            {
                this._myPartsQty = value;
            }
        }
        private void FormForPartsQtyImport_Load(object sender, EventArgs e)
        {
            ////读取项目名称，并且加入到列表框中；
            this.textBox1.Text = this.ProjectName;
            this.textBox2.Text = this.BlockName;
            this.label3.Text += _myPartsQty.Count;
            this.dataGridView1.Columns.Add("Panel Name", "Panel Name");
            this.dataGridView1.Columns.Add("Model Qty", "Model Qty");
            this.dataGridView1.Columns.Add("SPM Qty", "SPM Qty");
            this.dataGridView1.Columns[0].Width = 100;
            this.dataGridView1.Columns[1].Width = 70;
            this.dataGridView1.Columns[2].Width = 70;
            foreach (KeyValuePair<string, int> item in this._myPartsQty)
            {
                int index = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[index].Cells[0].Value = item.Key.ToString();
                this.dataGridView1.Rows[index].Cells[1].Value = item.Value.ToString();
                string panelName = item.Key.Substring(item.Key.IndexOf("/") + 1);
                string sqlcommand = $"select * from tb_Schedule where Project='{this.ProjectName}'and BLOCK='{this.BlockName}'and Panel='{panelName}'";
                SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
                string spmPanelPartsQty = ""; double spmPanelkPI = 0.0; double actualHours = 0.0;
                List<string> shopDrawingTime = new List<string>();
                if (sdr.HasRows)//判断行不为空
                {
                    while (sdr.Read())//循环读取数据，知道无数据为止
                    {
                        spmPanelPartsQty = sdr["KPIN"].ToString();//提取tb_Schedule表的KPIN字段
                        shopDrawingTime.AddRange(new string[]{ (sdr["PSDM"].ToString()),
                        (sdr["PSCM"].ToString()),
                       ( sdr["PKDM"].ToString()),
                        (sdr["PKCM"].ToString())
                        });
                        if ((shopDrawingTime[0] == "0" && shopDrawingTime[1] == "0" && shopDrawingTime[2] == "" && shopDrawingTime[3] == "") || (shopDrawingTime[0] == "" && shopDrawingTime[1] == "" && shopDrawingTime[2] == "" && shopDrawingTime[3] == ""))
                        {
                            spmPanelkPI = 0.0;
                        }
                        else
                        {
                            spmPanelkPI = Convert.ToDouble(shopDrawingTime[0]) + Convert.ToDouble(shopDrawingTime[1]) + Convert.ToDouble(shopDrawingTime[2]) + Convert.ToDouble(shopDrawingTime[3]);
                        }

                    }
                }
                else
                {
                    MessageBox.Show($"please import panel List for {panelName} to SPM first then run this program again.");
                    this.Close();
                }
                this.dataGridView1.Rows[index].Cells[2].Value = spmPanelPartsQty;
                if (spmPanelPartsQty != "")
                {
                    actualHours = CommonTools.MyGetShopDrawingKpi(Convert.ToInt32(spmPanelPartsQty))[0] +
                                        CommonTools.MyGetShopDrawingKpi(Convert.ToInt32(spmPanelPartsQty))[1] +
                                        CommonTools.MyGetShopDrawingKpi(Convert.ToInt32(spmPanelPartsQty))[2] +
                                        CommonTools.MyGetShopDrawingKpi(Convert.ToInt32(spmPanelPartsQty))[3];
                }

                if (dataGridView1.Rows[index].Cells[1].Value.ToString() != spmPanelPartsQty || actualHours != spmPanelkPI)
                {
                    if (dataGridView1.Rows[index].Cells[1].Value.ToString() != spmPanelPartsQty)
                    {
                        this.dataGridView1.Rows[index].Cells[2].Style.BackColor = Color.Red;
                    }
                    else
                    {
                        this.dataGridView1.Rows[index].Cells[2].Style.BackColor = Color.LightYellow;
                    }
                }
                else
                {
                    this.dataGridView1.Rows[index].Cells[2].Style.BackColor = Color.Green;
                }

            }
        }
        List<int> incorrtIndex = new List<int>();//有问题的行号
        private void button1_Click(object sender, EventArgs e)
        {
            GetincorrectRowsIndex();
            int qty = 0;
            for (int i = 0; i < incorrtIndex.Count; i++)
            {
                int index = incorrtIndex[i];
                double[] strTemp = CommonTools.MyGetShopDrawingKpi(Convert.ToInt32(this.dataGridView1.Rows[index].Cells[1].Value.ToString()));

                string panelName = this.dataGridView1.Rows[index].Cells[0].Value.ToString().Substring(this.dataGridView1.Rows[index].Cells[0].Value.ToString().IndexOf("/") + 1);
                string sqlcommand = $"Update tb_Schedule set [KPIN]='{this.dataGridView1.Rows[index].Cells[1].Value.ToString()}'" +
                    $",[PSDM]={strTemp[0]}" + $",[PSCM]={strTemp[1]}" + $",[PKDM]={strTemp[2]}" + $",[PKCM]={strTemp[3]}" +
                     $" where Project='{this.ProjectName}'and BLOCK='{this.BlockName}'and Panel='{panelName}'";
                qty += SqlHelper.MyExecuteNonQuery(sqlcommand);
            }
            MessageBox.Show("total " + qty + " panels Parts Qty have been Import to Spm!");
        }
        private void GetincorrectRowsIndex()
        {
            incorrtIndex.Clear();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                int qty = 0;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    if (dataGridView1.Rows[i].Cells[j].Style.BackColor == Color.Red || dataGridView1.Rows[i].Cells[j].Style.BackColor == Color.LightYellow)
                    {
                        qty++;
                    }
                }
                if (qty != 0)
                {
                    incorrtIndex.Add(i);
                }
            }
            MessageBox.Show("total " + incorrtIndex.Count.ToString() + " panels Parts Qty will Import to Spm!");
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
