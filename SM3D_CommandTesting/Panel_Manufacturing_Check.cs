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
    public partial class Panel_Manufacturing_Check : Form
    {
        public string Hull { get; set; }
        public string BLK { get; set; }
        public Dictionary<string, int> panelMfg { get; set; }
        public Panel_Manufacturing_Check()
        {
            InitializeComponent();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.listBox1.Items.Clear();
            string sqlcommand = $"select * from tb_Schedule where Project='{this.Hull}'and BLOCK='{this.BLK}' and BATCH='{this.comboBox1.Text}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    if (this.panelMfg.ContainsKey(sdr["Panel"].ToString()))
                    {
                        this.listBox1.Items.Add(sdr["Panel"].ToString() + "----> Have " + panelMfg[sdr["Panel"].ToString()] + " Panel MFG");
                    }
                }
            }
        }

        private void Panel_Manufacturing_Check_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = Hull;
            this.textBox2.Text = BLK;
            string sqlcommand = $"select * from tb_Schedule where Project='{this.Hull}'and BLOCK='{this.BLK}'";
            SqlDataReader sdr = SqlHelper.MyExecuteReader(sqlcommand);//调用自己编写的sqlhelper类
            Dictionary<string, int> dicBatch = new Dictionary<string, int>();
            if (sdr.HasRows)//判断行不为空
            {
                while (sdr.Read())//循环读取数据，知道无数据为止
                {
                    if (dicBatch.ContainsKey(sdr["BATCH"].ToString()))
                    {
                        dicBatch[sdr["BATCH"].ToString()] = 1;
                    }
                    else
                    {
                        dicBatch[sdr["BATCH"].ToString()] = 0;
                    }
                }
            }
            foreach (var item in dicBatch.Keys)
            {
                this.comboBox1.Items.Add(item);
            }
        }

    }
}
