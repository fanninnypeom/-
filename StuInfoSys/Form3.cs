using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StuInfoSys
{
    public partial class Form3 : Form
    {
        private int found;
        private Form1 f;
        public Form3(Form1 F)
        {
            f = F;
            InitializeComponent();
            Show();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            
        }
        private void Show() {
            if (f.radioButton1.Checked == true)
            {
                show("Name");
            }
            if (f.radioButton2.Checked == true)
            {
                show("Email");
            }
            if (f.radioButton3.Checked == true)
            {
                show("ID");
            }
        }
        private void show(string str) {
            OleDbConnection conn = new OleDbConnection();
            String s = System.Environment.CurrentDirectory;
            string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            connStr += @s + @"//StudentManager.mdb";
            // Console.WriteLine("当前连接字符串为:\n" + connStr + "\n");
            conn.ConnectionString = connStr;
            string strSelectQuery = "select * from students";
            OleDbCommand cmd = new OleDbCommand(strSelectQuery, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            da.SelectCommand = cmd;
            DataSet result = new DataSet();
            da.Fill(result, "students");
            int i;
            for (i = 0; i < result.Tables["students"].Rows.Count; i++)
            {
                if ((string)result.Tables[0].Rows[i][str] == f.textBox1.Text.ToString())
                {
                    this.textBox1.Text = (string) result.Tables[0].Rows[i]["Name"];
                    this.textBox2.Text = (string)result.Tables[0].Rows[i]["ID"];
                    this.textBox3.Text = (string)result.Tables[0].Rows[i]["Sex"];
                    this.textBox4.Text = (string)result.Tables[0].Rows[i]["Class"];
                    this.textBox5.Text = (string)result.Tables[0].Rows[i]["Email"];
                    this.textBox6.Text = (string)result.Tables[0].Rows[i]["Chinese"];
                    this.textBox7.Text = (string)result.Tables[0].Rows[i]["Math"];
                    this.textBox8.Text = (string)result.Tables[0].Rows[i]["English"];
                    this.Visible = true;
                    found = i;
                }
        }
            if (this.Visible ==false) {
                MessageBox.Show("找不到对象！");
            }

            conn.Close();
        }

        private void button1_Click(object sender, EventArgs e)//delete
        {
            operate op = new operate();
            op.delete(this,found);
            Close();
        }

        private void button2_Click(object sender, EventArgs e)//modify
        {
            operate op = new operate();
            op.modify(this,found);
            Close();
        }
    }
}
