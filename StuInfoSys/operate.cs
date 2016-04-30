using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace StuInfoSys
{
    class operate
    {
        public void modify(Form3 f,int found)
        {
            OleDbConnection conn = new OleDbConnection();
            String s = System.Environment.CurrentDirectory;
            string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            connStr += @s + @"//StudentManager.mdb";
            // Console.WriteLine("当前连接字符串为:\n" + connStr + "\n");
            conn.ConnectionString = connStr;
            string strSelectQuery = "select * from students";
            OleDbCommand cmd = new OleDbCommand(strSelectQuery, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            OleDbCommandBuilder cb = new OleDbCommandBuilder(da);
            da.SelectCommand = cmd;
            da.UpdateCommand = cb.GetUpdateCommand();
         
            
            DataSet result = new DataSet();
            da.Fill(result, "students");
            result.Tables[0].Rows[found]["Name"] = f.textBox1.Text;
            result.Tables[0].Rows[found]["ID"] = f.textBox2.Text;
            result.Tables[0].Rows[found]["Sex"] = f.textBox3.Text;
            result.Tables[0].Rows[found]["Class"] = f.textBox4.Text;
            result.Tables[0].Rows[found]["Email"] = f.textBox5.Text;
            result.Tables[0].Rows[found]["Chinese"]=f.textBox6.Text;
            result.Tables[0].Rows[found]["Math"] = f.textBox7.Text;
            result.Tables[0].Rows[found]["English"] = f.textBox8.Text;
            da.Update(result, "students");
            conn.Close();
        }
        public void delete(Form3 f,int found)
        {
            OleDbConnection conn = new OleDbConnection();
            String s = System.Environment.CurrentDirectory;
            string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            connStr += @s + @"//StudentManager.mdb";
            // Console.WriteLine("当前连接字符串为:\n" + connStr + "\n");
            conn.ConnectionString = connStr;
            string strSelectQuery = "select * from students";
            OleDbCommand cmd = new OleDbCommand(strSelectQuery, conn);
            OleDbDataAdapter da = new OleDbDataAdapter();
            OleDbCommandBuilder cb = new OleDbCommandBuilder(da);
            da.SelectCommand = cmd;
            da.UpdateCommand = cb.GetUpdateCommand();
      
            DataSet result = new DataSet();
            da.Fill(result, "students");
            result.Tables[0].Rows[found].Delete();
            da.Update(result, "students");
            conn.Close();
        }
        public void add(Form2 f) {
            String name=f.textBox1.Text.ToString();
            String ID = f.textBox2.Text.ToString();
            String sex = f.textBox3.Text.ToString();
            String Class = f.textBox4.Text.ToString();
            String Email = f.textBox5.Text.ToString();
            String chinese = f.textBox6.Text.ToString();
            String math = f.textBox7.Text.ToString();
            String english = f.textBox8.Text.ToString();
            OleDbConnection conn = new OleDbConnection();
            String str = System.Environment.CurrentDirectory;
            string connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            connStr +=@str+@"//StudentManager.mdb";
           // Console.WriteLine("当前连接字符串为:\n" + connStr + "\n");
            conn.ConnectionString = connStr;
            OleDbDataAdapter da = new OleDbDataAdapter();
            OleDbCommandBuilder cb = new OleDbCommandBuilder(da);
            string strSelectQuery = "select * from students";
            OleDbCommand cmd = new OleDbCommand(strSelectQuery, conn);
            da.SelectCommand = cmd;
            da.InsertCommand = cb.GetInsertCommand();
            DataSet result = new DataSet();
            da.Fill(result,"students");
            DataRow row1 = result.Tables["students"].NewRow();
            row1["Name"] = name;
            row1["ID"] = ID;
            row1["Sex"] = sex;
            row1["Class"] = Class;
            row1["Email"] = Email;
            row1["Chinese"] = chinese;
            row1["Math"] = math;
            row1["English"] = english;
            try {

                result.Tables["students"].Rows.Add(row1);

                da.Update(result, "students");
            }
            catch (System.Data.OleDb.OleDbException e) {
                MessageBox.Show(e.Message);
            }
                conn.Close();

            } 
    }
}
