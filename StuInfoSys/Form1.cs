using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using ADOX;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StuInfoSys
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
            string str = System.Environment.CurrentDirectory;
            string dbName1=str + "\\StudentManager.mdb";//注意扩展名必须为mdb,否则不能插入表
            string dbName = @dbName1;
            ADOX.CatalogClass cat = new ADOX.CatalogClass();
            try {
                cat.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbName );
            }
            catch (System.Runtime.InteropServices.COMException) {

            }
            ADODB.Connection cn = new ADODB.Connection();
            cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbName , null, null, -1);
            cat.ActiveConnection = cn;

            //新建表  
            ADOX.Table table = new ADOX.Table();
            table.Name = "students";

            ADOX.Column column = new ADOX.Column();
            column.ParentCatalog = cat;
            column.Type = ADOX.DataTypeEnum.adVarWChar; // 必须先设置字段类型  
            column.Name = "ID";
            column.DefinedSize = 50;
            column.Properties["AutoIncrement"].Value = true;
            table.Columns.Append(column, DataTypeEnum.adVarWChar, 50);
            //设置主键  
          
    table.Keys.Append("PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "ID", "", "");
           
            table.Columns.Append("Name", DataTypeEnum.adVarWChar, 50);
            table.Columns.Append("Sex", DataTypeEnum.adVarWChar, 50);
            table.Columns.Append("Class", DataTypeEnum.adVarWChar, 50);
            table.Columns.Append("Email", DataTypeEnum.adVarWChar, 50);
           table.Columns.Append("Chinese", DataTypeEnum.adVarWChar, 50);
            table.Columns.Append("Math", DataTypeEnum.adVarWChar, 50);
            table.Columns.Append("English", DataTypeEnum.adVarWChar, 50);
            try
            {
                cat.Tables.Append(table);
            }
            catch (Exception ex)
            {
                Console.WriteLine("111");
            }
            //此处一定要关闭连接，否则添加数据时候会出错  

            table = null;
            cat = null;
            Application.DoEvents();
            cn.Close();
        }   

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.Visible = true;
//            this.Activate();
//            this.button1.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true || radioButton2.Checked == true || radioButton3.Checked == true)
            {
                Form3 f = new Form3(this);
              
            }
        
        }


    }
}
