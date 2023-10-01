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
using СourseWork;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

namespace СourseWork
{
    internal class DataHandler
    {
//        //public static SqlConnection sqlConnection = null;
//        private SqlConnection sqlConnection;
//        public DataHandler(SqlConnection sqlConnection)
//        {
//            this.sqlConnection = sqlConnection;
//        }


//        public void FormOne()
//        {
//            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Database1"].ConnectionString);
//            sqlConnection.Open();
//            SqlCommand cmd = new SqlCommand
//("SELECT Id_Materials, Concat(Name_Material,Price) as Materials_detail FROM Materials",
//            sqlConnection);
//            SqlDataAdapter da = new SqlDataAdapter(cmd);
//            DataSet ds = new DataSet();
//            da.Fill(ds);
//            cmd.ExecuteNonQuery();

//            // Access the DataGridView control on Form1
//            Form1 form1 = (Form1)Application.OpenForms["Form1"];
//            form1.dataGridView1.DataSource = ds.Tables[0];
//        }
//        private void upd_table(string zapros)
//        {
//            SqlDataAdapter dataAdapter = new SqlDataAdapter(
//              zapros, sqlConnection);
//            DataSet dataSet = new DataSet();
//            dataAdapter.Fill(dataSet);

//            // Access the DataGridView control on Form1
//            Form1 form1 = (Form1)Application.OpenForms["Form1"];
//            form1.dataGridView1.DataSource = dataSet.Tables[0];         
//        }
//        private void zapros(SqlCommand command)
//        {
//            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
//            System.Data.DataTable dtable = new System.Data.DataTable();
//            dataAdapter.Fill(dtable);

//            // Access the DataGridView control on Form1
//            Form1 form1 = (Form1)Application.OpenForms["Form1"];
//            form1.dataGridView1.DataSource = dtable;
//        }


    }
}