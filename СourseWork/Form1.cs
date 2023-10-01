using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace СourseWork
{
    public partial class Form1 : Form
    {
        public static SqlConnection sqlConnection = null;
        public string checkTable;
        public Form1()
        {
            InitializeComponent();
        }

        private void upd_table(string zapros)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
              zapros, sqlConnection);
            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);
            dataGridView1.DataSource = dataSet.Tables[0];
        }
        private void delete(SqlCommand command, String s, String s1)
        {
            bool ok = false;
            try
            { command.ExecuteNonQuery(); }
            catch (Exception)
            {
                MessageBox.Show(s);
                ok = true;
            }
            finally
            { if (!ok) MessageBox.Show(s1); }
        }
        private void zapros(SqlCommand command)
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
            System.Data.DataTable dtable = new System.Data.DataTable();
            dataAdapter.Fill(dtable);
            dataGridView1.DataSource = dtable;
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Database1"].ConnectionString);
            sqlConnection.Open();
            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS Руководитель\r\nFROM [dbo].[Faculty];\r\n");
        }


        //Другие элементы
        #region OtherElements

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }


        #endregion

        //Кнопки взаимодействия
        #region IteractionButtons
        private void button1_Click(object sender, EventArgs e)
        {
            Form2 plForm = new Form2();
            plForm.panel4.Visible = true;//main panel

            plForm.panel1.Visible = false;
            plForm.panel2.Visible = false;
            plForm.panel3.Visible = false;
            plForm.panel5.Visible = false;
            plForm.panel6.Visible = false;
            DialogResult result;

            string panelElements1;
            string panelElements2;
            string panelElements3;
            string panelElements4;
            int panelIntElements1;
            int panelIntElements2;
            int panelIntElements3;
            int k = 0;
            switch (checkTable)
            {
                case "Факультет":
                    plForm.panel1.Visible = true;
                    plForm.panel1.BringToFront();

                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox1.Text;
                    panelElements2 = plForm.textBox2.Text;
                    if (panelElements1 == "" || panelElements2 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    else
                    {
                        //проверка на повторение с уже имеющимся факультетом 
                        for (int i = 0; i <= dataGridView1.RowCount - 2; i++)
                        {
                            if (dataGridView1[1, i] != null && dataGridView1[1, i].Value.ToString() == panelElements1)
                            { k = 1; break; }
                        }
                        if (k == 1)
                        {
                            MessageBox.Show("Факультет с названием " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec INS_FACULTY @NAME, @HEAD"
                              , sqlConnection);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("HEAD", panelElements1);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Факультет " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS Руководитель\r\nFROM [dbo].[Faculty];\r\n");
                        }
                    }
                    break;

                case "Кафедра":
                    plForm.panel2.Visible = true;
                    plForm.panel2.BringToFront();
                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox3.Text;
                    panelElements2 = plForm.textBox4.Text;
                    panelIntElements1 = (int)plForm.numericUpDown1.Value;
                    if (panelElements1 == "" || panelElements2 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    else
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 2; i++)
                        {
                            if (dataGridView1[1, i].Value.ToString() == panelElements1)
                            { k = 1; break; }
                        }
                        if (k == 1)
                        {
                            MessageBox.Show("Кафедра с названием " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec INS_DEPARTMENT @NAME, @HEAD, @FACULTYCODE"
                              , sqlConnection);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("HEAD", panelElements2);
                            command.Parameters.AddWithValue("FACULTYCODE", panelIntElements1);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Кафедра " + panelElements1 + " добавлена в Бд.");
                            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS НачальникКафедры, [FacultyCode] AS КодФакультета\r\nFROM [dbo].[Department]\r\n");//
                        }
                    }
                    break;

                case "Преподаватель":
                    plForm.panel3.Visible = true;
                    plForm.panel3.BringToFront();
                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox5.Text;
                    panelElements2 = plForm.textBox6.Text;
                    panelElements3 = plForm.textBox7.Text;
                    panelIntElements1 = (int)plForm.numericUpDown6.Value;
                    panelIntElements2 = (int)plForm.numericUpDown7.Value;
                    if (panelElements1 == "" || panelElements2 == "" || panelElements3 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    else
                    {

                        if (k == 1)
                        {
                            MessageBox.Show(" с названием " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            try
                            {
                                SqlCommand command = new SqlCommand(
                              $"exec INS_TEACHER @NAME, @POSITION, @STATUS, @ACADEMICDEGREEID, @DEPARTMENTID"
                              , sqlConnection);
                                command.Parameters.AddWithValue("NAME", panelElements1);
                                command.Parameters.AddWithValue("POSITION", panelElements2);
                                command.Parameters.AddWithValue("STATUS", panelElements3);
                                command.Parameters.AddWithValue("ACADEMICDEGREEID", panelIntElements1);
                                command.Parameters.AddWithValue("DEPARTMENTID", panelIntElements2);
                                command.ExecuteNonQuery();/////////
                                MessageBox.Show("Преподаватель " + panelElements1 + " добавлен в Бд.");
                                upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [Status] AS Статус, \r\n[AcademicDegreeID] AS ИДУченойСтепени, [DepartmentID] AS ИДОтделения\r\nFROM [dbo].[Teacher]\r\n");//

                            }
                            catch { MessageBox.Show("Ошибка. Введенное значение Ученой степени не существует в Бд. " +
                                "Добавьте его в базу, прежде чем присваивать преподавателям."); }
                        }
                            
                    }
                    break;

                case "Почасовик":
                    plForm.panel5.Visible = true;
                    plForm.panel5.BringToFront();
                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox9.Text;
                    panelElements2 = plForm.textBox10.Text;
                    panelIntElements1 = (int)plForm.numericUpDown2.Value;
                    if (panelElements1 == "" || panelElements2 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    else
                    {
                        
                        //for (int i = 0; i <= dataGridView1.RowCount - 2; i++)
                        //{
                        //    if (dataGridView1[0, i].Value.ToString() == panelElements1)
                        //    { k = 1; break; }
                        //}
                        if (k == 1)
                        {
                            MessageBox.Show("Почасовой работник " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec INS_HOURLYWORKER @NAME, @POSITION, @RANKID"
                              , sqlConnection);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("POSITION", panelElements2);
                            command.Parameters.AddWithValue("RANKID", panelIntElements1);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Почасовой работник " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [RankID] AS ИДРанга\r\nFROM [dbo].[HourlyWorker]\r\n");//
                        }
                    }
                    break;

                case "Плановая нагрузка почасовика":
                    plForm.panel6.Visible = true;
                    plForm.panel6.BringToFront();
                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelIntElements1 = (int)plForm.numericUpDown3.Value;
                    panelElements1 = plForm.textBox11.Text;
                    panelIntElements2 = (int)plForm.numericUpDown4.Value;
                    panelIntElements3 = (int)plForm.numericUpDown5.Value;
                    if (panelElements1 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    else
                    {
                        
                        //for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
                        //{
                        //    if (dataGridView1[0, i].Value.ToString() == panelElements1)
                        //    { k = 1; break; }
                        //}
                        if (k == 1)
                        {
                            MessageBox.Show("План для почасового работника " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec INS_PLANNEDLOAD @HOURLYID, @POSITION, @YEAR, @MONTHLYLOAD"
                              , sqlConnection);
                            command.Parameters.AddWithValue("HOURLYID", panelIntElements1);
                            command.Parameters.AddWithValue("POSITION", panelElements1);
                            command.Parameters.AddWithValue("YEAR", panelIntElements2);
                            command.Parameters.AddWithValue("MONTHLYLOAD", panelIntElements3);
                            command.ExecuteNonQuery();
                            MessageBox.Show("План для почасового работника " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT\r\n    [HourlyID] AS [ИдентификаторРабочего],\r\n    HourlyWorker.FullName AS [ФИО],\r\n    YEAR([Date]) AS [Год],\r\n    MONTH([Date]) AS [Месяц],\r\n    SUM([HoursWorked]) AS [Кол-во_раб_часов]\r\nFROM\r\n    [dbo].[Workload], [HourlyWorker]\r\nGROUP BY\r\n    [HourlyID], HourlyWorker.FullName, YEAR([Date]), MONTH([Date])\r\n");//
                        }
                    }
                    break;

                case "Фактическая нагрузка для почасовика":
                    MessageBox.Show("no insert function");
                    break;

                case "Ранг почасовика":
                    MessageBox.Show("no insert function");
                    break;

                case "Ученая степень":
                    MessageBox.Show("no insert function");
                    break;

                case "Звание":
                    MessageBox.Show("no insert function");
                    break;

                case "Почасовик-Кафедра":
                    MessageBox.Show("no insert function");
                    break;

                default:
                    plForm.panel2.Visible = false;
                    MessageBox.Show("error num table");
                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form2 plForm = new Form2();
            plForm.panel4.Visible = true;//main panel

            plForm.panel1.Visible = false;
            plForm.panel2.Visible = false;
            plForm.panel3.Visible = false;
            plForm.panel5.Visible = false;
            plForm.panel6.Visible = false;
            DialogResult result;

            string panelElements1;
            string panelElements2;
            string panelElements3;
            string panelElements4;
            int panelIntElements1;
            int panelIntElements2;
            int panelIntElements3;
            int index;
            int ID;
            int k = 0;
            switch (checkTable)
            {
                case "Факультет":
                    try
                    {
                        index = dataGridView1.SelectedRows[0].Index;
                        plForm.textBox1.Text = dataGridView1[1, index].Value.ToString(); 
                        plForm.textBox2.Text = dataGridView1[2, index].Value.ToString();
                    }
                    catch { MessageBox.Show("Выберите строку для изменения."); break; }
                    plForm.panel1.Visible = true;
                    plForm.panel1.BringToFront();

                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox1.Text;
                    panelElements2 = plForm.textBox2.Text;
                    ID = Convert.ToInt32(dataGridView1[0, index].Value);
                    while (panelElements1 == "" || panelElements2 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    {
                        //проверка на повторение с уже имеющимся факультетом 
                        for (int i = 0; i <= dataGridView1.RowCount - 2; i++)
                        {
                            if (dataGridView1[1, i] != null && dataGridView1[1, i].Value.ToString() == panelElements1)
                            { k = 1; break; }
                        }
                        if (k == 1)
                        {
                            MessageBox.Show("Факультет с названием " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec UpdateFaculty @Code, @NAME, @HEAD"
                              , sqlConnection);
                            command.Parameters.AddWithValue("Code", ID);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("HEAD", panelElements2);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Факультет " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS Руководитель\r\nFROM [dbo].[Faculty];\r\n");
                        }
                    }
                    break;

                case "Кафедра":
                    try
                    {
                        index = dataGridView1.SelectedRows[0].Index;
                        plForm.textBox3.Text = dataGridView1[1, index].Value.ToString(); // Устанавливаем название кафедры в textBox3
                        plForm.textBox4.Text = dataGridView1[2, index].Value.ToString(); // Устанавливаем начальника кафедры в textBox4
                    }
                    catch { MessageBox.Show("Выберите строку для изменения."); break; }
                    plForm.panel2.Visible = true;
                    plForm.panel2.BringToFront();
                    result = plForm.ShowDialog(this);
                    if (result == DialogResult.Cancel)
                        return;
                    panelElements1 = plForm.textBox3.Text;
                    panelElements2 = plForm.textBox4.Text;
                    panelIntElements1 = (int)plForm.numericUpDown1.Value;
                    ID = Convert.ToInt32(dataGridView1[0, index].Value);
                    while (panelElements1 == "" || panelElements2 == "") { MessageBox.Show("Все поля должны быть заполнены"); plForm.ShowDialog(this); }
                    {
                        for (int i = 0; i <= dataGridView1.RowCount - 2; i++)
                        {
                            if (dataGridView1[1, i].Value.ToString() == panelElements1)
                            { k = 1; break; }
                        }
                        if (k == 1)
                        {
                            MessageBox.Show("Кафедра с названием " + panelElements1 + " уже существует!");
                        }
                        else
                        {
                            SqlCommand command = new SqlCommand(
                              $"exec UpdateDepartment @code, @NAME, @HEAD, @FACULTYCODE"
                              , sqlConnection);
                            command.Parameters.AddWithValue("code", ID);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("HEAD", panelElements2);
                            command.Parameters.AddWithValue("FACULTYCODE", panelIntElements1);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Кафедра " + panelElements1 + " добавлена в Бд.");
                            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS НачальникКафедры, [FacultyCode] AS КодФакультета\r\nFROM [dbo].[Department]\r\n");//
                        }
                    }
                    break;

                case "Преподаватель":
                    try
                    {
                        index = dataGridView1.SelectedRows[0].Index;
                        plForm.textBox5.Text = dataGridView1[1, index].Value.ToString(); // Устанавливаем полное имя преподавателя в textBox5
                        plForm.textBox6.Text = dataGridView1[2, index].Value.ToString(); // Устанавливаем должность преподавателя в textBox6
                        plForm.textBox7.Text = dataGridView1[3, index].Value.ToString(); // Устанавливаем статус преподавателя в textBox7
                        plForm.numericUpDown6.Value = (int)dataGridView1[4, index].Value; // Устанавливаем идентификатор ученой степени в numericUpDown6
                        plForm.numericUpDown7.Value = (int)dataGridView1[5, index].Value; // Устанавливаем идентификатор отделения в numericUpDown7
                    }
                    catch { MessageBox.Show("Выберите строку для изменения."); break; }
                    plForm.panel3.Visible = true;
                    plForm.panel3.BringToFront();
                    result = plForm.ShowDialog(this);

                    if (result == DialogResult.Cancel)
                        return;

                    panelElements1 = plForm.textBox5.Text;
                    panelElements2 = plForm.textBox6.Text;
                    panelElements3 = plForm.textBox7.Text;
                    panelIntElements1 = (int)plForm.numericUpDown6.Value;
                    panelIntElements2 = (int)plForm.numericUpDown7.Value;
                    ID = Convert.ToInt32(dataGridView1[0, index].Value);

                    if (string.IsNullOrWhiteSpace(panelElements1) || string.IsNullOrWhiteSpace(panelElements2) || string.IsNullOrWhiteSpace(panelElements3))
                    {
                        MessageBox.Show("Все поля должны быть заполнены");
                    }
                    else
                    {
                       

                        try
                        {
                            SqlCommand command = new SqlCommand(
                                "exec UpdateTeacher @teacher_id, @NAME, @POSITION, @STATUS, @ACADEMICDEGREEID, @DEPARTMENTID"
                                , sqlConnection);
                            command.Parameters.AddWithValue("teacher_id", ID);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("POSITION", panelElements2);
                            command.Parameters.AddWithValue("STATUS", panelElements3);
                            command.Parameters.AddWithValue("ACADEMICDEGREEID", panelIntElements1);
                            command.Parameters.AddWithValue("DEPARTMENTID", panelIntElements2);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Преподаватель " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [Status] AS Статус, \r\n[AcademicDegreeID] AS ИДУченойСтепени, [DepartmentID] AS ИДОтделения\r\nFROM [dbo].[Teacher]\r\n");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
                        }
                    }
                    break;


                case "Почасовик":
                    try
                    {
                        index = dataGridView1.SelectedRows[0].Index;
                        plForm.textBox9.Text = dataGridView1[1, index].Value.ToString(); // Устанавливаем полное имя почасового работника в textBox9
                        plForm.textBox10.Text = dataGridView1[2, index].Value.ToString(); // Устанавливаем должность почасового работника в textBox10
                        plForm.numericUpDown2.Value = (int)dataGridView1[3, index].Value; // Устанавливаем идентификатор ранга в numericUpDown2
                    }
                    catch { MessageBox.Show("Выберите строку для изменения."); break; }
                    plForm.panel5.Visible = true;
                    plForm.panel5.BringToFront();
                    result = plForm.ShowDialog(this);

                    if (result == DialogResult.Cancel)
                        return;

                    panelElements1 = plForm.textBox9.Text;
                    panelElements2 = plForm.textBox10.Text;
                    panelIntElements1 = (int)plForm.numericUpDown2.Value;
                    ID = Convert.ToInt32(dataGridView1[0, index].Value);

                    if (string.IsNullOrWhiteSpace(panelElements1) || string.IsNullOrWhiteSpace(panelElements2))
                    {
                        MessageBox.Show("Все поля должны быть заполнены");
                    }
                    else
                    {
                        

                        try
                        {
                            SqlCommand command = new SqlCommand(
                                "exec UpdateHourlyWorker @ID, @NAME, @POSITION, @RANKID"
                                , sqlConnection);
                            command.Parameters.AddWithValue("ID", ID);
                            command.Parameters.AddWithValue("NAME", panelElements1);
                            command.Parameters.AddWithValue("POSITION", panelElements2);
                            command.Parameters.AddWithValue("RANKID", panelIntElements1);
                            command.ExecuteNonQuery();
                            MessageBox.Show("Почасовой работник " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [RankID] AS ИДРанга\r\nFROM [dbo].[HourlyWorker]\r\n");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
                        }
                    }
                    break;


                case "Плановая нагрузка почасовика":
                    try
                    {
                        index = dataGridView1.SelectedRows[0].Index;
                        plForm.numericUpDown3.Value = (int)dataGridView1[0, index].Value; // Устанавливаем идентификатор почасового работника в numericUpDown3
                        plForm.textBox11.Text = dataGridView1[1, index].Value.ToString(); // Устанавливаем должность в textBox11
                        plForm.numericUpDown4.Value = (int)dataGridView1[2, index].Value; // Устанавливаем год в numericUpDown4
                    }
                    catch { MessageBox.Show("Выберите строку для изменения."); break; }
                    plForm.panel6.Visible = true;
                    plForm.panel6.BringToFront();
                    result = plForm.ShowDialog(this);

                    if (result == DialogResult.Cancel)
                        return;

                    panelIntElements1 = (int)plForm.numericUpDown3.Value;
                    panelElements1 = plForm.textBox11.Text;
                    panelIntElements2 = (int)plForm.numericUpDown4.Value;
                    panelIntElements3 = (int)plForm.numericUpDown5.Value;
                    ID = Convert.ToInt32(dataGridView1[0, index].Value);

                    if (string.IsNullOrWhiteSpace(panelElements1))
                    {
                        MessageBox.Show("Должность должна быть заполнена");
                    }
                    else
                    {

                        try
                        {
                            SqlCommand command = new SqlCommand(
                                "exec UpdatePlannedLoad @HourlyID, @Position, @YEAR, @MONTHLYLOAD"
                                , sqlConnection);
                            command.Parameters.AddWithValue("@HourlyID", panelIntElements1);
                            command.Parameters.AddWithValue("Position", panelElements1);
                            command.Parameters.AddWithValue("YEAR", panelIntElements2);
                            command.Parameters.AddWithValue("MONTHLYLOAD", panelIntElements3);
                            command.ExecuteNonQuery();
                            MessageBox.Show("План для почасового работника " + panelElements1 + " добавлен в Бд.");
                            upd_table("SELECT [HourlyID] AS ИДПочасовикаРабочего, [Position] AS Должность, [Year] AS Год, [MonthlyLoad] AS МесячнаяНагрузка\r\nFROM [dbo].[PlannedLoad]\r\n");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
                        }
                    }
                    break;


                case "Фактическая нагрузка для почасовика":
                    MessageBox.Show("no insert function");
                    break;

                case "Ранг почасовика":
                    MessageBox.Show("no insert function");
                    break;

                case "Ученая степень":
                    MessageBox.Show("no insert function");
                    break;

                case "Звание":
                    MessageBox.Show("no insert function");
                    break;

                case "Почасовик-Кафедра":
                    MessageBox.Show("no insert function");
                    break;

                default:
                    plForm.panel2.Visible = false;
                    MessageBox.Show("error num table");
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            switch (checkTable)
            {
                case "Факультет":
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        int index = dataGridView1.SelectedRows[0].Index;
                        int id_faculty = Convert.ToInt32(dataGridView1[0, index].Value);
                        SqlCommand command = new SqlCommand(
                          $"exec DEL_FACULTY @Id_faculty"
                          , sqlConnection);
                        command.Parameters.AddWithValue("Id_faculty", id_faculty);
                        delete(command, "Таблица содержит связанные данные!", "Объект удалён!");
                        upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS Руководитель\r\nFROM [dbo].[Faculty];\r\n");
                    }
                    break;

                case "Кафедра":
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        int index = dataGridView1.SelectedRows[0].Index;
                        int id_department = Convert.ToInt32(dataGridView1[0, index].Value);
                        SqlCommand command = new SqlCommand(
                          $"exec DEL_DEPARTMENT @Id_department"
                          , sqlConnection);
                        command.Parameters.AddWithValue("Id_department", id_department);
                        delete(command, "Таблица содержит связанные данные!", "Объект удалён!");
                        upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS НачальникКафедры, [FacultyCode] AS КодФакультета\r\nFROM [dbo].[Department]\r\n");
                    }
                    break;

                case "Преподаватель":
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        int index = dataGridView1.SelectedRows[0].Index;
                        int id_teacher = Convert.ToInt32(dataGridView1[0, index].Value);
                        SqlCommand command = new SqlCommand(
                          $"exec DEL_TEACHER @Id_teacher"
                          , sqlConnection);
                        command.Parameters.AddWithValue("Id_teacher", id_teacher);
                        delete(command, "Таблица содержит связанные данные!", "Объект удалён!");
                        upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [Status] AS Статус, \r\n[AcademicDegreeID] AS ИДУченойСтепени, [DepartmentID] AS ИДОтделения\r\nFROM [dbo].[Teacher]\r\n");
                    }
                    break;

                case "Почасовик":
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        int index = dataGridView1.SelectedRows[0].Index;
                        int id_hourlyworker = Convert.ToInt32(dataGridView1[0, index].Value);
                        SqlCommand command = new SqlCommand(
                          $"exec DEL_HOURLYWORKER @Id_hourlyworker"
                          , sqlConnection);
                        command.Parameters.AddWithValue("Id_hourlyworker", id_hourlyworker);
                        delete(command, "Таблица содержит связанные данные!", "Объект удалён!");
                        upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [RankID] AS ИДРанга\r\nFROM [dbo].[HourlyWorker]\r\n");
                    }
                    break;

                case "Плановая нагрузка почасовика":
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        int index = dataGridView1.SelectedRows[0].Index;
                        int hourlyid = Convert.ToInt32(dataGridView1[0, index].Value);
                        string position = dataGridView1[1, index].Value.ToString();
                        int year = Convert.ToInt32(dataGridView1[2, index].Value);
                        SqlCommand command = new SqlCommand(
                          $"exec DEL_PLANNED_LOAD @HourlyID, @Position, @Year"
                          , sqlConnection);
                        command.Parameters.AddWithValue("HourlyID", hourlyid);
                        command.Parameters.AddWithValue("Position", position);
                        command.Parameters.AddWithValue("Year", year);
                        delete(command, "Таблица содержит связанные данные!", "Объект удалён!");
                        upd_table("SELECT [HourlyID] AS ИДПочасовикаРабочего, [Position] AS Должность, [Year] AS Год, [MonthlyLoad] AS МесячнаяНагрузка\r\nFROM [dbo].[PlannedLoad]\r\n");
                    }
                    break;

                case "Фактическая нагрузка для почасовика":
                    MessageBox.Show("no del function");
                    break;

                case "Ранг почасовика":
                    MessageBox.Show("no del function");
                    break;

                case "Ученая степень":
                    MessageBox.Show("no del function");
                    break;

                case "Звание":
                    MessageBox.Show("no del function");
                    break;

                case "Почасовик-Кафедра":
                    MessageBox.Show("no del function");
                    break;

                default:
                    MessageBox.Show("error num table");
                    break;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application exApp = new Excel.Application();
            if (exApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
                return;
            }

            Excel.Workbook workbook = exApp.Workbooks.Add();
            Excel.Worksheet worksheet = (Excel.Worksheet)exApp.ActiveSheet;
            if (worksheet == null)
            {
                MessageBox.Show("Worksheet could not be created!");
                return;
            }

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("There are no rows in the DataGridView!");
                return;
            }

            Excel.Range range = worksheet.Range["A1:C1"];
            range.Merge();
            worksheet.Range["A1"].Value = "Данные в формате Excel:";
            worksheet.Range["A1"].Font.Bold = true;
            int startingColumnIndex = 4;

            for (int j = 0; j <= dataGridView1.ColumnCount - 1; j++)
            {
                worksheet.Cells[startingColumnIndex, j + 1] = dataGridView1.Columns[j].HeaderText.ToString();
                ((Excel.Range)worksheet.Cells[1, j + 1]).Columns.AutoFit();
            }

            for (int i = 0; i <= dataGridView1.RowCount - 1; i++)
            {
                for (int j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    if (dataGridView1[j, i].Value != null)
                    {
                        worksheet.Cells[startingColumnIndex + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                    }
                }
            }

((Excel.Range)worksheet.Cells[startingColumnIndex + 1, 1]).EntireRow.AutoFit();
            ((Excel.Range)worksheet.Range[worksheet.Cells[startingColumnIndex + 1, 1], worksheet.Cells[dataGridView1.RowCount + startingColumnIndex + 1, dataGridView1.ColumnCount]]).EntireColumn.AutoFit();

            string fileName = "MyExcelDocument.xlsx";
            workbook.SaveAs(fileName);

            exApp.Visible = true;

        }

        #endregion

        //Кнопки для Таблиц
        #region TableButtons
        private void button5_Click(object sender, EventArgs e)
        {
            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS Руководитель\r\nFROM [dbo].[Faculty];\r\n");
            checkTable = "Факультет";
        }

        private void button6_Click(object sender, EventArgs e)
        {
            upd_table("SELECT [Code] AS Код, [Name] AS Название, [Head] AS НачальникКафедры, [FacultyCode] AS КодФакультета\r\nFROM [dbo].[Department]\r\n");
            checkTable = "Кафедра";
        }

        private void button10_Click(object sender, EventArgs e)
        {
            upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [Status] AS Статус, \r\n[AcademicDegreeID] AS ИДУченойСтепени, [DepartmentID] AS ИДОтделения\r\nFROM [dbo].[Teacher]\r\n");
            checkTable = "Преподаватель";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            upd_table("SELECT [ID] AS ИД, [FullName] AS ПолноеИмя, [Position] AS Должность, [RankID] AS ИДРанга\r\nFROM [dbo].[HourlyWorker]\r\n");
            checkTable = "Почасовик";
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            //плановая нагрузка почасовиков
            upd_table("SELECT [HourlyID] AS ИДПочасовикаРабочего, [Position] AS Должность, [Year] AS Год, [MonthlyLoad] AS МесячнаяНагрузка\r\nFROM [dbo].[PlannedLoad]\r\n");
            checkTable = "Плановая нагрузка почасовика";
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //нужно будет переводить данные в нагрузку по месяцам.
            upd_table("SELECT\r\n    [HourlyID] AS [ИдентификаторРабочего],\r\n    HourlyWorker.FullName AS [ФИО],\r\n    YEAR([Date]) AS [Год],\r\n    MONTH([Date]) AS [Месяц],\r\n    SUM([HoursWorked]) AS [Кол-во_раб_часов]\r\nFROM\r\n    [dbo].[Workload], [HourlyWorker]\r\nGROUP BY\r\n    [HourlyID], HourlyWorker.FullName, YEAR([Date]), MONTH([Date])\r\n");
            checkTable = "Фактическая нагрузка для почасовика";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            upd_table("SELECT \r\n    [ID] AS 'Идентификатор',\r\n    [AcademicDegreeID] AS 'Идентификатор ученой степени',\r\n    [TitleID] AS 'Идентификатор должности',\r\n    [SalaryPerHour] AS 'Зарплата в час'\r\nFROM [dbo].[Rank];\r\n");
            checkTable = "Ранг почасовика";
        }

        private void button12_Click(object sender, EventArgs e)
        {
            upd_table("SELECT \r\n    [ID] AS 'Идентификатор',\r\n    [AcademicDegree] AS 'Ученая степень'\r\nFROM [dbo].[AcademicDegree];\r\n");
            checkTable = "Ученая степень";
        }

        private void button13_Click(object sender, EventArgs e)
        {
            upd_table("SELECT \r\n    [ID] AS 'Идентификатор',\r\n    [Title] AS 'Должность'\r\nFROM [dbo].[Title];\r\n");
            checkTable = "Звание";
        }

        private void button14_Click(object sender, EventArgs e)
        {
            upd_table("SELECT \r\n    [HourlyID] AS 'ID почасового работника',\r\n    [DepartmentCode] AS 'Код кафедры'\r\nFROM [dbo].[HourlyDepartment];\r\n");
            checkTable = "Почасовик-Кафедра";
        }


        #endregion
        //Кнопки для Запросов
        #region RequestButtons
        private void button17_Click(object sender, EventArgs e)
        {
            Form3 plForm = new Form3();
            plForm.panel1.Visible = false;
            plForm.panel5.Visible = true;
            plForm.panel2.Visible = false;
            plForm.panel3.Visible = false;

            DialogResult result;
            result = plForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            int panelIntElements1 = (int)plForm.numericUpDown1.Value;
            int panelIntElements2 = (int)plForm.numericUpDown2.Value;

            SqlCommand command = new SqlCommand("HourlyRegister", sqlConnection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@Month", panelIntElements1);
            command.Parameters.AddWithValue("@Year", panelIntElements2);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);

            dataGridView1.DataSource = table;

        }

        private void button16_Click(object sender, EventArgs e)
        {
            Form3 plForm = new Form3();
            plForm.panel1.Visible = false;
            plForm.panel5.Visible = false;
            plForm.panel2.Visible = false;
            plForm.panel3.Visible = true;

            DialogResult result;
            result = plForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            int panelIntElements1 = (int)plForm.numericUpDown5.Value;

            SqlCommand command = new SqlCommand("HourlyLoad", sqlConnection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@HourlyWorkerID", panelIntElements1);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);

            dataGridView1.DataSource = table;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            Form3 plForm = new Form3();
            plForm.panel1.Visible = false;
            plForm.panel5.Visible = false;
            plForm.panel2.Visible = true;
            plForm.panel3.Visible = false;

            DialogResult result;
            result = plForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            int panelIntElements1 = (int)plForm.numericUpDown6.Value;

            SqlCommand command = new SqlCommand("AllDepartments", sqlConnection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@year", panelIntElements1);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);

            dataGridView1.DataSource = table;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            Form3 plForm = new Form3();
            plForm.panel1.Visible = true;
            plForm.panel2.Visible = false;
            plForm.panel5.Visible = false;
            plForm.panel3.Visible = false;

            DialogResult result;
            result = plForm.ShowDialog(this);
            if (result == DialogResult.Cancel)
                return;
            int panelIntElements1 = (int)plForm.numericUpDown3.Value;
            int panelIntElements2 = (int)plForm.numericUpDown4.Value;

            SqlCommand command = new SqlCommand("SelectDepartment", sqlConnection);
            command.CommandType = CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@year", panelIntElements1);
            command.Parameters.AddWithValue("@departmentCode", panelIntElements2);

            SqlDataAdapter adapter = new SqlDataAdapter(command);
            System.Data.DataTable table = new System.Data.DataTable();
            adapter.Fill(table);

            dataGridView1.DataSource = table;


        }
        #endregion
        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }


    }

}


