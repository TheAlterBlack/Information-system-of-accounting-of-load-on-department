using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Scripting.Hosting;
using System.IO;
using System.Diagnostics;
using Spire.Xls;
using Spire.DataExport;
using MySql.Data.MySqlClient;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System.Reflection;


namespace Diploma_Add_Py
{
    public partial class Form1 : Form
    {
        public Form2 form;
        public Form3 form3;
        
        public Form1()
        {
            InitializeComponent();
            WindowState = FormWindowState.Maximized;
            tabControl1.Dock = DockStyle.Fill;

            this.dataGridView3.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridView3.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.dataGridView3.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
        }

        private void button1_Click(object sender, EventArgs e)//Парсинг файлу
        {
            try
            {
                Process PYprocces = new Process();
                PYprocces.StartInfo.FileName = @"Parcing_module.py";
                PYprocces.Start();
                PYprocces.WaitForExit();
                MessageBox.Show("Парсинг завершено! Перейдіть до редагування");
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void button2_Click(object sender, EventArgs e)//Формування таблиці для редагування
        {
            try
            {
                int i = 0;
                for (int j = 0; j < 50; j++)
                {
                    string fileName = String.Format(@"ITP{0}.xlsx", j);
                    if (File.Exists(fileName) == true) { i += 1; }
                }

                Spire.Xls.Workbook MerBook = new Spire.Xls.Workbook();
                MerBook.LoadFromFile(@"ITP0.xlsx");

                Spire.Xls.Worksheet MerSheet = MerBook.Worksheets[0];
                int index = 1;
                MessageBox.Show("Початок формування таблиці Excel!");
                do
                {
                    Spire.Xls.Workbook SouBook1 = new Spire.Xls.Workbook();
                    string fileName = String.Format(@"ITP{0}.xlsx", index);
                    SouBook1.LoadFromFile(fileName);

                    int a = SouBook1.Worksheets[0].LastRow;
                    int b = SouBook1.Worksheets[0].LastColumn;

                    SouBook1.Worksheets[0].Range[1, 1, a, b].Copy(MerSheet.Range[MerSheet.LastRow + 1, 1, a + MerSheet.LastRow, b - 1]);
                    MerBook.SaveToFile(@"result_1.xlsx", ExcelVersion.Version2010);
                    index++;
                } while (index < i);
                MessageBox.Show("Таблиця Excel з навантаженням сформована!\nПерейдіть до редагування");


                //Add data grid view
                //Create a new workbook
                Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

                //Load a file and imports its data
                workbook.LoadFromFile(@"result_1.xlsx");

                //Initialize worksheet
                Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
                int i1 = workbook.Worksheets[0].LastRow;
                int i2 = workbook.Worksheets[0].LastColumn - 1;

                // get the data source that the grid is displaying data for
                this.dataGridView1.DataSource = sheet.ExportDataTable(sheet.Range[1, 1, i1, i2], false);
            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void button3_Click(object sender, EventArgs e)//Додавання даних у БД
        {
            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();
            
            
            string query = "";
            string semester = "Осінній семестр";
            int rows = 0;
            do
            {
                if (dataGridView1.Rows[rows].Cells[0].Value.ToString() == "Осінній семестр")
                {
                    semester = "Осінній семестр";
                    rows++;
                }
                else if (dataGridView1.Rows[rows].Cells[0].Value.ToString() == "Весняний семестр")
                {
                    semester = "Весняний семестр";
                    rows++;
                }
                else
                {
                    string zeroCell = dataGridView1.Rows[rows].Cells[0].Value.ToString(); //№ з./п.
                    string firstCell = dataGridView1.Rows[rows].Cells[1].Value.ToString(); //Шифр циклу
                    int res;
                    bool isInt = Int32.TryParse(zeroCell, out res);
                    bool isDiscipline = firstCell == "ПП" && firstCell == "ПН" &&
                        firstCell == "ГС" && firstCell == "ПА" &&
                        firstCell == "П";

                    if ((isInt == true) || (isDiscipline == true))
                    {
                        string year = dataGridView1.Rows[rows].Cells[3].Value.ToString();
                        string groups = dataGridView1.Rows[rows].Cells[4].Value.ToString();
                        string[] delimiterChars = { ", " };
                        string[] group = groups.Split(delimiterChars, System.StringSplitOptions.RemoveEmptyEntries);

                        //Цикл додавання груп студентів
                        foreach (var g in group)
                        {
                            string st = g.Replace("\n", "");
                            query = "SELECT group_name FROM `group`";
                            MySqlCommand command = new MySqlCommand(query, conn);
                            MySqlDataReader reader = command.ExecuteReader();
                            int check = 0;
                            while (reader.Read())
                            {
                                if (reader[0].ToString() == st) { check = 1; }
                            }
                            reader.Close();

                            if (check == 0)
                            {
                                query = String.Format("INSERT INTO `group`(group_name, year) VALUES (\"{0}\", {1})", st, year);
                                command = new MySqlCommand(query, conn);
                                command.ExecuteNonQuery();
                            }
                        }
                        //Додаємо дисципліну
                        string discipline_name = dataGridView1.Rows[rows].Cells[2].Value.ToString();

                        string groups1 = dataGridView1.Rows[rows].Cells[4].Value.ToString();
                        string[] delimiterChars1 = { ", " };
                        string[] group1 = groups1.Split(delimiterChars1, System.StringSplitOptions.RemoveEmptyEntries);

                        query = "SELECT group_id, group_name FROM `group`";
                        MySqlCommand comm = new MySqlCommand(query, conn);
                        string group_id_list = "";
                        MySqlDataReader read = comm.ExecuteReader();


                        while (read.Read())
                        {
                            foreach (var g in group1)
                            {
                                string st = g.Replace("\n", "");
                                if (read[1].ToString() == st)
                                {
                                    group_id_list = group_id_list + String.Format(",{0}", read[0].ToString());
                                    //MessageBox.Show(read[0].ToString() + " " + read[1].ToString());
                                }
                            }

                        }
                        read.Close();

                        //Оставить сумму или сделать проверку?
                        string control_d = dataGridView1.Rows[rows].Cells[16].Value.ToString() + dataGridView1.Rows[rows].Cells[17].Value.ToString();
                        string ind_d = dataGridView1.Rows[rows].Cells[18].Value.ToString();

                        query = String.Format("SELECT discipline_name FROM discipline WHERE semester = \"{0}\"", semester);
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        MySqlDataReader reader1 = cmd.ExecuteReader();
                        int check1 = 0;
                        while (reader1.Read())
                        {
                            if (reader1[0].ToString() == discipline_name) { check1 = 1; }
                        }
                        reader1.Close(); // закрываем reader
                        int discipline_id = 0;
                        if (check1 == 0)
                        {
                            query = String.Format("INSERT INTO `discipline`(discipline_name, code, semester, check_type, individual, id_group_list)" +
                            "VALUES (\"{0}\",\"{1}\",\"{2}\",\"{3}\",\"{4}\",\"{5}\")", discipline_name, firstCell, semester, control_d, ind_d, group_id_list);
                            cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();

                            query = String.Format("SELECT discipline_id FROM `discipline` WHERE discipline_name=\"{0}\"", discipline_name);
                            cmd = new MySqlCommand(query, conn);
                            string discipline_id_st = cmd.ExecuteScalar().ToString();
                            discipline_id = Int32.Parse(discipline_id_st);//Запам'ятовуємо id доданої дисципліни

                            //Додаємо навантаження
                            string[] work_type = { "Лекція", "Практичне, семінарське заняття", "Лабораторне заняття", "Інд. зан. і конс. роб. за розкладом", "Атестаційні заходи", "Проведення тестування", "Перевірка індивідуальних завдань", "Керівництво і приймання інд. завд.", "Консул. з навч. дисц. протягом. сем.", "Консультації передекзаменаційні", "Заліки, підсумковий сем. контроль", "Екзамени, додатковий сем. контроль", "Керів., конс. та реценз. кваліф. проектів", "Державна атестація", "Керів. аспір., докт., здоб., стажування викл.", "Керівництво практикою", "Інше", "Погодинна оплата" };
                            //'Лекція',"Практичне, семінарське заняття",'Лабораторне заняття','Інд. зан. і конс. роб. за розкладом','Атестаційні заходи','Проведення тестування','Перевірка індивідуальних завдань','Керівництво і приймання інд. завд.','Консул. з навч. дисц. протягом. сем.','Консультації передекзаменаційні','Заліки, підсумковий сем. контроль','Екзамени, додатковий сем. контроль','Керів., конс. та реценз. кваліф. проектів','Державна атестація','Керів. аспір., докт., здоб., стажування викл.','Керівництво практикою','Інше','Погодинна оплата'
                            string hours = "";
                            for (int i = 20; i < 37; i++)
                            {
                                var value = dataGridView1.Rows[rows].Cells[i].Value;
                                hours = (value != null || value!="" ? value.ToString() : "0");
                                if (hours == "") { hours = "0"; }
                                hours = hours.Replace(",", ".");
                                query = String.Format("INSERT INTO computation(discipline_id, work_type, hours) VALUES ({0}, \"{1}\", {2})", discipline_id, work_type[i - 20], hours);
                                cmd = new MySqlCommand(query, conn);
                                cmd.ExecuteNonQuery();
                            }

                            hours = dataGridView1.Rows[rows].Cells[38].Value.ToString();
                            if (hours == "") { hours = "0"; }
                            hours = hours.Replace(",", ".");
                            query = String.Format("INSERT INTO computation(discipline_id, work_type, hours) VALUES ({0}, \"Погодинна оплата\", {1})", discipline_id, hours);
                            cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    rows++;
                }
            } while (rows < dataGridView1.Rows.Count - 1);
            MessageBox.Show("Дані завантажені у БД!");
            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }

        private void button4_Click(object sender, EventArgs e)//Додавання викладача
        {


            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();

                string surname = textBox1.Text;
            string name = textBox2.Text;
            string patr = textBox3.Text;

            string query = "SELECT name, surname, patronymic FROM professor";
            MySqlCommand command = new MySqlCommand(query, conn);
            MySqlDataReader reader = command.ExecuteReader();

            int check = 0;
            while (reader.Read())
            {
                if (reader[0].ToString() == name ||
                    reader[1].ToString() == surname ||
                    reader[2].ToString() == patr) { check = 1; }
            }
            reader.Close();

            if (check == 0)
            {
                query = String.Format("INSERT INTO professor(name, surname, patronymic) VALUES (\"{0}\", \"{1}\", \"{2}\")", name, surname, patr);
                command = new MySqlCommand(query, conn);
                command.ExecuteNonQuery();
                MessageBox.Show("Викладач: " + surname + " " + name + " " + patr + " доданий до БД!");
                textBox1.Clear();
                textBox2.Clear();
                textBox3.Clear();
            }
            else { MessageBox.Show("Виклач " + surname + " " + name + " " + patr + " вже доданий до БД!"); }

            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }

        private void button5_Click(object sender, EventArgs e)//Завантаження груп для редагування
        {
            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();

                string query = "SELECT count(group_name) FROM `group`";
            MySqlCommand com = new MySqlCommand(query, conn);
            int groups_count;
            bool check = Int32.TryParse(com.ExecuteScalar().ToString(), out groups_count);
            if (groups_count != 0)
            {
                string[,] groups = new string[2, groups_count];

                query = "SELECT group_id, group_name FROM `group`";
                com = new MySqlCommand(query, conn);
                MySqlDataReader read = com.ExecuteReader();
                int k = 0;
                while (read.Read())
                {
                    groups[0, k] = read[0].ToString();
                    groups[1, k] = read[1].ToString();
                    k++;
                }
                read.Close();

                for (int i = 0; i < k; i++)
                {
                    comboBox1.Items.Add(groups[1, i]);
                }
                MessageBox.Show("Групи додані у випадаючий список");
            }
            else { MessageBox.Show("У БД немає груп!"); }
            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }

        private void button6_Click(object sender, EventArgs e)//Додавання групи до БД
        {
            string group = comboBox1.Text;
            if (group != "")
            {
                if (textBox4.Text != "")
                {
                    string conn_s;
                    using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
                    {
                        conn_s = sr.ReadToEnd();
                    }

                    try
                    {
                        MySqlConnection conn = new MySqlConnection(conn_s);
                        conn.Open();
                        string query = String.Format("SELECT group_id FROM `group` WHERE group_name=\"{0}\"", group);
                    MySqlCommand com = new MySqlCommand(query, conn);
                    string group_id = com.ExecuteScalar().ToString();

                    query = String.Format("UPDATE `group` SET quantity=\"{0}\" WHERE group_id={1}", textBox4.Text, group_id);
                    com = new MySqlCommand(query, conn);
                    com.ExecuteNonQuery();
                    conn.Close();
                    }
                    catch (Exception ex) { MessageBox.Show(ex.ToString()); }
                    MessageBox.Show("Дані додані до БД!");
                    textBox4.Clear();
                }
                else { MessageBox.Show("Введіть кількість студентів"); }
            }
            else { MessageBox.Show("Оберіть групу у випадаючому списку"); }
        }

        private void button7_Click(object sender, EventArgs e)//Додавання даних у таблицю для розподілу
        {
            dataGridView2.Rows.Clear();

            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();

                string query = "SELECT COUNT( discipline_name ) FROM discipline";
            MySqlCommand com = new MySqlCommand(query, conn);
            int count;
            bool n = Int32.TryParse(com.ExecuteScalar().ToString(), out count);
            if (n == true)
            {
                string[,] discipline_names = new string[count, 3];

                //query = "SELECT d.semester, d.discipline_name, d.id_group_list, c.work_type, c.hours FROM discipline d, computation c WHERE c.discipline_id = d.discipline_id AND(c.work_type = \"Лекція\" OR c.work_type = \"Консультації передекзаменаційні\" OR c.work_type = \"Заліки, підсумковий сем. контроль\" OR c.work_type = \"Практичне, семінарське заняття\" OR c.work_type = \"Лабораторне заняття\" OR c.work_type = \"Інд. зан. і конс. роб. за розкладом\" OR c.work_type = \"Атестаційні заходи\" OR c.work_type = \"Перевірка індивідуальних завдань\" OR c.work_type = \"Керівництво і приймання інд. завд.\")";
                query = "SELECT semester, discipline_name, id_group_list FROM discipline";
                com = new MySqlCommand(query, conn);
                MySqlDataReader read = com.ExecuteReader();
                int i = 0;
                while (read.Read())
                {
                    discipline_names[i, 0] = read[0].ToString();
                    discipline_names[i, 1] = read[1].ToString();
                    discipline_names[i, 2] = read[2].ToString();
                    i++;
                }
                read.Close();
                
                for (int j = 0; j < count; j++) {
                    
                    query = String.Format("SELECT c.hours FROM discipline d, computation c "+
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Лекція\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                    com = new MySqlCommand(query, conn);

                    var value = com.ExecuteScalar().ToString();
                    string c21 = ((value != null || value != "") ? value.ToString() : "");



                    query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Консультації передекзаменаційні\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                    com = new MySqlCommand(query, conn);

                    value = com.ExecuteScalar().ToString();
                    string c30 = ((value != null || value != "") ? value.ToString() : "");



                    query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Заліки, підсумковий сем. контроль\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                    com = new MySqlCommand(query, conn);

                    value = com.ExecuteScalar().ToString();
                    string c32 = ((value != null || value != "") ? value.ToString() : "");
                    
                    string[] group = discipline_names[j, 2].Split(new char[] {','});
                    string groups="";
                    int groups_count=0;

                    foreach (string s in group)
                    {
                        if (s != "")
                        {
                            query = String.Format("SELECT group_name FROM `group` WHERE group_id={0}", s);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            groups += ((value != null || value != "") ? value.ToString() : "") + ",";
                            groups_count++;
                        }
                    }
                    if (discipline_names[j, 1].Contains("- КР") == false)
                        dataGridView2.Rows.Add(discipline_names[j, 0], discipline_names[j, 1], groups, c21, c30, c32);

                    foreach (string s in group)
                    {
                        if (s != "") {
                            query = String.Format("SELECT group_name FROM `group` WHERE group_id={0}", s);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            groups += ((value != null || value != "") ? value.ToString() : "") + ",";

                            string group1 = ((value != null || value != "") ? value.ToString() : "");

                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Практичне, семінарське заняття\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);

                            value = com.ExecuteScalar().ToString();
                            string c22 = ((value != null || value != "") ? value.ToString() : "");
                            float res = float.Parse(c22);
                            if (res != 0 || groups_count!=0)
                            {
                                res = res / (float)groups_count;
                                c22 = res.ToString("0.##");
                            }



                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Лабораторне заняття\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            string c23 = ((value != null || value != "") ? value.ToString() : "");
                            res = float.Parse(c23);
                            if (res != 0 || groups_count != 0)
                            {
                                res = res / (float)groups_count;
                                c23 = res.ToString("0.##");
                            }


                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Інд. зан. і конс. роб. за розкладом\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            string c24 = ((value != null || value != "") ? value.ToString() : "");
                            res = float.Parse(c24);
                            if (res != 0 || groups_count != 0)
                            {
                                res = res / (float)groups_count;
                                c24 = res.ToString("0.##");
                            }


                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Атестаційні заходи\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            string c25 = ((value != null || value != "") ? value.ToString() : "");
                            res = float.Parse(c25);
                            if (res != 0 || groups_count != 0)
                            {
                                res = res / (float)groups_count;
                                c25 = res.ToString("0.##");
                            }


                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Перевірка індивідуальних завдань\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            string c27 = ((value != null || value != "") ? value.ToString() : "");
                            res = float.Parse(c27);
                            if (res != 0 || groups_count != 0)
                            {
                                res = res / (float)groups_count;
                                c27 = res.ToString("0.##");
                            }


                            query = String.Format("SELECT c.hours FROM discipline d, computation c " +
                                            "WHERE c.discipline_id = d.discipline_id " +
                                            "AND c.work_type = \"Керівництво і приймання інд. завд.\" " +
                                            "AND d.discipline_name = \"{0}\"", discipline_names[j, 1]);
                            com = new MySqlCommand(query, conn);
                            value = com.ExecuteScalar().ToString();
                            string c28 = ((value != null || value != "") ? value.ToString() : "");
                            res = float.Parse(c28);
                            if (res != 0 || groups_count != 0)
                            {
                                res = res / (float)groups_count;
                                c28 = res.ToString("0.##");
                            }

                            dataGridView2.Rows.Add(discipline_names[j, 0], discipline_names[j, 1], group1,null,null,null, c22, c23, c24, c25, c27, c28);
                        }
                    }
                }
            }

            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }
        
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)//Считывание клика на колонку с преподами
        {
            if (dataGridView2.CurrentCell.ColumnIndex.Equals(12) && e.RowIndex != -1)
            {
                if (dataGridView2.CurrentCell != null && dataGridView2.CurrentCell.Value == null)
                {
                    //Инициализируем DataTransfer.data
                    DataTransfer.data = new object[] { "" };

                    int CursorX = Cursor.Position.X;
                    int CursorY = Cursor.Position.Y;
                    
                    Form2 form2 = new Form2(CursorX - 277, CursorY - 175);
                    form2.ShowDialog();
                    form2.Location = new System.Drawing.Point(CursorX, CursorY);
                    dataGridView2.CurrentCell.Value = DataTransfer.data[0].ToString();
                    form2.Close();
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)//Зберегти тимчасову таблицю у файл
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                ExcelWorkSheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            }

            for (int i = 2; i < dataGridView2.Rows.Count+1; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    var value = dataGridView2.Rows[i - 2].Cells[j].Value;
                    string s = (value == null ? "" : value.ToString());

                    if (s != "")
                    {
                        ExcelWorkSheet.Cells[i, j+1] = s;
                    }
                }
            }

            ExcelWorkBook.SaveAs(System.IO.Directory.GetCurrentDirectory().ToString()+"\\"+"temp_table.xls");
            ExcelApp.AlertBeforeOverwriting = false;
            ExcelApp.Visible = false;
            ExcelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ExcelApp);
            GC.Collect();
            MessageBox.Show("Таблиця успішно збережена!", "Збереження таблиці", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button10_Click(object sender, EventArgs e)//Вивід тимчасової таблиці
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "temp_table.xls");
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = ObjExcel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            Microsoft.Office.Interop.Excel.Range rg = null;

            Int32 row = 2;
            dataGridView2.Rows.Clear();
            List<String> arr = new List<string>();
            while (ObjWorkSheet.get_Range("a" + row, "a" + row).Value != null)
            {
                rg = ObjWorkSheet.get_Range("a" + row, "u" + row);
                foreach (Microsoft.Office.Interop.Excel.Range item in rg)
                {
                    try
                    {
                        arr.Add(item.Value.ToString().Trim());
                    }
                    catch { arr.Add(""); }
                }
                if (arr[12] == "") { arr[12] = null; }
                dataGridView2.Rows.Add(arr[0], arr[1], arr[2], arr[3], arr[4], arr[5], arr[6], arr[7], arr[8], arr[9], arr[10], arr[11], arr[12]);
                arr.Clear();
                row++;
            }
            MessageBox.Show("Таблиця успішно додана!", "Відкриття таблиці", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ObjWorkBook.Close(false, "", null);
            
            ObjExcel.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ObjExcel);
            GC.Collect();
        }

        private void завантажитиToolStripMenuItem_Click(object sender, EventArgs e)//Додавання файлу навантаження
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Pdf Files|*.pdf";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.pdf"))
                    File.Delete(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.pdf");
                var filePath = openFileDialog1.FileName;
                File.Copy(filePath, System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.pdf");
                MessageBox.Show("Файл збережено!");
            }
        }

        private void коефСтавкиToolStripMenuItem_Click(object sender, EventArgs e)//Форма коефіцієнта ставки
        {
            form3 = new Form3();
            form3.ShowDialog();
            form3.Close();
        }

        private void button8_Click(object sender, EventArgs e)//Збереження розподіленого навантаження у БД
        {
            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();
                MySqlCommand comm;
            string query;
            if (dataGridView2.RowCount > 1)
            {
                int RowsCount = dataGridView2.RowCount;
                for (int i = 0; i < RowsCount - 1; i++)
                {
                    if (dataGridView2.Rows[i].Cells[12].Value != null)
                    {
                        string groups_id = "";
                        string[] groups = dataGridView2.Rows[i].Cells[2].Value.ToString().Split(',');
                        foreach (var elem in groups)
                        {
                            if (elem != "")
                            {
                                query = String.Format("SELECT group_id FROM `group` WHERE group_name=\"{0}\"", elem);
                                comm = new MySqlCommand(query, conn);
                                var v = comm.ExecuteScalar();
                                groups_id += ((v != null || v.ToString() != "") ? v.ToString() + "," : "");
                            }
                        }

                        string[] professor = dataGridView2.Rows[i].Cells[12].Value.ToString().Split(' ');
                        string surname = professor[0];
                        string name = professor[1];
                        string patr = professor[2];
                        query = String.Format("SELECT professor_id FROM `professor` WHERE name=\"{0}\" AND surname=\"{1}\" AND patronymic=\"{2}\"", name, surname, patr);
                        comm = new MySqlCommand(query, conn);
                        var value = comm.ExecuteScalar();
                        string professor_id = ((value != null) ? value.ToString() : "");

                        query = String.Format("SELECT discipline_id FROM `discipline` WHERE discipline_name=\"{0}\" AND semester=\"{1}\"", dataGridView2.Rows[i].Cells[1].Value.ToString(), dataGridView2.Rows[i].Cells[0].Value.ToString());
                        comm = new MySqlCommand(query, conn);
                        value = comm.ExecuteScalar();
                        string discipline_id = ((value != null) ? value.ToString() : "");

                        string[] comp = new string[9];
                        float sum_hours = 0;

                        for (int j = 3; j < 12; j++)
                        {
                            value = dataGridView2.Rows[i].Cells[j].Value;
                            comp[j - 3] = (value != null ? value.ToString() : "0");
                            if (comp[j - 3] == "") { comp[j - 3] = "0"; }
                            sum_hours += float.Parse(comp[j - 3]);
                            comp[j - 3] = comp[j - 3].Replace(',', '.');
                        }

                        query = String.Format("SELECT comp_prof_id FROM comp_prof WHERE discipline_id={0} AND professor_id={1} AND groups=\"{2}\"", discipline_id, professor_id, groups_id);
                        comm = new MySqlCommand(query, conn);
                        var check = comm.ExecuteScalar();
                        bool chek = (check != null ? chek = true : chek = false);

                        if (chek == false)
                        {
                            if (discipline_id != "" || professor_id != "")
                            {
                                query = String.Format("INSERT INTO comp_prof(discipline_id, professor_id, " +
                                                        "groups, lectures21, ConsultBeforeExam30, FinalSemControl31, " +
                                                        "Practice22, Laboratory23, Indiv24, Attestation25, CheckInd27, " +
                                                        "LeaderInd28) VALUES (\"{0}\", \"{1}\", \"{2}\", {3}, {4}, {5}, {6}, {7}, " +
                                                        "{8}, {9}, {10}, {11})", discipline_id, professor_id, groups_id,
                                                        comp[0], comp[1], comp[2], comp[3], comp[4], comp[5], comp[6],
                                                        comp[7], comp[8]);
                                comm = new MySqlCommand(query, conn);
                                comm.ExecuteNonQuery();

                                query = String.Format("SELECT hours FROM professor WHERE professor_id={0}", professor_id);
                                comm = new MySqlCommand(query, conn);
                                var a = comm.ExecuteScalar();
                                float hours = (a != null ? float.Parse(a.ToString()) : 0);
                                hours = hours + sum_hours;

                                float r;
                                string path = System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.txt";
                                using (StreamReader sr = new StreamReader(path))
                                {
                                    r = float.Parse(sr.ReadToEnd());
                                }
                                float rate = hours / r;

                                string hours_s = hours.ToString();
                                hours_s = hours_s.Replace(',', '.');

                                string rate_s = rate.ToString();
                                rate_s = rate_s.Replace(',', '.');

                                query = String.Format("UPDATE professor SET hours={0}, rate={1} WHERE professor_id={2}",
                                    hours_s, rate_s, professor_id);
                                comm = new MySqlCommand(query, conn);
                                comm.ExecuteNonQuery();

                                MessageBox.Show("Дані додані до БД!");
                            }
                            else { MessageBox.Show("Ви не вказали дисципліну або викладача у " + i + " рядку"); }
                        }
                        else { MessageBox.Show("Навантаження у рядку " + (i + 1) + " уже є в БД!"); }
                    }
                    else { MessageBox.Show("У рядку " + (i+1) + " не вказаний викладач"); }
                }
            }
            else { MessageBox.Show("Додайте дані в таблицю"); }
            
            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }

        private void button12_Click(object sender, EventArgs e)//Перегляд викладачів з навантаженням
        {
            button11.Enabled = true;

            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();

                string query = "SELECT name, surname, patronymic, hours, rate FROM professor WHERE hours IS NOT NULL";
            MySqlCommand comm = new MySqlCommand(query, conn);
            MySqlDataReader read = comm.ExecuteReader();
            
            while (read.Read()) {
                string professor = "", hours = "", rate = "";
                if (read[1].ToString() != null) {
                    professor += read[1].ToString() + " ";
                }
                if (read[0].ToString() != null)
                {
                    professor += read[0].ToString() + " ";
                }
                if (read[2].ToString() != null)
                {
                    professor += read[2].ToString() + " ";
                }
                if (read[3].ToString() != null)
                {
                    hours = read[3].ToString();
                }
                if (read[4].ToString() != null)
                {
                    rate = read[4].ToString();
                }
                comboBox2.Items.Add(professor);
                dataGridView3.Rows.Add(professor, hours, rate);
            }
            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }
        }

        private void button11_Click(object sender, EventArgs e)//Вивод у Word
        {
            Microsoft.Office.Interop.Word.Application wordApp;
            Microsoft.Office.Interop.Word.Document wordDoc;
            wordApp = new Microsoft.Office.Interop.Word.Application();
            wordDoc = wordApp.Documents.Add(System.IO.Directory.GetCurrentDirectory().ToString() + "\\Res.docx");
            Object missing = Type.Missing;
            Microsoft.Office.Interop.Word.Range selText;
            selText = wordDoc.Range(wordDoc.Content.Start, wordDoc.Content.End);
            Microsoft.Office.Interop.Word.Find find = wordApp.Selection.Find;

            string conn_s;
            using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
            {
                conn_s = sr.ReadToEnd();
            }

            try
            {
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();


                string[] s = comboBox2.Text.Split(' ');
            string name=s[1];
            string surname=s[0];
            string patronymic=s[2];
            string hours = "";
            string rate = "";

            string professor_id = "";
            string query = "SELECT professor_id, name, surname, patronymic, hours, rate FROM professor";
            MySqlCommand comm = new MySqlCommand(query, conn);
            MySqlDataReader read = comm.ExecuteReader();
            while (read.Read()) {
                if (read[0] != null || read[1] != null || read[2] != null || read[3] != null || read[4] != null || read[5] != null) {
                    if (read[1].ToString() == name || read[2].ToString() == surname || read[3].ToString() == patronymic)
                    {
                        professor_id = read[0].ToString();
                        hours = read[4].ToString();
                        rate = read[5].ToString();
                    }
                }
            }
            read.Close();

            find.Text = "[professor]";
            find.Replacement.Text = Convert.ToString(surname + " " + name + " " + patronymic);

            Object wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            Object replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
            MatchCase: false,
            MatchWholeWord: false,
            MatchWildcards: false,
            MatchSoundsLike: missing,
            MatchAllWordForms: false,
            Forward: true,
            Wrap: wrap,
            Format: false,
            ReplaceWith: missing, Replace: replace);

            find.Text = "[date]";
            find.Replacement.Text = Convert.ToString(DateTime.Now.ToString("dd MMMM yyyy"));

            wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
            MatchCase: false,
            MatchWholeWord: false,
            MatchWildcards: false,
            MatchSoundsLike: missing,
            MatchAllWordForms: false,
            Forward: true,
            Wrap: wrap,
            Format: false,
            ReplaceWith: missing, Replace: replace);

            find.Text = "[hours]";
            find.Replacement.Text = Convert.ToString(hours);

            wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
            MatchCase: false,
            MatchWholeWord: false,
            MatchWildcards: false,
            MatchSoundsLike: missing,
            MatchAllWordForms: false,
            Forward: true,
            Wrap: wrap,
            Format: false,
            ReplaceWith: missing, Replace: replace);

            find.Text = "[rate]";
            find.Replacement.Text = Convert.ToString(rate);

            wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;
            replace = Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll;
            find.Execute(FindText: Type.Missing,
            MatchCase: false,
            MatchWholeWord: false,
            MatchWildcards: false,
            MatchSoundsLike: missing,
            MatchAllWordForms: false,
            Forward: true,
            Wrap: wrap,
            Format: false,
            ReplaceWith: missing, Replace: replace);


            query = "SELECT count(groups) FROM comp_prof";
            comm = new MySqlCommand(query, conn);
            var value = comm.ExecuteScalar();
            int count_groups = (value != null ? Int32.Parse(value.ToString()) : 0);

            query = "SELECT groups FROM comp_prof";
            comm = new MySqlCommand(query, conn);
            read = comm.ExecuteReader();
            string[,] groups_id = new string[count_groups, 4];
            int k = 0;
            while (read.Read()) {
                if (read[0] != null)
                {
                    groups_id[k, 0] = read[0].ToString();
                    k++;
                }
            }
            read.Close();
            int kol=0;
            for (int i = 0; i < count_groups; i++)
            {
                string[] gr = groups_id[i, 0].Split(',');
                kol = 0;
                foreach (var g in gr) {
                    if (g != "") {
                        query = String.Format("SELECT group_name FROM `group` WHERE group_id={0}", g);
                        comm = new MySqlCommand(query, conn);
                        var v = comm.ExecuteScalar();
                        groups_id[i, 1] += (v!=null ? v.ToString() : "")+",";

                        query = String.Format("SELECT year FROM `group` WHERE group_id={0}", g);
                        comm = new MySqlCommand(query, conn);
                        v = comm.ExecuteScalar();
                        groups_id[i, 2] = (v != null ? v.ToString() : "");

                        query = String.Format("SELECT quantity FROM `group` WHERE group_id={0}", g);
                        comm = new MySqlCommand(query, conn);
                        v = comm.ExecuteScalar();
                        kol = kol + (v != null ? Int32.Parse(v.ToString()) : 0);
                        groups_id[i, 3] = kol.ToString();
                    }
                }
            }

            query = String.Format("SELECT d.discipline_name, d.check_type, d.individual, cp.groups, cp.lectures21, cp.ConsultBeforeExam30, cp.FinalSemControl31, cp.Practice22, cp.Laboratory23, cp.Indiv24, cp.Attestation25, cp.CheckInd27, cp.LeaderInd28 "+
                                    "FROM discipline d, comp_prof cp "+
                                    "WHERE cp.discipline_id = d.discipline_id AND cp.professor_id = {0} AND d.semester = \"Осінній семестр\"",professor_id);
            comm = new MySqlCommand(query, conn);
            read = comm.ExecuteReader();
            int row = 0;
            float[] sum_cells = new float[9]; 
            while (read.Read()) {

                if (read[0] != null) {
                    wordDoc.Tables[1].Cell(2 + row, 1).Range.Text = read[0].ToString();
                    if (read[1] != null || read[2] != null)
                    {
                        wordDoc.Tables[1].Cell(2 + row, 5).Range.Text = read[1].ToString();
                        wordDoc.Tables[1].Cell(2 + row, 6).Range.Text = read[2].ToString();
                    }
                    if (read[3] != null) {
                        for (int i = 0; i < count_groups; i++) {
                            if (read[3].ToString().Equals(groups_id[i, 0])) {
                                wordDoc.Tables[1].Cell(2 + row, 2).Range.Text = groups_id[i, 2].ToString();
                                wordDoc.Tables[1].Cell(2 + row, 3).Range.Text = groups_id[i, 1].ToString().Trim(',');
                                wordDoc.Tables[1].Cell(2 + row, 4).Range.Text = groups_id[i, 3].ToString();
                            }
                        }
                    }
                    float sum_c = 0;
                    for (int i = 4; i <= 12; i++) {
                        if (read[i] != null) {
                            wordDoc.Tables[1].Cell(2 + row, i + 3).Range.Text = read[i].ToString();
                            sum_c += (read[i].ToString()!="" ? float.Parse(read[i].ToString()) : 0);
                            sum_cells[i - 4] += (read[i].ToString() != "" ? float.Parse(read[i].ToString()) : 0);
                        }
                    }
                    wordDoc.Tables[1].Cell(2 + row, 16).Range.Text = sum_c.ToString();
                }
                wordDoc.Tables[1].Rows.Add();
                row++;
            }
            read.Close();
            float sum_sum_cells = 0;
            for (int i = 0; i < 9; i++) {
                wordDoc.Tables[1].Cell(2 + row, i + 7).Range.Text = sum_cells[i].ToString();
                sum_sum_cells += sum_cells[i];
            }
            wordDoc.Tables[1].Cell(2 + row, 16).Range.Text = sum_sum_cells.ToString();




            query = String.Format("SELECT d.discipline_name, d.check_type, d.individual, cp.groups, cp.lectures21, cp.ConsultBeforeExam30, cp.FinalSemControl31, cp.Practice22, cp.Laboratory23, cp.Indiv24, cp.Attestation25, cp.CheckInd27, cp.LeaderInd28 " +
                                    "FROM discipline d, comp_prof cp " +
                                    "WHERE cp.discipline_id = d.discipline_id AND cp.professor_id = {0} AND d.semester = \"Весняний семестр\"", professor_id);
            comm = new MySqlCommand(query, conn);
            read = comm.ExecuteReader();
            row = 0;
            float[] sum_cells1 = new float[9];
            while (read.Read())
            {

                if (read[0] != null)
                {
                    wordDoc.Tables[2].Cell(2 + row, 1).Range.Text = read[0].ToString();
                    if (read[1] != null || read[2] != null)
                    {
                        wordDoc.Tables[2].Cell(2 + row, 5).Range.Text = read[1].ToString();
                        wordDoc.Tables[2].Cell(2 + row, 6).Range.Text = read[2].ToString();
                    }
                    if (read[3] != null)
                    {
                        for (int i = 0; i < count_groups; i++)
                        {
                            if (read[3].ToString().Equals(groups_id[i, 0]))
                            {
                                wordDoc.Tables[2].Cell(2 + row, 2).Range.Text = groups_id[i, 2].ToString();
                                wordDoc.Tables[2].Cell(2 + row, 3).Range.Text = groups_id[i, 1].ToString().Trim(',');
                                wordDoc.Tables[2].Cell(2 + row, 4).Range.Text = groups_id[i, 3].ToString();
                            }
                        }
                    }
                    float sum_c = 0;
                    for (int i = 4; i <= 12; i++)
                    {
                        if (read[i] != null)
                        {
                            wordDoc.Tables[2].Cell(2 + row, i + 3).Range.Text = read[i].ToString();
                            sum_c += (read[i].ToString() != "" ? float.Parse(read[i].ToString()) : 0);
                            sum_cells1[i - 4] += (read[i].ToString() != "" ? float.Parse(read[i].ToString()) : 0);
                        }
                    }
                    wordDoc.Tables[2].Cell(2 + row, 16).Range.Text = sum_c.ToString();
                }
                wordDoc.Tables[2].Rows.Add();
                row++;
            }
            read.Close();
            sum_sum_cells = 0;
            for (int i = 0; i < 9; i++)
            {
                wordDoc.Tables[2].Cell(2 + row, i + 7).Range.Text = sum_cells1[i].ToString();
                sum_sum_cells += sum_cells1[i];
            }
            wordDoc.Tables[2].Cell(2 + row, 16).Range.Text = sum_sum_cells.ToString();

            wordApp.Visible = false;
            wordDoc.Save();
            wordDoc.Close();
            wordApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordDoc);
            GC.Collect();
            conn.Close();
            }
            catch (Exception ex){ MessageBox.Show(ex.ToString()); }

        }

        private void серверToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 form4 = new Form4();
            form4.ShowDialog();
        }
    }
}
