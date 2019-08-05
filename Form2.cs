using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
namespace Diploma_Add_Py
{
    public partial class Form2 : Form
    {
        public Form2(int x, int y)
        {
            InitializeComponent();
            this.Location = new Point(x-277, y-175);

            string professor = comboBox1.Text;
            
            if (professor == "")
            {
                string conn_s;
                using (StreamReader sr = new StreamReader(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini"))
                {
                    conn_s = sr.ReadToEnd();
                }
                MySqlConnection conn = new MySqlConnection(conn_s);
                conn.Open();

                string query = "SELECT count(name) FROM professor";
                MySqlCommand com = new MySqlCommand(query, conn);

                int count;
                bool a = Int32.TryParse(com.ExecuteScalar().ToString(), out count);
                if (count != 0)
                {
                    string[] prof = new string[count];
                    query = "SELECT name, surname, patronymic FROM professor";
                    com = new MySqlCommand(query, conn);
                    MySqlDataReader read = com.ExecuteReader();
                    int i = 0;
                    while (read.Read())
                    {
                        prof[i] = read[1].ToString() +" "+ read[0].ToString() +" "+ read[2].ToString();
                        i++;
                    }
                    read.Close();
                    for (i = 0; i < count; ++i) {
                        comboBox1.Items.Add(prof[i]);
                    }
                }
                else { MessageBox.Show("Викладачі не додані!\nДодайте викладчів і поверніться до розподілу навантаження"); }
            }
        }
        private void OK_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "")
            {
                //Изменяем dataTransfer
                DataTransfer.data[0] = comboBox1.Text;
                Dispose();
            }
            else { MessageBox.Show("Оберіть викладача!"); }
        }
    }
}
