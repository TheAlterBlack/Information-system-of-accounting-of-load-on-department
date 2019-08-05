using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
namespace Diploma_Add_Py
{
    public partial class Form4 : Form
    {   public Form4()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = new StreamWriter(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "Server.ini", false, System.Text.Encoding.Default))
            {
                if (textBox1.Text != "" && textBox1.Text!=null)
                {
                    if (textBox2.Text != "" && textBox2.Text != null)
                    {
                        if (textBox3.Text != "" && textBox3.Text != null)
                        {
                            string connect = String.Format("server={0};user={1};database={2};password={3}", textBox1.Text, textBox2.Text,
                                textBox3.Text, textBox4.Text);
                            sw.WriteLine(connect);
                            MessageBox.Show("Інформація додана!");
                            Close();
                        }
                        else { MessageBox.Show("Введіть назву БД!"); }
                    }
                    else { MessageBox.Show("Введіть ім'я користувача!"); }
                }
                else { MessageBox.Show("Введіть назву серверу!"); }}}}}
