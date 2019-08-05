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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();

            string path = System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.txt";
            using (StreamReader sr = new StreamReader(path))
            {
                label2.Text = "Коефіцієнт: " + sr.ReadToEnd();
            }
        }
        private void OK_Click(object sender, EventArgs e)
        {
            using (StreamWriter sw = new StreamWriter(System.IO.Directory.GetCurrentDirectory().ToString() + "\\" + "ITP.txt", false, System.Text.Encoding.Default))
            {
                sw.WriteLine(textBox1.Text);
                MessageBox.Show("Коефіцієнт: " + textBox1.Text + " доданий!");
                Close();
            }
        }
    }
}
