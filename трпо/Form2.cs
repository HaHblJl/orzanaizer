using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace трпо
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        public Form2(Form1 f)
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Text == "") return;
            string workspace = textBox1.Text;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if ((dataGridView1[6, i].Value.ToString() != workspace) | (Convert.ToDouble(dataGridView1[4, i].Value) != (Convert.ToDouble(dataGridView1[5, i].Value) / 2.0)))
                {
                    dataGridView1.Rows.RemoveAt(i--);
                }
            }
        }

    }
}
