using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bogcha
{
    public partial class Form2 : Form
    {
        bool close = true;
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Mahsulot().Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            close = true;
            this.Close();
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Taom().Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Taqsimot().Show();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (close)
            {
                new Form1().Show();
            }
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Jadval().Show();
        }
    }
}
