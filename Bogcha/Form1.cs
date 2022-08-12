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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
           
                if (textBox2.PasswordChar == '*')
                {
                    textBox2.PasswordChar = '\0';
                    pictureBox4.BringToFront();
                }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text.Equals("admin") && textBox2.Text.Equals("1"))
            {
                this.Hide();
                new Form2().Show();
            }
            else
            {
                MessageBox.Show("Login yoki parol xato kiritildi");
                textBox1.Text = textBox2.Text = "";
            }
        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {

            if (textBox2.PasswordChar == '\0')
            {
                textBox2.PasswordChar = '*';
                pictureBox5.BringToFront();
            }
        }

       

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
    }
}
