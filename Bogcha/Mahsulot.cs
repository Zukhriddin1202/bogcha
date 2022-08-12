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
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Bogcha
{
    public partial class Mahsulot : Form
    {
        private SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MTMDB"].ConnectionString);
        private int index = -1;
        bool close = true;
        public Mahsulot()
        {
            InitializeComponent();
        }

        private String reg(String s)
        {
            Regex r = new Regex("'");
            return r.Replace(s, "`");
        }

        private void view()
        {
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                DataTable data = new DataTable();
                cmd.CommandText = $"select id, nomi as 'Nomi', chiqit as 'Chiqit', yog as 'Yog`', oqsil as 'Oqsil', uglevod as 'Uglevod', kkal as 'KKal',narx as 'Narx', birlik as 'Birligi' from Mahsulot order by nomi desc";
                con.Open();
                data.Load(cmd.ExecuteReader());
                guna2DataGridView1.DataSource = data;
                con.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void clear()
        {
            nomi.Text = chiqit.Text = yog.Text = oqsil.Text = uglevod.Text = kkal.Text =narx.Text= "";
            birlik.SelectedIndex = 0;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Form2().Show();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            nomi.Text = reg(nomi.Text);
            try
            {
                if(!nomi.Text.Equals(string.Empty) && !chiqit.Text.Equals(string.Empty) && !yog.Text.Equals(string.Empty) && !oqsil.Text.Equals(string.Empty) && !uglevod.Text.Equals(string.Empty) && !kkal.Text.Equals(string.Empty) && !narx.Text.Equals(string.Empty) && birlik.SelectedIndex!=-1)
                {
                    SqlCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    DataTable data = new DataTable();
                    cmd.CommandText = $"select * from Mahsulot where nomi ='{nomi.Text}'";
                    con.Open();
                    data.Load(cmd.ExecuteReader());
                    con.Close();
                    if (data.Rows.Count > 0)
                        MessageBox.Show("Bunday mahsulot mavjud");
                    else
                    {
                        cmd.CommandText = $"insert into Mahsulot(nomi,chiqit,yog,oqsil,uglevod,kkal,birlik,narx) values(" +
                            $"'{nomi.Text}',{double.Parse(chiqit.Text)},{double.Parse(yog.Text)},{double.Parse(oqsil.Text)},{double.Parse(uglevod.Text)},{double.Parse(kkal.Text)},'{birlik.SelectedItem.ToString()}',{narx.Text})";
                        con.Open(); 
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Mahsulot saqalandi");
                        con.Close();
                        clear();
                        view();
                    }
                    
                }
                else
                {
                    MessageBox.Show("Maydonnlarni to'ldiring");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void Mahsulot_Load(object sender, EventArgs e)
        {
            view();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {

            if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
            {
                nomi.Text = reg(nomi.Text);
                try
                {

                    if (!nomi.Text.Equals(string.Empty) && !chiqit.Text.Equals(string.Empty) && !yog.Text.Equals(string.Empty) && !oqsil.Text.Equals(string.Empty) && !uglevod.Text.Equals(string.Empty) && !kkal.Text.Equals(string.Empty) && !narx.Text.Equals(string.Empty) && birlik.SelectedIndex != -1)
                    {
                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        DataTable data = new DataTable();
                        cmd.CommandText = $"select * from Mahsulot where nomi ='{nomi.Text}' and id!={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        con.Open();
                        data.Load(cmd.ExecuteReader());
                        con.Close();
                        if (data.Rows.Count > 0)
                            MessageBox.Show("Bunday mahsulot mavjud");
                        else
                        {
                            cmd.CommandText = $"update  Mahsulot set nomi='{nomi.Text}',chiqit={double.Parse(chiqit.Text)},yog={double.Parse(yog.Text)},oqsil={double.Parse(oqsil.Text)},uglevod={double.Parse(uglevod.Text)},kkal={double.Parse(kkal.Text)},birlik='{birlik.SelectedItem.ToString()}', narx={double.Parse(narx.Text)} where id={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                            con.Open();
                            cmd.ExecuteNonQuery();
                            con.Close();
                            MessageBox.Show("Mahsulot yangilandi");
                            index = -1;
                            view();
                            clear();
                        }
                    }
                    else
                    {
                        MessageBox.Show("Maydonnlarni to'ldiring");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    con.Close();
                }
            }
            else
                MessageBox.Show("Yangilayotgan mahsulotni tanlang");
        }

        private void guna2DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                index = e.RowIndex;
                if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
                {
                    DataGridViewRow row = guna2DataGridView1.Rows[index];
                    nomi.Text = row.Cells[1].Value.ToString();
                    chiqit.Text = row.Cells[2].Value.ToString();
                    yog.Text = row.Cells[3].Value.ToString();
                    oqsil.Text = row.Cells[4].Value.ToString();
                    uglevod.Text = row.Cells[5].Value.ToString();
                    kkal.Text = row.Cells[6].Value.ToString();
                    narx.Text = row.Cells[7].Value.ToString();
                    birlik.SelectedItem =row.Cells[8].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            clear();
            if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
            {
                try
                {
                    if (MessageBox.Show("Tanlangan mahsulotni o'chirmoqchimisiz?", "Ogohlantirish!!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        DataTable data = new DataTable();
                        cmd.CommandText = $"select * from TaomMahsulot where mahsulotId={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        con.Open(); 
                        data.Load(cmd.ExecuteReader());
                        con.Close();
                        if (data.Rows.Count > 0)
                        {
                            MessageBox.Show("Siz bu mahsulotni o'chiraolmaysiz. Sababi bu mahsulot qaysidir taom uchun ishlatilgan. Agar o'chirmoqchi bo'lsangiz dasavval bu mahsulot qatnashgan taomlarni o'chirishga to'g'ri keladi!!");
                        }
                        else
                        {
                            cmd.CommandText = $"delete from Mahsulot where id={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                            con.Open();
                            cmd.ExecuteNonQuery();
                            con.Close();
                            MessageBox.Show("Mahsulot o'chirildi");
                            guna2DataGridView1.Rows.RemoveAt(index);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    con.Close();
                }
            }
            else MessageBox.Show("O'chirayotgan mahsulotini tanlang");

        }

        private void Mahsulot_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (close) new Form1().Show();
          
        }
    }
}
