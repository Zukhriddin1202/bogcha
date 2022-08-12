using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bogcha
{
    public partial class Taom : Form
    {
        private SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MTMDB"].ConnectionString);
        private int index = -1;
        int[] ID = null;
        bool close = true;
        public Taom()
        {
            InitializeComponent();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Form2().Show();
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
                cmd.CommandText = $"select id as'Id', nomi as 'Taom nomi' from Taom order by nomi desc";
                con.Open();
                data.Load(cmd.ExecuteReader());
                guna2DataGridView1.DataSource = data;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void clear()
        {
            nomi.Text = "";
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            try
            {
                nomi.Text = reg(nomi.Text);
                if (!nomi.Text.Equals(string.Empty) && checkedListBox1.CheckedItems.Count > 0)
                {
                    SqlCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    DataTable data = new DataTable();
                    cmd.CommandText = $"select * from Taom where nomi='{nomi.Text}'";
                    con.Open();
                    data.Load(cmd.ExecuteReader());
                    con.Close();
                    if (data.Rows.Count > 0)
                        MessageBox.Show("Bunday taom mavjud");
                    else
                    {
                        con.Open();
                        cmd.CommandText = $"insert into Taom(nomi) values('{nomi.Text}')";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = $"select id from Taom where nomi='{nomi.Text}'";
                        data = new DataTable();
                        data.Load(cmd.ExecuteReader());
                        string s = "";
                        int f = checkedListBox1.CheckedItems.Count;
                        for (int i = 0; i < f; i++)
                        {
                            s += $"({data.Rows[0][0].ToString()},{ID[checkedListBox1.CheckedIndices[i]]}),";
                        }
                        cmd.CommandText = $"insert into TaomMahsulot(taomId,mahsulotId) values" + s.Substring(0, s.Length - 1);
                        cmd.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("Taom saqlandi");
                        for (int i = 0; i < ID.Length; i++)
                            checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
                        clear();
                        view();

                    }
                }
                else
                {
                    MessageBox.Show("Maydonlarni to'ldiring");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void Taom_Load(object sender, EventArgs e)
        {
            view();
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                DataTable data = new DataTable();
                cmd.CommandText = $"select id,nomi from Mahsulot";
                con.Open();
                data.Load(cmd.ExecuteReader());
                ID = new int[data.Rows.Count];
                for(int i=0; i<data.Rows.Count; i++)
                {
                    ID[i] = int.Parse(data.Rows[i][0].ToString());
                    checkedListBox1.Items.Add(data.Rows[i][1].ToString());
                }
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
            {
                try
                {
                    nomi.Text = reg(nomi.Text);
                    if (!nomi.Text.Equals(string.Empty) && checkedListBox1.CheckedItems.Count > 0)
                    {
                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        DataTable data = new DataTable();
                        cmd.CommandText = $"select * from Taom where nomi='{nomi.Text}' and id!={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        con.Open();
                        data.Load(cmd.ExecuteReader());
                        con.Close();
                        if (data.Rows.Count > 0)
                            MessageBox.Show("Bunday taom mavjud");
                        else
                        {
                            con.Open();
                            cmd.CommandText = $"update Taom set nomi='{nomi.Text}' where id={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                            cmd.ExecuteNonQuery();
                            cmd.CommandText = $"delete from TaomMahsulot where taomId={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                            cmd.ExecuteNonQuery();
                            string s = "";
                            int f = checkedListBox1.CheckedItems.Count;
                            for (int i = 0; i < f; i++)
                            {
                                s += $"({guna2DataGridView1.Rows[index].Cells[0].Value.ToString()},{ID[checkedListBox1.CheckedIndices[i]]}),";
                            }
                            cmd.CommandText = $"insert into TaomMahsulot(taomId,mahsulotId) values" + s.Substring(0, s.Length - 1);
                            cmd.ExecuteNonQuery();
                            con.Close();
                            MessageBox.Show("Taom yangilandi");
                            for (int i = 0; i < ID.Length; i++)
                                checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);

                            clear();
                            view();

                        }
                    }
                    else
                    {
                        MessageBox.Show("Maydonlarni to'ldiring");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    con.Close();
                }
            }
            else MessageBox.Show("Yangilayotgan taomni tanlang");

        }

        private void guna2DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                index = e.RowIndex;
                if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
                {
                    for (int j = 0; j < ID.Length; j++)
                        checkedListBox1.SetItemChecked(j, false);
                    DataGridViewRow row = guna2DataGridView1.Rows[index];
                    nomi.Text = row.Cells[1].Value.ToString();
                    SqlCommand cmd = con.CreateCommand();
                    cmd.CommandType = CommandType.Text;
                    DataTable data = new DataTable();
                    cmd.CommandText = $"select * from TaomMahsulot where taomId={row.Cells[0].Value.ToString()}";
                    con.Open();
                    data.Load(cmd.ExecuteReader());
                    con.Close();
                    for(int i=0; i<data.Rows.Count; i++)
                    {
                        for(int j=0; j<ID.Length; j++)
                        {
                            if (data.Rows[i][2].ToString().Equals(ID[j].ToString())){
                                checkedListBox1.SetItemChecked(j, true);
                            }
                           
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            if (index != -1 && index != guna2DataGridView1.Rows.Count - 1)
            {
                try
                {
                    if (MessageBox.Show("Tanlangan taom o'chirmoqchimisiz?", "Ogohlantirish!!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {

                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        cmd.CommandText = $"delete from TaomMahsulot where taomId={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = $"delete from Taom where id={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        cmd.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("Taom o'chirildi");
                        clear();
                        for (int i = 0; i < ID.Length; i++)
                            checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);

                        guna2DataGridView1.Rows.RemoveAt(index);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    con.Close();
                }
            }
            else MessageBox.Show("O'chirayotgan taomni tanlang");

        }

        private void Taom_FormClosed(object sender, FormClosedEventArgs e)
        {

            if (close) new Form1().Show();
        }
    }
}
