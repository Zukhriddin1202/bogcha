using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;
using System.Windows.Forms;

namespace Bogcha
{
    public partial class Taqsimot : Form
    {
        private SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MTMDB"].ConnectionString);
        private int index = -1;
        int[] ID = null;
        int[] taomMahsulotId = null;
        string[] masalih = null;
        int k = 0;
        string query = "";
        int idmasalih;
        double jamiMiqdor = 0;
        bool close = true;
        public Taqsimot()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (k < taomMahsulotId.Length - 1)
                {
                    if (!miqdor.Text.Equals(string.Empty))
                    {
                        jamiMiqdor += double.Parse(miqdor.Text);
                        query += $"({taomMahsulotId[k]},{double.Parse(miqdor.Text)},{idmasalih}),";
                        k++;
                        nom.Text = masalih[k];
                        if (k + 1 == taomMahsulotId.Length)
                            guna2Button1.Text = "Saqlash";
                        miqdor.Text = "";
                    }
                    else
                    {
                        MessageBox.Show("Miqdorni kiriting");
                    }
                }
                else
                {
                    if (!miqdor.Text.Equals(string.Empty))
                    {
                        jamiMiqdor += double.Parse(miqdor.Text);
                        query += $"({taomMahsulotId[k]},{double.Parse(miqdor.Text)},{idmasalih}),";
                        k = 0;
                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        cmd.CommandText = $"insert into TaomTaqsimot(taomMahsulot,miqdor,taqsimotId) values {query.Substring(0, query.Length - 1)}";
                        con.Open();
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = $"update Taqsimot set miqdor={jamiMiqdor} where id={idmasalih}";
                        cmd.ExecuteNonQuery();
                        con.Close();
                        MessageBox.Show("Saqlandi");
                        view();
                        guna2Button1.Text = "Keyingisi";
                        miqdor.Text = "";
                        enebl();
                    }
                    else
                    {
                        MessageBox.Show("Miqdorni kiriting");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Form2().Show();
        }

        private void enebl()
        {
            radioButton1.Enabled = !radioButton1.Enabled;
            radioButton2.Enabled = !radioButton2.Enabled;
            kun.Enabled = !kun.Enabled;
            tur.Enabled = !tur.Enabled;
            taom.Enabled = !taom.Enabled;
            panel1.Visible = !panel1.Visible;
            guna2DataGridView1.Enabled = !guna2DataGridView1.Enabled;
            panel2.Visible = !panel2.Visible;
        }
        private void taom_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                jamiMiqdor = 0;
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                DataTable data = new DataTable();
                string s = radioButton1.Checked ? radioButton1.Text : radioButton2.Text;
                cmd.CommandText = $"select ta.nomi from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taommahsulot inner join Taom ta on ta.id=tm.taomId  where yosh='{s}' and kun='{kun.SelectedItem}' and tur='{tur.SelectedItem}' and ta.nomi='{taom.SelectedItem}'";
                con.Open();
                data.Load(cmd.ExecuteReader());
                con.Close();
                if (data.Rows.Count > 0)
                {
                    MessageBox.Show("Bunday taom bu vaqt uchun allaqachon ajratilgan");
                }
                else
                {
                    MessageBox.Show("Siz tanlagan taomning masaliqlarini birin ketin miqdorini kiriting.");
                    cmd.CommandText = $"select tm.id , m.nomi from Mahsulot m inner join TaomMahsulot tm on m.id=tm.mahsulotId where tm.taomId={ID[taom.SelectedIndex]}";
                    con.Open();
                    data = new DataTable();
                    data.Load(cmd.ExecuteReader());
                    taomMahsulotId = new int[data.Rows.Count];
                    masalih = new string[data.Rows.Count];
                    for (int i = 0; i < data.Rows.Count; i++)
                    {
                        masalih[i] = data.Rows[i][1].ToString();
                        taomMahsulotId[i] = int.Parse(data.Rows[i][0].ToString());
                    }
                    cmd.CommandText = $"insert into Taqsimot(yosh,kun,tur,miqdor) values('{s}','{kun.SelectedItem}','{tur.SelectedItem}',-1)";
                    cmd.ExecuteNonQuery();
                    cmd.CommandText = $"select id from Taqsimot where miqdor=-1";
                    data = new DataTable();
                    data.Load(cmd.ExecuteReader());
                    idmasalih = int.Parse(data.Rows[0][0].ToString());
                    con.Close();
                    nom.Text = masalih[0];
                    if (masalih.Length == 1)
                        guna2Button1.Text = "Saqlash";
                    enebl();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void view()
        {
            try
            {

                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                DataTable data = new DataTable();
                string s = radioButton1.Checked ? radioButton1.Text : radioButton2.Text;
                cmd.CommandText = $"select distinct t.id,yosh as'Yosh', kun as 'Kun', tur as 'Taom turi', ta.nomi as'Taom',t.miqdor as 'Taom miqdori(gr)' from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taommahsulot inner join Taom ta on ta.id=tm.taomId  where yosh='{s}' and kun='{kun.SelectedItem}' and tur='{tur.SelectedItem}'";
                con.Open();
                data.Load(cmd.ExecuteReader());
                data.Columns.Add("Mahsulotlar va miqdorlari");
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    cmd.CommandText = $"select distinct m.nomi,tt.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taommahsulot inner join Mahsulot m on m.id=tm.mahsulotId inner join Taom ta on ta.id=tm.taomId  where yosh='{s}' and kun='{kun.SelectedItem}' and tur='{tur.SelectedItem}' and ta.nomi='{data.Rows[i][4].ToString()}'";
                    DataTable mah = new DataTable();
                    mah.Load(cmd.ExecuteReader());
                    string z = "";
                    for (int j = 0; j < mah.Rows.Count; j++)
                    {
                        z += (mah.Rows[j][0] + "-" + mah.Rows[j][1] + ",\n");
                    }
                    data.Rows[i][6] = z.Substring(0, z.Length - 2);
                }
                guna2DataGridView1.DataSource = data;
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }

        }
        private void Taqsimot_Load(object sender, EventArgs e)
        {
            try
            {
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.Text;
                DataTable data = new DataTable();
                cmd.CommandText = $"select id ,nomi from Taom";
                con.Open();
                data.Load(cmd.ExecuteReader());
                ID = new int[data.Rows.Count];
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    taom.Items.Add(data.Rows[i][1].ToString());
                    ID[i] = int.Parse(data.Rows[i][0].ToString());
                }
                con.Close();
                view();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            view();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            view();
        }

        private void kun_SelectedIndexChanged(object sender, EventArgs e)
        {
            view();
        }

        private void tur_SelectedIndexChanged(object sender, EventArgs e)
        {
            view();
        }

        private void guna2DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            index = e.RowIndex;
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (index != -1 && index != guna2DataGridView1.RowCount - 1)
                {
                    if (MessageBox.Show("Belgilangan taom taqsimotini o'chirasizmi?", "Ogohlantirish!!!", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        SqlCommand cmd = con.CreateCommand();
                        cmd.CommandType = CommandType.Text;
                        con.Open();
                        cmd.CommandText = $"delete from TaomTaqsimot where taqsimotId={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        cmd.ExecuteNonQuery();
                        cmd.CommandText = $"delete from Taqsimot where id={guna2DataGridView1.Rows[index].Cells[0].Value.ToString()}";
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("O'chirildi");
                        con.Close();
                        guna2DataGridView1.Rows.RemoveAt(index);
                    }
                }
                else
                {
                    MessageBox.Show("O'chiriladigan satrni tanlang!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }

        }

        private void Taqsimot_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (close) new Form1().Show();
        }
    }
}
