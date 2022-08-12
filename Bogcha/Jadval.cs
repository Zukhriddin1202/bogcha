using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Bogcha
{
    public partial class Jadval : Form
    {
        private SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["MTMDB"].ConnectionString);
        bool close = true;
        public Jadval()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

            if (!bola.Text.Equals(string.Empty))
            {
                using (SaveFileDialog saveFile = new SaveFileDialog() { Filter = "Excel|*.xlsx" })
                {
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        var fileInfo = new FileInfo(saveFile.FileName);
                        using (ExcelEngine excelEngine = new ExcelEngine())
                        {
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Xlsx;
                            
                            int zet = 1;
                            DateTime moment = DateTime.Now;
                            int f = 0;
                            List<int> ishkuni = new List<int>();
                            for (int kun = 1; kun <= DateTime.DaysInMonth(moment.Year, moment.Month); kun++)
                            {
                                DateTime newDate = new DateTime(moment.Year, moment.Month, kun);
                                if (newDate.DayOfWeek != DayOfWeek.Sunday && newDate.DayOfWeek != DayOfWeek.Saturday)
                                {
                                    f++;
                                    ishkuni.Add(kun);
                                }
                            }
                            ishkuni.Sort();
                            int[] ishkunlari = ishkuni.ToArray();

                            IWorkbook workbook = application.Workbooks.Create(f+3);
                            for (int kun = 1; kun <= DateTime.DaysInMonth(moment.Year, moment.Month); kun++)
                            {
                                DateTime newDate = new DateTime(moment.Year, moment.Month, kun);
                                if (newDate.DayOfWeek != DayOfWeek.Sunday && newDate.DayOfWeek != DayOfWeek.Saturday)
                                {
                                    IWorksheet worksheet = workbook.Worksheets[zet - 1];
                                   

                                    worksheet.Name = kun + "-kun";
                                    string s = radioButton1.Checked ? radioButton1.Text : radioButton2.Text;
                                    int day = zet % 10==0?10: zet % 10;
                                    int bolasoni = int.Parse(bola.Text);
                                    SqlCommand cmd = con.CreateCommand();
                                    cmd.CommandType = CommandType.Text;
                                    con.Open();
                                    DataTable data = new DataTable();
                                    cmd.CommandText = $" select  count(*) from Taqsimot t  where kun='{day + "-kun"}' and yosh='{s}'";
                                    data.Load(cmd.ExecuteReader());
                                    int n = int.Parse(data.Rows[0][0].ToString()) + 1;
                                    cmd.CommandText = "select nomi,id,narx,birlik from Mahsulot";
                                    data = new DataTable();
                                    data.Load(cmd.ExecuteReader());
                                    int m = data.Rows.Count + 3;
                                    String[] mahsulot = new String[data.Rows.Count];
                                    int[] mahsulotId = new int[data.Rows.Count];
                                    double[] mahsulotNarx = new double[data.Rows.Count];
                                    double[] mahsulotOylikMiqdor = new double[data.Rows.Count];
                                    string[] mahsulotBirlik = new string[data.Rows.Count];
                                    for (int i = 0; i < data.Rows.Count; i++)
                                    {
                                        mahsulot[i] = data.Rows[i][0].ToString();
                                        mahsulotId[i] = int.Parse(data.Rows[i][1].ToString());
                                        mahsulotNarx[i] = double.Parse(data.Rows[i][2].ToString());
                                        mahsulotBirlik[i] = data.Rows[i][3].ToString();
                        }
                                    cmd.CommandText = $"select distinct ta.nomi,t.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot inner join Taom ta on ta.id=tm.taomId where kun='{day + "-kun"}' and yosh='{s}' and tur='Ertalabki nonushta'";
                                    data = new DataTable();
                                    data.Load(cmd.ExecuteReader());
                                    String[] toamNomi1 = new String[data.Rows.Count];
                                    double[] toamMiqdori1 = new double[data.Rows.Count];
                                    for (int i = 0; i < data.Rows.Count; i++)
                                    {
                                        toamNomi1[i] = data.Rows[i][0].ToString();
                                        toamMiqdori1[i] = double.Parse(data.Rows[i][1].ToString());
                                    }
                                    cmd.CommandText = $"select distinct ta.nomi,t.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot inner join Taom ta on ta.id=tm.taomId where kun='{day + "-kun"}' and yosh='{s}' and tur='Ikkinchi nonushta'";
                                    data = new DataTable();
                                    data.Load(cmd.ExecuteReader());
                                    String[] toamNomi2 = new String[data.Rows.Count];
                                    double[] toamMiqdori2 = new double[data.Rows.Count];
                                    for (int i = 0; i < data.Rows.Count; i++)
                                    {
                                        toamNomi2[i] = data.Rows[i][0].ToString();
                                        toamMiqdori2[i] = double.Parse(data.Rows[i][1].ToString());
                                    }
                                    cmd.CommandText = $"select distinct ta.nomi,t.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot inner join Taom ta on ta.id=tm.taomId where kun='{day + "-kun"}' and yosh='{s}' and tur='Tushlik'";
                                    data = new DataTable();
                                    data.Load(cmd.ExecuteReader());
                                    String[] toamNomi3 = new String[data.Rows.Count];
                                    double[] toamMiqdori3 = new double[data.Rows.Count];
                                    for (int i = 0; i < data.Rows.Count; i++)
                                    {
                                        toamNomi3[i] = data.Rows[i][0].ToString();
                                        toamMiqdori3[i] = double.Parse(data.Rows[i][1].ToString());
                                    }
                                    cmd.CommandText = $"select distinct ta.nomi,t.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot inner join Taom ta on ta.id=tm.taomId where kun='{day + "-kun"}' and yosh='{s}' and tur='Ikkinchi tushlik'";
                                    data = new DataTable();
                                    data.Load(cmd.ExecuteReader());
                                    String[] toamNomi4 = new String[data.Rows.Count];
                                    double[] toamMiqdori4 = new double[data.Rows.Count];
                                    for (int i = 0; i < data.Rows.Count; i++)
                                    {
                                        toamNomi4[i] = data.Rows[i][0].ToString();
                                        toamMiqdori4[i] = double.Parse(data.Rows[i][1].ToString());
                                    }
                                    void styl(IWorksheet worksheet1, int i, int j)
                                    {
                                        worksheet1.Range[i, j].CellStyle.Font.FontName = "Times New Roman";
                                        worksheet1.Range[i, j].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                        worksheet1.Range[i, j].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                                        worksheet1.Range[i, j].CellStyle.Font.Size = 12;
                                    }
                                    String[] vaqt = { "Ertalabki nonushta", "Ikkinchi nonushta", "Tushlik", "Ikkinchi tushlik" };
                                    for (int i = 1; i <= n; i++)
                                    {
                                        if (i == 1)
                                        {
                                            for (int j = 1; j <= m; j++)
                                            {
                                                styl(worksheet, i, j);
                                                worksheet.Range[i, j].CellStyle.Font.Bold = true;
                                                if (j == 2)
                                                {
                                                    worksheet.Range[i, j].ColumnWidth = 26;
                                                    worksheet.Range[i, j].Text = "Ovqatning nomlanishi";
                                                }
                                                else if (j == 3)
                                                {
                                                    worksheet.Range[i, j].Text = "1-nafar tarbiyalanuvchi\nuchun ovqat hajmi"; ;
                                                    worksheet.Range[i, j].CellStyle.Rotation = 90;
                                                }
                                                else if (j > 3)
                                                {
                                                    worksheet.Range[i, j].Text = mahsulot[j - 4];
                                                    worksheet.Range[i, j].CellStyle.Rotation = 90;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            String[] foodNomi = null;
                                            double[] foodMiqdor = null;
                                            String time = "";
                                            int l = 0;
                                            if (i >= 2 && i <= toamNomi1.Length + 1)
                                            {
                                                l = 2;
                                                time = vaqt[0];
                                                foodNomi = toamNomi1;
                                                foodMiqdor = toamMiqdori1;
                                            }
                                            else if (i >= 2 + toamNomi1.Length && i <= 1 + toamNomi1.Length + toamNomi2.Length)
                                            {
                                                l = 2 + toamNomi1.Length;
                                                time = vaqt[1];
                                                foodNomi = toamNomi2;
                                                foodMiqdor = toamMiqdori2;
                                            }
                                            else if (i >= 2 + toamNomi2.Length + toamNomi1.Length && i <= 1 + toamNomi3.Length + toamNomi1.Length + toamNomi2.Length)
                                            {
                                                l = 2 + toamNomi2.Length + toamNomi1.Length;
                                                time = vaqt[2];
                                                foodNomi = toamNomi3;
                                                foodMiqdor = toamMiqdori3;
                                            }
                                            else if (i >= 2 + toamNomi2.Length + toamNomi1.Length + toamNomi3.Length && i <= toamNomi4.Length + 1 + toamNomi3.Length + toamNomi1.Length + toamNomi2.Length)
                                            {
                                                l = 2 + toamNomi2.Length + toamNomi1.Length + toamNomi3.Length;
                                                time = vaqt[3];
                                                foodNomi = toamNomi4;
                                                foodMiqdor = toamMiqdori4;
                                            }
                                            for (int j = 1; j <= m; j++)
                                            {
                                                styl(worksheet, i, j);
                                                if (j == 1)
                                                {
                                                    if (i == l)
                                                    {
                                                        worksheet.Range[i, j].Text = time;
                                                        if (worksheet.Range[i, j].Text.Length * 6 > 36 * foodNomi.Length)
                                                        {

                                                            for (int k = i; k < i + foodNomi.Length; k++)
                                                                worksheet.Range[k, j].RowHeight = (worksheet.Range[i, j].Text.Length * 6) / foodNomi.Length;
                                                        }
                                                        else
                                                            for (int k = i; k < i + foodNomi.Length; k++)
                                                                worksheet.Range[k, j].RowHeight = 36;


                                                        worksheet.Range[i, j, i + foodNomi.Length - 1, j].Merge();
                                                        worksheet.Range[i, j].CellStyle.Rotation = 90;
                                                    }
                                                }

                                                else if (j == 2)
                                                    worksheet.Range[i, j].Text = foodNomi[i - l];
                                                else if (j == 3)
                                                    worksheet.Range[i, j].Number = foodMiqdor[i - l];
                                                else if (j > 3)
                                                {
                                                    cmd.CommandText = $"select distinct tt.miqdor from TaomTaqsimot tt inner join TaomMahsulot tm on tm.id=tt.taomMahsulot inner join Mahsulot m on m.id=tm.mahsulotId inner join Taom ta on ta.id=tm.taomId inner join Taqsimot t on t.id=tt.taqsimotId where m.nomi='{mahsulot[j - 4]}' and ta.nomi='{foodNomi[i - l]}' and t.tur='{time}' and t.kun='{day + "-kun"}' and yosh='{s}'";
                                                    data = new DataTable();
                                                    data.Load(cmd.ExecuteReader());
                                                    if (data.Rows.Count > 0)
                                                        worksheet.Range[i, j].Number = double.Parse(data.Rows[0][0].ToString());
                                                    else
                                                        worksheet.Range[i, j].Text="/";
                                                }
                                            }
                                        }
                                    }
                                    string[] soz = { $"Bir nafar {s} tarbiyalanuvchi uchun", "Jami mahsulot miqdori", "Mahsulot narxi", "Sarflangan mablag' miqdori" };
                                    for (int i = 1; i < 5; i++)
                                    {
                                        for (int j = 1; j <= m; j++)
                                        {
                                            styl(worksheet, i + n, j);
                                            worksheet.Range[i + n, j].RowHeight = 36;
                                            if (j == 1)
                                            {
                                                worksheet.Range[i + n, j, i + n, j + 2].Merge();
                                                worksheet.Range[i + n, j].Text = soz[i - 1];
                                            }
                                            else if (j > 3)
                                            {
                                                if (i == 1)
                                                {
                                                    cmd.CommandText = $"select distinct t.id,tt.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot where tm.mahsulotId={mahsulotId[j - 4]} and yosh='{s}' and kun='{day + "-kun"}'";
                                                    data = new DataTable();
                                                    data.Load(cmd.ExecuteReader());
                                                    double jami = 0;
                                                    for (int u = 0; u < data.Rows.Count; u++)
                                                    {
                                                        jami += double.Parse(data.Rows[u][1].ToString());
                                                    }
                                                    worksheet.Range[i + n, j].Number = jami;
                                                    worksheet.Range[i + n + 1, j].Number = jami * bolasoni / 1000;
                                                    worksheet.Range[i + n + 2, j].Number = mahsulotNarx[j - 4];
                                                    worksheet.Range[i + n + 3, j].Number = (jami * bolasoni / 1000) * mahsulotNarx[j - 4];
                                                }

                                            }
                                        }

                                    }
                                    worksheet.Range.BorderInside(ExcelLineStyle.Thin, Color.Black);
                                    worksheet.Range.BorderAround(ExcelLineStyle.Thin, Color.Black);
                                    zet++;
                                    if (kun == DateTime.DaysInMonth(moment.Year, moment.Month))
                                    {

                                       //narx
                                        worksheet = workbook.Worksheets[zet - 1];
                                        worksheet.Name = "Narx";
                                        double sumNarx = 0;
                                        for (int i = 1; i <= mahsulot.Length + 1; i++)
                                        {
                                            if (i == 1)
                                            {
                                                styl(worksheet, i, 1);
                                                worksheet.Range[i, 1].Text = "№";
                                                styl(worksheet, i, 2);
                                                worksheet.Range[i, 2].Text = "Mahsulot nomlari";
                                                styl(worksheet, i, 3);
                                                worksheet.Range[i, 3].Text = "Narxi";
                                                worksheet.Range[i, 1].RowHeight = 36;
                                                worksheet.Range[i, 2].ColumnWidth = worksheet.Range[i, 2].Text.Length * 3;
                                            }
                                            else
                                            {
                                                styl(worksheet, i, 1);
                                                worksheet.Range[i, 1].Text = (i - 1).ToString();
                                                styl(worksheet, i, 2);
                                                worksheet.Range[i, 2].Text = mahsulot[i - 2];
                                                styl(worksheet, i, 3);
                                                worksheet.Range[i, 3].Number = mahsulotNarx[i - 2];
                                                sumNarx += mahsulotNarx[i - 2];
                                            }
                                        }
                                        styl(worksheet, mahsulot.Length + 2, 1);
                                        worksheet.Range[mahsulot.Length + 2, 1, mahsulot.Length + 2, 2].Merge();
                                        worksheet.Range[mahsulot.Length + 2, 1].Text = "Jami";
                                        styl(worksheet, mahsulot.Length + 2, 3);
                                        worksheet.Range[mahsulot.Length + 2, 3].Number = sumNarx;
                                        worksheet.Range.BorderInside(ExcelLineStyle.Thin, Color.Black);
                                        worksheet.Range.BorderAround(ExcelLineStyle.Thin, Color.Black);

                                        //Jami
                                        string oyNomi = "";
                                        switch (moment.Month)
                                        {
                                            case 1:
                                                oyNomi = "yanvar";
                                                break;
                                            case 2:
                                                oyNomi = "fevral";
                                                break;
                                            case 3:
                                                oyNomi = "mart";
                                                break;
                                            case 4:
                                                oyNomi = "aprel";
                                                break;
                                            case 5:
                                                oyNomi = "may";
                                                break;
                                            case 6:
                                                oyNomi = "iyun";
                                                break;
                                            case 7:
                                                oyNomi = "iyul";
                                                break;
                                            case 8:
                                                oyNomi = "avgust";
                                                break;
                                            case 9:
                                                oyNomi = "sentyabr";
                                                break;
                                            case 10:
                                                oyNomi = "oktyabr";
                                                break;
                                            case 11:
                                                oyNomi = "noyabr";
                                                break;
                                            case 12:
                                                oyNomi = "dekabr";
                                                break;
                                        }
                                        worksheet = workbook.Worksheets[zet];
                                        worksheet.Name = s + " hisoboti";

                                        styl(worksheet, 1, 1);
                                        worksheet.Range[1, 1].CellStyle.Font.Size = 16;
                                        worksheet.Range[1, 1, 1, 32].Merge();
                                        worksheet.Range[1, 1].Text = "Sharof Rashidov tuman MTBga qarashli _______________ nomli davlat xususiy sherikchilik asosidagi oilaviy nodavlat maktabgacha ta'lim tashkilotining " + moment.Year + " -yil " + oyNomi + " oyi oziq-ovqat xisoboti";
                                        styl(worksheet, 1, 35);
                                        worksheet.Range[1, 34, 1, 37].Merge();
                                        worksheet.Range[1, 34].CellStyle.Font.Size = 13;
                                        worksheet.Range[1, 34].Text = "Tasdiqlayman.N.Toyloqova Yatt\nElmurod Shod DXSHАNОМТТ.";
                                        worksheet.Range[1, 34].RowHeight = 60;
                                        for (int i = 1; i <= 12 + f; i++)
                                            styl(worksheet, 4, i);
                                        worksheet.Range[4, 1, 6, 1].Merge();
                                        worksheet.Range[4, 1].Text = "№";
                                        worksheet.Range[4, 2, 6, 2].Merge();
                                        worksheet.Range[4, 2].Text = "Mahsulot nomlari";
                                        worksheet.Range[4, 2].ColumnWidth = "Mahsulot nomlari".Length * 2;
                                        worksheet.Range[4, 3, 6, 3].Merge();
                                        worksheet.Range[4, 3].Text = "O'lchov birligi";
                                        worksheet.Range[4, 3].ColumnWidth = "O'lchov birligi".Length;
                                        worksheet.Range[4, 4, 6, 4].Merge();
                                        worksheet.Range[4, 4].Text = "Narxi";
                                        worksheet.Range[4, 4].ColumnWidth = "O'lchov birligi".Length;
                                        worksheet.Range[4, 5, 5, 6].Merge();
                                        worksheet.Range[4, 5].Text = "Oy boshiga qoldiq";
                                        styl(worksheet, 6, 5);
                                        worksheet.Range[6, 5].Text = "Miqdori";
                                        worksheet.Range[6, 5].ColumnWidth = "Miqdori".Length * 2;
                                        styl(worksheet, 6, 6);
                                        worksheet.Range[6, 6].Text = "Summasi";
                                        worksheet.Range[6, 6].ColumnWidth = "Miqdori".Length * 2;
                                        worksheet.Range[4, 7, 5, 8].Merge();
                                        worksheet.Range[4, 7].Text = "Jami kirim";
                                        styl(worksheet, 6, 7);
                                        worksheet.Range[6, 7].Text = "Miqdori";

                                        worksheet.Range[6, 7].ColumnWidth = "Miqdori".Length * 2;
                                        styl(worksheet, 6, 8);
                                        worksheet.Range[6, 8].Text = "Summasi";

                                        worksheet.Range[6, 8].ColumnWidth = "Miqdori".Length * 2;

                                        worksheet.Range[4, 9, 4, 9 + f - 1].Merge();
                                        worksheet.Range[4, 9].Text = "Sanasi";

                                        worksheet.Range[6, 5].RowHeight = 36;
                                        for (int i = 9; i < 9 + f; i++)
                                        {
                                            styl(worksheet, 5, i);
                                            styl(worksheet, 6, i);
                                            worksheet.Range[5, i].Text = moment.Month.ToString();
                                            worksheet.Range[6, i].Number = ishkunlari[i - 9];
                                        }

                                        worksheet.Range[4, 9 + f, 5, 9 + f + 1].Merge();
                                        worksheet.Range[4, 9 + f].Text = "Jami chiqim";
                                        styl(worksheet, 6, 9 + f);
                                        worksheet.Range[6, 9 + f].Text = "Miqdori";
                                        worksheet.Range[6, 9 + f].ColumnWidth = "Miqdori".Length * 2;
                                        styl(worksheet, 6, 9 + f + 1);
                                        worksheet.Range[6, 9 + f + 1].Text = "Summasi";
                                        worksheet.Range[6, 9 + f + 1].ColumnWidth = "Miqdori".Length * 2;
                                        worksheet.Range[4, 9 + f + 2, 5, 9 + f + 3].Merge();
                                        worksheet.Range[4, 9 + f + 2].Text = "Oy oxiridagi qoldiq";
                                        styl(worksheet, 6, 9 + f + 2);
                                        worksheet.Range[6, 9 + f + 2].Text = "Miqdori";
                                        worksheet.Range[6, 9 + f + 2].ColumnWidth = "Miqdori".Length * 2;
                                        styl(worksheet, 6, 9 + f + 3);
                                        worksheet.Range[6, 9 + f + 3].Text = "Summasi";
                                        worksheet.Range[6, 9 + f + 3].ColumnWidth = "Miqdori".Length * 2;

                                        for (int i = 7; i < mahsulot.Length + 7; i++)
                                        {
                                            for (int j = 1; j <= 12 + f; j++)
                                            {
                                                styl(worksheet, i, j);
                                            }
                                            worksheet.Range[i, 1].Text = (i - 6).ToString();
                                            worksheet.Range[i, 2].Text = mahsulot[i - 7];
                                            worksheet.Range[i, 3].Text = mahsulotBirlik[i - 7];
                                            worksheet.Range[i, 4].Number = mahsulotNarx[i - 7];

                                            worksheet.Range[i, 5].Number = 0;
                                            worksheet.Range[i, 6].Number = 0;


                                            double jami1 = 0, jami2 = 0;

                                            int kunmi = 1;
                                            for (int q = 9; q < 9 + f; q++)
                                            {
                                                kunmi = kunmi % 10 == 0 ? 10 : kunmi % 10;
                                                cmd.CommandText = $"select distinct t.id,tt.miqdor from Taqsimot t inner join TaomTaqsimot tt on tt.taqsimotId=t.id inner join TaomMahsulot tm on tm.id=tt.taomMahsulot where tm.mahsulotId={mahsulotId[i - 7]} and yosh='{s}' and kun='{kunmi + "-kun"}'";
                                                data = new DataTable();
                                                data.Load(cmd.ExecuteReader());
                                                jami1 = 0;
                                                for (int u = 0; u < data.Rows.Count; u++)
                                                {
                                                    jami1 += double.Parse(data.Rows[u][1].ToString());
                                                }
                                                jami2 += jami1 / 1000;
                                                worksheet.Range[i, q].Number = jami1 / 1000;
                                                kunmi++;
                                            }

                                            mahsulotOylikMiqdor[i-7] = jami2;
                                            worksheet.Range[i, 7].Number = jami2;
                                            worksheet.Range[i, 8].Number = worksheet.Range[i, 7].Number * worksheet.Range[i, 4].Number;
                                            worksheet.Range[i, 9 + f].Number = Math.Round(worksheet.Range[i, 7].Number);
                                            worksheet.Range[i, 9 + f + 1].Number = Math.Round(worksheet.Range[i, 7].Number) * worksheet.Range[i, 4].Number;

                                            worksheet.Range[i, 9 + f + 2].Number = Math.Abs(worksheet.Range[i, 9 + f].Number - worksheet.Range[i, 7].Number);

                                            worksheet.Range[i, 9 + f + 3].Number = worksheet.Range[i, 9 + f + 2].Number * worksheet.Range[i, 4].Number;
                                        }
                                        worksheet.Range[mahsulot.Length + 7, 1, mahsulot.Length + 7, 3].Merge();
                                        styl(worksheet, mahsulot.Length + 7, 1);
                                        worksheet.Range[mahsulot.Length + 7, 1].Text = "Jami";
                                        char belgi = 'D',perbelgi=' ';
                                        for (int q = 4; q <=12 + f; q++)
                                        {
                                            styl(worksheet, mahsulot.Length + 7, q);
                                            worksheet.Range[mahsulot.Length + 7, q].Formula = $"=SUM({perbelgi}{belgi}7:{perbelgi}{belgi}{mahsulot.Length+6})";
                                            if (belgi == 'Z')
                                            {
                                                perbelgi = 'A';
                                                belgi = 'A';
                                            }
                                            else belgi=(char)(belgi+ 1);
                                        }

                                        worksheet.Range[4,1, mahsulot.Length + 7, 12 + f].BorderInside(ExcelLineStyle.Thin, Color.Black);
                                        worksheet.Range[4, 1, mahsulot.Length + 7, 12 + f].BorderAround(ExcelLineStyle.Thin, Color.Black);
                                        worksheet.Range[mahsulot.Length + 10, 2, mahsulot.Length + 10, 4].Merge();
                                        styl(worksheet, mahsulot.Length + 10, 2);
                                        worksheet.Range[mahsulot.Length + 10, 2].Text = "Xisobot tuzuvch Yatt raxbari";
                                        styl(worksheet, mahsulot.Length + 10, 6);
                                        worksheet.Range[mahsulot.Length + 10, 6].Text = "_______________________________";

                                        //Faktura

                                        worksheet = workbook.Worksheets[zet+1];
                                        worksheet.Name = "Faktura";


                                        worksheet.Range[2, 2, 2, 10].Merge();
                                        worksheet.Range[2, 2].Text = $"{oyNomi.ToUpper()} {moment.Year} yildagi ________ sonli";
                                        styl(worksheet, 2, 2);
                                        worksheet.Range[3, 2, 3, 10].Merge();
                                        worksheet.Range[3, 2].Text = $"HISOB VARAQ -FAKTURA ";
                                        styl(worksheet, 3, 2);
                                        worksheet.Range[4, 2, 4, 10].Merge();
                                        worksheet.Range[4, 2].Text = $"__________{moment.Year} yil №_________shartnoma";
                                        styl(worksheet, 4, 2);
                                        worksheet.Range[6, 3].Text = $"«Ijrochi»: ________________________________";
                                        styl(worksheet, 6, 3);
                                        worksheet.Range[6, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                        
                                        worksheet.Range[6, 7].Text = $"  «Buyurtmachi»: ________________________";
                                        styl(worksheet, 6, 7);
                                        worksheet.Range[6, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[7, 3].Text = $"Маnzil: ___________________________________          ";
                                        styl(worksheet, 7, 3);
                                        worksheet.Range[7, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[7, 7].Text = $"Маnzil: ___________________________________       ";
                                        styl(worksheet, 7, 7);
                                        worksheet.Range[7, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[8, 3].Text = $"Теl./faks:__________________________________";
                                        styl(worksheet, 8, 3);
                                        worksheet.Range[8, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[8, 7].Text = $"Теl./faks:__________________________________";
                                        styl(worksheet, 8, 7);
                                        worksheet.Range[8, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[9,3].Text = $"Xisob-kitob xisob varaq raqami:";
                                        styl(worksheet, 9, 3);
                                        worksheet.Range[9, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[9, 7].Text = $"ShG'X:________________________________";
                                        styl(worksheet, 9, 7);
                                        worksheet.Range[9, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[10, 3].Text = $"№ ________________________________________";
                                        styl(worksheet, 10, 3);
                                        worksheet.Range[10, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[10, 7].Text = $"Tashkilotning STIRIi: ";
                                        styl(worksheet, 10, 7);
                                        worksheet.Range[10, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[11, 3].Text = $"Bank ______________________________________";
                                        styl(worksheet, 11, 3);
                                        worksheet.Range[11, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[11, 7].Text = $"IFUT: ";
                                        styl(worksheet, 11, 7);
                                        worksheet.Range[11, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[12, 7].Text = $"O'zbekiston Respublikasi Moliya  ";
                                        styl(worksheet, 12, 7);
                                        worksheet.Range[12, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[13,3].Text = $"STIR:___________________________";
                                        styl(worksheet, 13, 3);
                                        worksheet.Range[13, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[13, 7].Text = $"vazirligi G'aznachiligi ";
                                        styl(worksheet, 13, 7);
                                        worksheet.Range[13, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                        
                                        string[] faktura = { "YaG'H: 23 402 000 300 100 001 010 ", "Bank nomi: Markaziy bankning Тоshкеnт  ", "shahаr Bosh boshqarmasi (HККМ) ", "МFО:00014  G'aznachilik bo'linmasining", "SТIRi:201122919      " };
                                        for(int i=14; i<19; i++)
                                        {
                                            worksheet.Range[i, 7].Text = faktura[i-14];
                                            styl(worksheet, i, 7);
                                            worksheet.Range[i, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                        }

                                        string[] fakturaName = { "№", "Tovar (ish, xizmatlar) nomi", "O'lchov birligi", "Miqdori", "Narxi", "Yetkazib Berish\nQiymati", "QQS","So'mmasi", "Yetkazib\nBerishning\nQQS hisobga olingan" };
                                        for(int i=2;i<11; i++)
                                        {
                                            
                                            if (i == 9)
                                            {
                                                worksheet.Range[21, i].Text = fakturaName[i - 2];
                                                styl(worksheet, 21, i);
                                                worksheet.Range[21, i].RowHeight = 30;
                                                worksheet.Range[21, i].ColumnWidth = 15;
                                            }
                                            else if(i == 8)
                                            {
                                                worksheet.Range[20, i].Text = fakturaName[i-2];
                                                styl(worksheet, 20, i);
                                                worksheet.Range[21, i].Text = "Stavkasi";
                                                worksheet.Range[21, i].ColumnWidth = 15;
                                                styl(worksheet, 21, i);
                                                worksheet.Range[20, i, 20, i+1].Merge();
                                                worksheet.Range[20, i].RowHeight = 30;
                                            }
                                            else
                                            {
                                                worksheet.Range[20, i].Text = fakturaName[i - 2];
                                                styl(worksheet, 20, i);
                                                worksheet.Range[20, i, 21, i].Merge();
                                                worksheet.Range[20, i].ColumnWidth = 15;
                                            }
                                            if(i==3)
                                                worksheet.Range[20, i].ColumnWidth = 25;

                                        }
                                       
                                        for(int i=0; i<mahsulot.Length; i++)
                                        {
                                            worksheet.Range[i + 22, 2].RowHeight = 32;
                                            worksheet.Range[i + 22, 2].Number = i + 1;
                                            styl(worksheet, i + 22, 2);
                                            worksheet.Range[i + 22, 3].Text=mahsulot[i];
                                            styl(worksheet, i + 22, 3);
                                            worksheet.Range[i + 22, 4].Text = mahsulotBirlik[i];
                                            styl(worksheet, i + 22, 4);
                                            worksheet.Range[i + 22, 5].Number = mahsulotOylikMiqdor[i];
                                            styl(worksheet, i + 22, 5);
                                            worksheet.Range[i + 22, 6].Number = mahsulotNarx[i];
                                            styl(worksheet, i + 22, 6);
                                            worksheet.Range[i + 22, 7].Number = mahsulotNarx[i]* mahsulotOylikMiqdor[i];
                                            styl(worksheet, i + 22, 7);



                                        }
                                        for(int i=2; i<11; i++)
                                        {
                                            styl(worksheet, mahsulot.Length + 22, i);
                                        }
                                        worksheet.Range[mahsulot.Length + 22, 3].Text="Jami";

                                        worksheet.Range[mahsulot.Length + 22, 7].Formula=$"=SUM(G22:G{mahsulot.Length+21})";

                                        worksheet.Range[20,2,(mahsulot.Length + 22),10].BorderInside(ExcelLineStyle.Thin, Color.Black);
                                        worksheet.Range[20, 2, (mahsulot.Length + 22), 10].BorderAround(ExcelLineStyle.Thin, Color.Black);
                                        string[] yetkazib = { "So'mma so'z bilan:", "Rahbar: ", "Bosh hisobchi: ", "М.O'", "Tovarni berdim______________________" };
                                        for(int i= mahsulot.Length + 25; i< mahsulot.Length + 30; i++)
                                        {
                                            worksheet.Range[i, 3].Text = yetkazib[i- mahsulot.Length - 25];
                                            styl(worksheet, i, 3);
                                            worksheet.Range[i, 3].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                        }

                                        worksheet.Range[mahsulot.Length + 26, 7].Text = "Oldim: ___________________";
                                        styl(worksheet, mahsulot.Length + 26, 7);
                                        worksheet.Range[mahsulot.Length + 26, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[mahsulot.Length + 27, 7].Text = "                             (Buyurtmachi yoki vakolatli shaxs)";

                                        worksheet.Range[mahsulot.Length + 28, 7].Text = $"                  {moment.Year}-yil «____» «___________» dagi";
                                        styl(worksheet, mahsulot.Length + 28, 7);
                                        worksheet.Range[mahsulot.Length + 28, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                        worksheet.Range[mahsulot.Length + 29, 7].Text = $"        ______ - sonli ishonchnoma bo'yicha";
                                        styl(worksheet, mahsulot.Length + 29, 7);
                                        worksheet.Range[mahsulot.Length + 29, 7].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;

                                    }

                                    con.Close();
                                }
                                
                            }

                            





                            workbook.SaveAs(fileInfo.FullName);
                            MessageBox.Show("Saqalndi");

                        }
                    }
                }
            }
          /*  }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }*/

        }

        private void Jadval_Load(object sender, EventArgs e)
        {

        }

        private void Jadval_FormClosed(object sender, FormClosedEventArgs e)
        {
            if(close) new Form1().Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            close = false;
            this.Close();
            new Form2().Show();

        }
    }
}
