using System.Collections;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDBEntegrasyonProj
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        SqlConnection connection = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=ProjelerVT;Integrated Security=True");

        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excelUygulama = new Excel.Application();
            excelUygulama.Visible = true;
            Excel.Workbook wb = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            Excel.Worksheet ws1 = wb.Sheets[1];

            string[] basliklar = { "Personel No", "Ad", "Soyad", "Semt", "Sehir" };
            Excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = ws1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
            }

            try
            {
                connection.Open();
                string sqlCom = "SELECT PersonelNo, Ad, Soyad, Semt, Sehir FROM Personel";
                SqlCommand cmd = new SqlCommand(sqlCom, connection);
                SqlDataReader reader = cmd.ExecuteReader();


                int satir = 2;  //ilk satýr baþlýktý, ikinci satýr ile devam ediyoruz.
                while (reader.Read())
                {
                    string pno = reader[0].ToString();
                    string ad = reader[1].ToString();
                    string soyad = reader[2].ToString();
                    string semt = reader[3].ToString();
                    string sehir = reader[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + " " + pno + " " + ad + " " + soyad + " " + semt + " " + sehir + "\n";

                    range = ws1.Cells[satir, 1];
                    range.Value2 = pno;
                    range = ws1.Cells[satir, 2];
                    range.Value2 = ad;
                    range = ws1.Cells[satir, 3];
                    range.Value2 = soyad;
                    range = ws1.Cells[satir, 4];
                    range.Value2 = semt;
                    range = ws1.Cells[satir, 5];
                    range.Value2 = sehir;
                    satir++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sql query sýrasýnda bir hata oluþtu, Hata Kodu: SQLREAD01 \n" + ex.ToString());
            }
            finally
            {
                if (connection != null)
                {
                    connection.Close();
                }
            }
        }

        private void btnExceldenOku_Click(object sender, EventArgs e)
        {
            Excel.Application exlApp;
            Excel.Workbook exlWorkbook;
            Excel.Worksheet exlWorksheet;
            Excel.Range range;
            int rCnt = 0;
            int cCnt = 0;
            exlApp = new Excel.Application();
            exlWorkbook = exlApp.Workbooks.Open("C:\\test\\test.xlsx");
            exlWorksheet = (Excel.Worksheet)exlWorkbook.Worksheets.get_Item(1);
            range = exlWorksheet.UsedRange;

            // Ýlk olarak richTextBox2 içeriðini temizleyelim.
            richTextBox2.Clear();

            // Ýlk satýr baþlýklarý içerdiði için row Count (rCnt)' u 2'den baþlatmamýz gerekiyor.
            // Eðer ilk satýrda veriler baþlamýþ olsaydý 1 ' den baþlatmamýz gerekirdi.
            for (rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                ArrayList list = new ArrayList();
                for(cCnt = 1;  cCnt <= range.Columns.Count; cCnt++)
                {
                    string okunanHucre = Convert.ToString((range.Cells[rCnt, cCnt] as Excel.Range).Value2);
                    richTextBox2.Text = richTextBox2.Text + okunanHucre + " ";
                    list.Add(okunanHucre);
                }
                richTextBox2.Text = richTextBox2.Text + "\n";

                try 
                { 
                    connection.Open();
                    SqlCommand cmd = new SqlCommand("INSERT INTO Personel (PersonelNo, Ad, Soyad, Semt, Sehir)" 
                                                    + "VALUES (@P1, @P2, @P3, @P4, @P5)", connection);
                    cmd.Parameters.AddWithValue("@P1", list[0]);
                    cmd.Parameters.AddWithValue("@P2", list[1]);
                    cmd.Parameters.AddWithValue("@P3", list[2]);
                    cmd.Parameters.AddWithValue("@P4", list[3]);
                    cmd.Parameters.AddWithValue("@P5", list[4]);
                    cmd.ExecuteNonQuery();

                } 
                catch (Exception ex) 
                {
                    MessageBox.Show("Veritabanýna yazarken hata oluþtu! Hata kodu: SQLWRITE01\n" + ex.ToString());
                }
                finally 
                {
                    if (connection != null)
                        connection.Close(); 
                }
            }
            exlApp.Quit();
            RelaseObject(exlWorksheet);
            RelaseObject(exlWorkbook);
            RelaseObject(exlApp);
        }

        private void RelaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
