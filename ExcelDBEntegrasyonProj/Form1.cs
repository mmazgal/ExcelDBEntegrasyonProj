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
                range = ws1.Cells[1, (1+i)];
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
    }
}
