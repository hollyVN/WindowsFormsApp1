using System;
using System.IO; // to create a stream for file
using System.Net; // for request and respond with http protocol
using System.Windows.Forms;
using System.Data.SqlClient;
using ExcelDataReader;
using System.Configuration; // this package to reflect the server
using ClosedXML.Excel; // this is a library for working with file *.xlsx
using System.Globalization;
using System.Resources;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private ResourceManager resourceManager;
        public Form1()
        {
            InitializeComponent();
            CultureInfo culture = CultureInfo.CurrentCulture;
           // SetLanguage("en_US");
        }
       
        
        private void Form1_Load(object sender, EventArgs e)
        {
            // set english is default
            System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("en_US");

            
        }

        string excelFilePath = "sample.xlsx";
        string sqlServerStr = @"Data Source=DESKTOP-I6017GJ\SQLSERVER;Initial Catalog=SampleTest;Integrated Security=True";
       // string serverUrl = ConfigurationManager.AppSettings["ServerUrl"];
        

        private void btnReadFile_Click(object sender, EventArgs e)
        {
            using (var stream = File.Open(excelFilePath,FileMode.Open,FileAccess.Read)) // create a stream use System.IO push filepath into
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))// create a reader and push stream into
                {
                    bool isFirstRow = true;
                    do
                    {
                        while (reader.Read())
                        {
                            if (isFirstRow)
                            {
                                isFirstRow = false;
                                continue;
                            }
                            string data = reader.GetString(0);// read data from column 0 in fileexcel

                            // after get data will connect to sql server
                            using (SqlConnection sqlConn = new SqlConnection(sqlServerStr))
                            {
                                sqlConn.Open();
                                
                                string sqlCmd = "MERGE INTO dbo.matches AS target " +
                                                        "USING (SELECT @epc_code AS epc_code) AS source " +
                                                        "ON target.epc_code = source.epc_code " +
                                                        "WHEN MATCHED THEN " +
                                                        "   UPDATE SET updated_at = GETDATE() " +
                                                        "WHEN NOT MATCHED THEN " +
                                                        "   INSERT (created_at, updated_at,ri_date, epc_code) " +
                                                        "VALUES  (GETDATE(), GETDATE(), GETDATE(), @epc_code);";
                                using (SqlCommand cmd = new SqlCommand(sqlCmd, sqlConn))
                                {
                                    cmd.Parameters.AddWithValue("@epc_code", data);
                                    
                                    cmd.ExecuteNonQuery();
                                }
                                // send data to server use HTTP POST request
                                //var postData = "data=" + data; // this data we already declare
                                //var dataBytes = Encoding.UTF8.GetBytes(postData);

                                //var request = (HttpWebRequest)WebRequest.Create(serverUrl);
                                //request.Method = "POST";
                                //request.ContentType = "application/x-www-form-urlencoded";
                                //request.ContentLength = dataBytes.Length;

                                //using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                                //{
                                //    streamWriter.Write(postData);
                                //}

                                //var response = (HttpWebResponse)request.GetResponse();
                                //var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();
                            }

                        }
                    } while (reader.NextResult());
                }
            }
            MessageBox.Show("Bạn đã import dữ liệu thành công lên database");
        }

        string exportQuery = "select ri_date, epc_code from dbo.matches";
        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            SqlConnection sqlConn = new SqlConnection(sqlServerStr);
            sqlConn.Open();
            SqlCommand cmd = new SqlCommand(exportQuery, sqlConn);
            // truyền data từ sql server vào một đầu đọc reader
            SqlDataReader reader = cmd.ExecuteReader();
            var workbook = new XLWorkbook(); // create a workbook
            var worksheet = workbook.Worksheets.Add("ExportData"); // create a worksheet
            worksheet.Cell(1, 1).Value = "ri_date";// header of column 1
            worksheet.Cell(1, 2).Value = "epc_code"; // header of column 2

            // tao loop truyền vào file
            int row = 2;
            while (reader.Read())
            {
                worksheet.Cell(row, 1).Value = (DateTime)reader["ri_date"];
                worksheet.Cell(row, 2).Value = (string)reader["epc_code"];
                row++;
            }
            workbook.SaveAs("exportEpcCode.xlsx");
            MessageBox.Show("Export thành công");
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            string selectedLanguage = comboBox.SelectedItem.ToString();
            if(selectedLanguage == "english")
            {
                SetLanguage("en_US");
            }
            else
            {
                SetLanguage("vi_VN");
            }
            
        }
        private void SetLanguage(string lang)
        {
            CultureInfo culture = CultureInfo.CreateSpecificCulture(lang);
            ResourceManager rm = new ResourceManager("WindowsFormsApp1.Properties.Resources", typeof(Form1).Assembly);

            btnReadFile.Text = rm.GetString("ReadFile", culture);
            btnStop.Text = rm.GetString("Stop", culture);
            btnExportExcel.Text = rm.GetString("ExportExcel", culture);

            
        }
    }
}
