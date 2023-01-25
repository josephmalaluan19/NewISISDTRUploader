
using ExcelDataReader;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace NewUploader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void ExcelUploader_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "C://";
            openFileDialog1.Title = "Select file to be upload.";
            openFileDialog1.Filter = "Select Valid Document(*.xlsx)| *.xlsx";
            openFileDialog1.FilterIndex = 1;
            try
            {
                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (openFileDialog1.CheckFileExists)
                    {
                        string path = System.IO.Path.GetFullPath(openFileDialog1.FileName);
                        label1.Text = path;
                    }
                }
                else
                {
                    MessageBox.Show("Please Upload document.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnUploadExcel_Click(object sender, EventArgs e)
        {
            try
            {
                string filename = System.IO.Path.GetFileName(openFileDialog1.FileName);
                if (filename == null)
                {
                    MessageBox.Show("Please select a valid document.");
                }
                else
                {
                    System.IO.File.Delete(@"C:\\ISISFile\" + filename);
                    System.IO.File.Copy(openFileDialog1.FileName, @"C:\\ISISFile\" + filename);

                    ReadExcelFile();
                    MessageBox.Show("Document uploaded.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void ReadExcelFile()
        {
            int counter = 0;
            EmployeeData EmployeeData = new EmployeeData();
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(@"C:\ISISFile\"+ System.IO.Path.GetFileName(openFileDialog1.FileName), FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    do
                    {
                        while (reader.Read()) //Each ROW
                        {
                            if (counter > 0)
                            {
                                for (int column = 0; column < reader.FieldCount; column++)
                                {
                                    if (column == 0)
                                        EmployeeData.Person_ID = reader.GetValue(column).ToString();
                                    if (column == 0)
                                        EmployeeData.EmployeeNumber = reader.GetValue(column).ToString(); //Get this from ISIS new table as I instructed
                                    if (column == 2)
                                        EmployeeData.CheckDate = reader.GetValue(column).ToString();
                                    if (column == 3)
                                        EmployeeData.CheckInTime = reader.GetValue(column).ToString();
                                    if (column == 4)
                                        EmployeeData.CheckoutTime = reader.GetValue(column).ToString();


                                }
                                EmployeeData = new EmployeeData();
                            }

                            counter = counter + 1;
                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET

                }
            }
        }
    }
}

public class EmployeeData
{
    public string Person_ID { get; set; }
    public string EmployeeNumber { get; set; }
    public string CheckDate { get; set; }
    public string CheckInTime { get; set; }
    public string CheckoutTime { get; set; }
}