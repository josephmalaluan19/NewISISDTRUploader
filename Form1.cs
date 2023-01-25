
using ExcelDataReader;
using IronXL;
using System.Data.OleDb;
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

            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = workbook.CreateWorkSheet("Result Sheet");
            sheet["A1"].Value = "badgeno";
            sheet["B1"].Value = "badgename";
            sheet["C1"].Value = "empno";
            sheet["D1"].Value = "date";
            sheet["E1"].Value = "in_";
            sheet["F1"].Value = "out_";

            workbook.SaveAs("FingerTechFormatted.xlsx");

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
                                    {
                                        EmployeeData.Person_ID = reader.GetValue(column).ToString().Replace("'", "");
                                        sheet["A" + (counter + 1)].Value = EmployeeData.Person_ID;
                                    }
                                    if (column == 1)
                                    {
                                        EmployeeData.EmployeeNumber = reader.GetValue(column).ToString(); //Get this from ISIS new table as I instructed
                                    }
                                    if (column == 2)
                                    {
                                        if (reader.GetValue(column) == null)
                                            EmployeeData.CheckDate = "";
                                        else
                                            EmployeeData.CheckDate = reader.GetValue(column).ToString();

                                        sheet["D" + (counter + 1)].Value = EmployeeData.CheckDate;
                                    }
                                    if (column == 3)
                                    {
                                        if (reader.GetValue(column) == null)
                                            EmployeeData.CheckInTime = "";
                                        else
                                            EmployeeData.CheckInTime = reader.GetValue(column).ToString();

                                        sheet["E" + (counter + 1)].Value = EmployeeData.CheckInTime;
                                    }
                                    if (column == 4)
                                    {
                                        if (reader.GetValue(column) == null)
                                            EmployeeData.CheckoutTime = "";
                                        else
                                            EmployeeData.CheckoutTime = reader.GetValue(column).ToString();

                                        sheet["F" + (counter + 1)].Value = EmployeeData.CheckoutTime;
                                    }


                                }
                                EmployeeData = new EmployeeData();
                            }

                            counter = counter + 1;
                        }
                    } while (reader.NextResult()); //Move to NEXT SHEET

                }
            }

            workbook.SaveAs("C:\\ISISFile\\FingerTechFormatted.xlsx");
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
