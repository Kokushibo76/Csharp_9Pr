using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Win32;
using System.Windows.Documents;
using System.Collections.Generic;
using System.Windows.Controls;

namespace _9_praktos
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void WordCreate_Click(object sender, RoutedEventArgs e)
        {
            Word word = new Word();
            word.Show();
            this.Close();
        }

        private void WordOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            bool? success = fileDialog.ShowDialog();
            if (success == true)
            {
                string path = fileDialog.FileName;
                LoadWordFile(path);
            }
        }

        private void ExcelCreate_Click(object sender, RoutedEventArgs e)
        {
            Excel excel = new Excel();
            excel.Show();
            this.Close();
        }

        private void ExcelOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            bool? success = fileDialog.ShowDialog();
            if (success == true)
            {
                string path = fileDialog.FileName;
                LoadExcelFile(path);
            }
        }

        private void SendEmail(string from, string to, string subject, string body)
        {
            MailMessage mail = new MailMessage(from, to, subject, body);
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new NetworkCredential("your-email@gmail.com", "your-password");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
        }

        private void SaveWordFile(string _fileName)
        {
            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            FileStream fStream = new FileStream(_fileName, FileMode.Create);
            range.Save(fStream, DataFormats.Rtf);
            fStream.Close();
        }

        private void LoadWordFile(string _fileName)
        {
            if (File.Exists(_fileName))
            {
                TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                FileStream fStream = new FileStream(_fileName, FileMode.OpenOrCreate);
                range.Load(fStream, DataFormats.Rtf);
                fStream.Close();
            }
        }

        private void SaveExcelFile(string _fileName)
        {
            try
            {
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                for (int i = 0; i < MyDataGrid.Items.Count; i++)
                {
                    for (int j = 0; j < MyDataGrid.Columns.Count; j++)
                    {
                        DataGridColumn column = MyDataGrid.Columns[j];
                        string cellValue = (column.GetCellContent(MyDataGrid.Items[i]) as TextBlock).Text;
                        worksheet.Range[i + 1, j + 1].Value = cellValue;
                    }
                }

                workbook.SaveToFile(_fileName, ExcelVersion.Version2016);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении файла Excel: " + ex.Message);
            }
        }

        private void LoadExcelFile(string _fileName)
        {
            try
            {
                Workbook workbook = new Workbook();

                workbook.LoadFromFile(_fileName);

                Worksheet worksheet = workbook.Worksheets[0];

                MyDataGrid.ItemsSource = null;

                List<object[]> data = new List<object[]>();

                for (int i = 1; i <= worksheet.Cells.MaxDataRow; i++)
                {
                    object[] row = new object[worksheet.Cells.MaxDataColumn];
                    for (int j = 1; j <= worksheet.Cells.MaxDataColumn; j++)
                    {
                        row[j - 1] = worksheet.Range[i, j].Value;
                    }
                    data.Add(row);
                }

                MyDataGrid.ItemsSource = data;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке файла Excel: " + ex.Message);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                if (saveFileDialog.ShowDialog() == true)
                {
                    string path = saveFileDialog.FileName;
                    SaveExcelFile(path);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении файла Excel: " + ex.Message);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                if (openFileDialog.ShowDialog() == true)
                {
                    string path = openFileDialog.FileName;
                    LoadExcelFile(path);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при загрузке файла Excel: " + ex.Message);
            }
        }
    }
}