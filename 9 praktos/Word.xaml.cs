using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace _9_praktos
{
    /// <summary>
    /// Логика взаимодействия для Word.xaml
    /// </summary>
    public partial class Word : Window
    {
        public Word()
        {
            InitializeComponent();
        }

        private void SaveRtfFile(string _fileName)
        {
            TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
            FileStream fStream = new FileStream(_fileName, FileMode.Create);
            range.Save(fStream, DataFormats.Rtf);
            fStream.Close();
        }

        private void LoadRtfFile(string _fileName)
        {
            if (File.Exists(_fileName))
            {
                TextRange range = new TextRange(MyRtb.Document.ContentStart, MyRtb.Document.ContentEnd);
                FileStream fStream = new FileStream(_fileName, FileMode.OpenOrCreate);
                range.Load(fStream, DataFormats.Rtf);
                fStream.Close();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                bool? success = fileDialog.ShowDialog();
                if (success == true)
                {
                    string path = @"" + fileDialog.FileName;
                    SaveRtfFile(path);
                }
            } catch (Exception ex)
            {
                MessageBox.Show(e.ToString());
            }

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog();
                bool? success = fileDialog.ShowDialog();
                if (success == true)
                {
                    string path = @"" + fileDialog.FileName;
                    LoadRtfFile(path);
                }
            } catch (Exception ex)
            {
                MessageBox.Show(e.ToString());
            }
        }
    }
}
