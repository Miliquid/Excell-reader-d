using System.IO;
using System.Windows;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel;

namespace Excel_reader
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {   private Application.Application range;
        private Application.Application app;
        private Application.Workbooks wrbks=null;
        public Application.Workbook wrbk = null;
        private Application.Worksheet wrsh;
        public string b;
        
        public MainWindow()
        {
            InitializeComponent();

        }


        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
          


        
            using  (var fl = new FolderBrowserDialog())

            if (fl.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {var d = fl.SelectedPath;

                    FileInfo targetDir = new FileInfo(d);
                    Sfile.Text = targetDir.FullName;
                    
            }

        }

        private void OpenFileF_Click(object sender, RoutedEventArgs e)
        {
            var d = new OpenFileDialog();


            d.InitialDirectory = @"C:\\";
            d.Filter = "excel files (*.xls, *.xlsx)|*.xls;*.xlsx";
            d.FilterIndex = 2;
            d.RestoreDirectory = true;

            if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {

                FFile.Text = d.FileName;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {app = new Application.Application { DisplayAlerts = true };
            var Lo = new Logic();


            if (Sfile.Text == null)
            {
                System.Windows.MessageBox.Show("Введите путь сохранения");
            }
            else { Lo.Read(FFile.Text, Fio.Text, semestr.SelectedIndex, Sfile.Text); }

        }


    }
}