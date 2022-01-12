using System;
using System.IO;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;



using PostSharp.Patterns.Threading;

namespace SAP_Import
{
    public class ExcelClass
    {
        public string path = "";
        public Excel.Application excel = new Excel.Application();
        public Excel.Workbook wb;
        public Excel.Worksheet ws;
        public Excel.Range range;

        public int rows;
        public int cols;


        public ExcelClass(string path, int sheet)
        {
            try
            {
                this.path = path;
                wb = excel.Workbooks.Open(path);
                ws = wb.Worksheets[sheet];

                range = ws.UsedRange;
                rows = range.Rows.Count;
                cols = range.Columns.Count;
            }
            catch
            {
                MessageBox.Show("Unable To Create Excel Class");
            }
        }

        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null)
            {
                return (ws.Cells[i, j].Value2).ToString();
            }
            else
            {
                return "";
            }
        }

        public void Release()
        {
            Marshal.ReleaseComObject(ws);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(excel);
        }
    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ExcelClass ImportExcelFile;
        bool hasFile = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void DragBlock_DragOver(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(200, 200, 255));
        }

        private void DragBlock_DragLeave(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(200, 200, 200));
        }

        [Background]
        private void DragBlock_Drop(object sender, DragEventArgs e)
        {
            ProgressStart();
            hasFile = false;
            Dispatcher.BeginInvoke((Action)(() =>
            {
                Convert.IsEnabled = false;
            }));

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string filename = System.IO.Path.GetFileName(files[0]);
                string path = System.IO.Path.GetFullPath(files[0]);
                string extension = System.IO.Path.GetExtension(files[0]);

                Dispatcher.BeginInvoke((Action)(() =>
                {
                    DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(200, 255, 200));
                    Filename.Text = filename + "\n\n" + extension;
                }));

                if (extension == ".xlsx")
                {
                    ImportExcelFile = new ExcelClass(path, 1);
                    hasFile = true;
                    Dispatcher.BeginInvoke((Action)(() =>
                    {
                        Convert.IsEnabled = true;
                    }));
                }
                else
                {
                    MessageBox.Show("File Does Not Have .xlsx Extension", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(255, 200, 200));
                }));
            }

            ProgressEnd();
        }


        [Background]
        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            if (hasFile == true)
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    Convert.IsEnabled = false;
                }));
                ProgressStart();

                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine("Rows: " + ImportExcelFile.rows.ToString());
                Console.WriteLine("Cols: " + ImportExcelFile.cols.ToString());
                Console.WriteLine("------------------------------------------------------------");

                for (int i = 0; i < ImportExcelFile.rows; i++)
                {
                    for (int j = 0; j < ImportExcelFile.cols; j++)
                    {
                        Console.Write(ImportExcelFile.ReadCell(i, j) + "|");
                    }
                    Console.WriteLine();
                }

                ImportExcelFile.Release();

                Console.WriteLine("------------------------------------------------------------");

                ProgressEnd();
            }
            else
            {
                MessageBox.Show("No File Selected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // Begin Progress Bar Loading Animation
        [Dispatched]
        private void ProgressStart()
        {
            ProgressBar.Visibility = Visibility.Visible;
            ProgressBar.IsIndeterminate = true;
        }

        // Ends Progress Bar Loading Animation
        [Dispatched]
        private void ProgressEnd()
        {
            ProgressBar.Visibility = Visibility.Collapsed;
            ProgressBar.IsIndeterminate = false;
        }
    }
}
