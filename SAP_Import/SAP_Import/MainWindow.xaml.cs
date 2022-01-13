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

using forms = System.Windows.Forms;


using Microsoft.Win32;



using PostSharp.Patterns.Threading;

namespace SAP_Import
{
    public class ExcelRead
    {
        public string filename = "";
        public string path = "";
        public Excel.Application excel;
        public Excel.Workbook wb;
        public Excel.Worksheet ws;
        public Excel.Range range;

        public int rows;
        public int cols;


        public ExcelRead(string path, int sheet, string name)
        {
            try
            {
                this.filename = name;
                this.path = path;

                this.excel = new Excel.Application();                
                this.wb = excel.Workbooks.Open(path);
                this.ws = wb.Worksheets[sheet];
                this.range = ws.UsedRange;

                this.rows = range.Rows.Count;
                this.cols = range.Columns.Count;
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
            //excel.Workbooks.Close();
            //wb.Close(true, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            this.excel.Quit();

            Marshal.ReleaseComObject(this.ws);
            Marshal.ReleaseComObject(this.wb);
            Marshal.ReleaseComObject(this.excel);
        }
    }

    public class ExcelCreate
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        public ExcelCreate()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
        }
    }


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        ExcelRead ImportExcelFile;
        bool hasFile = false;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void DragBlock_DragOver(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 200, 255));
        }

        private void DragBlock_DragLeave(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 200, 200));
        }

        [Background]
        private void DragBlock_Drop(object sender, DragEventArgs e)
        {
            ProgressStart();
            hasFile = false;
            Dispatcher.BeginInvoke((Action)(() =>
            {
                Convert_Back.IsEnabled = false;
            }));

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                string filename = System.IO.Path.GetFileName(files[0]);
                string path = System.IO.Path.GetFullPath(files[0]);
                string extension = System.IO.Path.GetExtension(files[0]);

                
                if (extension == ".xlsx")
                {
                    ImportExcelFile = new ExcelRead(path, 1, filename);
                    hasFile = true;

                    Dispatcher.Invoke((Action)(() =>
                    {
                        ConsoleWindow.Children.Clear();

                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 255, 200));
                        Filename.Text = filename + "\n\n" + extension;

                        print("Excel File Information:", 12, "bold");
                        print("line", 0, "");
                        print("Rows: " + ImportExcelFile.rows, 12, "");
                        print("Cols: " + ImportExcelFile.cols, 12, "");
                        print("line", 0, "");
                    }));

                    
                    Dispatcher.BeginInvoke((Action)(() =>
                    {
                        Convert_Back.IsEnabled = true;
                    }));
                }
                else
                {
                    MessageBox.Show("File Does Not Have .xlsx Extension", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Dispatcher.BeginInvoke((Action)(() =>
                    {
                        ConsoleWindow.Children.Clear();
                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 255, 200, 200));
                        Filename.Text = "Drag .XLSX File";

                        
                    }));
                }
            }
            else
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    ConsoleWindow.Children.Clear();
                    DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 255, 200, 200));
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
                    Convert_Back.IsEnabled = false;
                }));
                ProgressStart();

                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine("Rows: " + ImportExcelFile.rows.ToString());
                Console.WriteLine("Cols: " + ImportExcelFile.cols.ToString());
                Console.WriteLine("------------------------------------------------------------");

                //for (int i = 0; i < ImportExcelFile.rows; i++)
                //{
                //    for (int j = 0; j < ImportExcelFile.cols; j++)
                //    {
                //        Console.Write(ImportExcelFile.ReadCell(i, j) + "|");
                //    }
                //    Console.WriteLine();
                //}


                bool isConverted;
                try
                {
                    ConvertBOM();

                    ImportExcelFile.Release();
                    isConverted = true;
                }
                catch
                {
                    isConverted = false; ;
                }

                ImportExcelFile.Release();

                Console.WriteLine("------------------------------------------------------------");

                Dispatcher.Invoke((Action)(() =>
                {
                    Filename.Text = "Drag .XLSX File";

                    if (isConverted == true)
                    {
                        Status.Text = "Success";
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 255, 200));
                    }
                    else
                    {
                        Status.Text = "Failed";
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 255, 200, 200));
                    }
                }));

                ProgressEnd();
            }
            else
            {
                MessageBox.Show("No File Selected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [Background]
        private void ConvertBOM()
        {
            print("Begnning Conversion:", 12, "bold");
            print("line", 0, "");

            //CreateBOMs();
        }

        [Background]
        private void CreateBOMs()
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            //xlWorkBook.Worksheets.Add
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "ID";
            xlWorkSheet.Cells[1, 2] = "Name";
            xlWorkSheet.Cells[2, 1] = "1";
            xlWorkSheet.Cells[2, 2] = "One";
            xlWorkSheet.Cells[3, 1] = "2";
            xlWorkSheet.Cells[3, 2] = "Two";

            string savePath = "";

            savePath = ImportExcelFile.path;

            // Keep This Code To Select Folder To Save IN
            //Dispatcher.Invoke((Action)(() =>
            //{
            //    forms.FolderBrowserDialog openFileDialog = new forms.FolderBrowserDialog();

            //    if (openFileDialog.ShowDialog() == forms.DialogResult.OK)
            //    {
            //        savePath = openFileDialog.SelectedPath;
            //    }
            //}));

            //txtEditor.Text = File.ReadAllText(openFileDialog.FileName);
            //savePath = Path.Get

            string temp = savePath.Replace(ImportExcelFile.filename, "BOMS.csv");

            Console.WriteLine(temp);
            print(temp, 14, "");

            try
            {
                xlWorkBook.SaveAs(temp);
            }
            catch
            {

            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        [Dispatched]
        private void print(string text, double size, string weight)
        {
            if (text == "line")
            {
                Border line = new Border();
                line.Height = 1;
                line.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(150, 150, 150));
                line.Margin = new Thickness(0, 1, 0, 1);
                line.CornerRadius = new CornerRadius(0);

                ConsoleWindow.Children.Add(line);
            }
            else
            {

                TextBlock tb = new TextBlock();
                tb.Text = text;
                tb.FontSize = size;
                if (weight == "bold")
                {
                    tb.FontWeight = FontWeights.Bold;
                }
                else
                {
                    tb.FontWeight = FontWeights.Regular;
                }

                ConsoleWindow.Children.Add(tb);
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
