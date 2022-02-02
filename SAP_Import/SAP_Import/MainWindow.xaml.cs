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
using System.Windows.Threading;



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
        string path = "";
        string filename = "";
        bool hasFile = false;
        bool isPos;
        bool isConverted;
        string name = "";
        string warehouse = "";
        int partCount = 0;

        ProgressBar Prog_BOMs = new ProgressBar();
        ProgressBar Prog_BOM_Items = new ProgressBar();
        ProgressBar Prog_BOM_Scraps = new ProgressBar();
        ProgressBar Prog_BOM_Coproducts = new ProgressBar();
        ProgressBar Prog_OITM = new ProgressBar();
        ProgressBar Prog_OITW = new ProgressBar();

        private Boolean AutoScroll = true;
        DispatcherTimer dispatcherTimer;
        DispatcherTimer dispatcherClock;

        int timer = 0;
        int totalTime = 0;

        public MainWindow()
        {
            dispatcherTimer = new DispatcherTimer();
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 100);
            dispatcherTimer.Start();

            dispatcherClock = new DispatcherTimer();
            dispatcherClock.Tick += new EventHandler(Clock_Timer);
            dispatcherClock.Interval = new TimeSpan(0, 0, 1);
            dispatcherClock.Start();

            InitializeComponent();
        }


        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
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
                filename = System.IO.Path.GetFileName(files[0]);
                path = System.IO.Path.GetFullPath(files[0]);
                string extension = System.IO.Path.GetExtension(files[0]);

                
                if (extension == ".xlsx")
                {
                    //ImportExcelFile = new ExcelRead(path, 1, filename);
                    hasFile = true;

                    Dispatcher.Invoke((Action)(() =>
                    {
                        ConsoleWindow.Children.Clear();

                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 255, 200));
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 233, 233, 233));
                        Filename.Text = filename;
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
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 233, 233, 233));
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

        // When "Convert" Button Clikced
        // Main Function
        [Background]
        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            totalTime = 0;
            name = "";
            Dispatcher.Invoke((Action)(() =>
            {
                name = BomName.Text;
                warehouse = Warehouse.Text;
            }));

            Console.WriteLine(name);

            if (hasFile == true && name != "" && warehouse.Length == 3)
            {
                ImportExcelFile = new ExcelRead(path, 1, filename);

                Dispatcher.BeginInvoke((Action)(() =>
                {
                    Convert_Back.IsEnabled = false;
                }));
                ProgressStart();

                Console.WriteLine("------------------------------------------------------------");
                Console.WriteLine("Rows: " + ImportExcelFile.rows.ToString());
                Console.WriteLine("Cols: " + ImportExcelFile.cols.ToString());
                Console.WriteLine("------------------------------------------------------------");


                isConverted = false;
                try
                {
                    ConvertBOM();
                    
                }
                catch
                {
                    isConverted = false; ;
                }

                ImportExcelFile.wb.Close(true, null, null);
                ImportExcelFile.excel.Quit();
                ImportExcelFile.Release();

                Console.WriteLine("------------------------------------------------------------");

                Dispatcher.Invoke((Action)(() =>
                {
                    Filename.Text = "Drag .XLSX File";

                    if (isConverted == true)
                    {
                        Status.Text = "Success";
                        printColor("Conversion Complete!", 14, "bold", 100, 155, 100);
                        printColor("Total Time: " + totalTime + "ms", 12, "", 100, 155, 100);
                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 233, 233, 233));
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 255, 200));
                    }
                    else
                    {
                        Status.Text = "Failed";
                        printColor("Conversion Failed!", 14, "bold", 255, 100, 100);
                        printColor("Total Conversion Time:" + (double)totalTime / 1000.0 + "Seconds", 14, "", 255, 100, 100);
                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 233, 233, 233));
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 255, 200, 200));
                    }
                }));

                ProgressEnd();
            }
            else
            {
                //MessageBox.Show("No File Selected", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

                if(name == "")
                {
                    
                        printColor("Please Enter BOM Name", 12, "bold", 255, 100, 100);
                    
                }

                if (warehouse == "")
                {

                    printColor("Please Enter a 3 Letter Warehouse Code", 12, "bold", 255, 100, 100);

                }
            }
        }

        // Helper Function that Initiates the Different Coverting Functions
        private void ConvertBOM()
        {
            isPos = false;
            CountPOS();

            if (isPos == true)
            {
                printColor("Correct Headings Found", 12, "", 100, 155, 100);
                print("Parts Count: " + partCount.ToString(), 10, "");
                print("Begnning Conversion:", 12, "bold");
                print("line", 0, "");

                try
                {
                    Create_BOMs();
                    Create_BOM_Items();
                    Create_BOM_Scraps();
                    Create_BOM_Coproducts();
                    Create_OITM();
                    Create_OITW();
                    isConverted = true;
                }
                catch
                {
                    isConverted = false;
                }
                
            }
            else
            {
                printColor("Unable to Find the Correct Headings", 12, "bold", 255, 100, 100);
            }
        }

        // Counts the Number of Parts in the Imported Excel File
        private void CountPOS()
        {
            partCount = 0;
            if(ImportExcelFile.ReadCell(0, 0) == "POS")
            {
                isPos = true;

                for(int i = 0; i < ImportExcelFile.rows; i++)
                {
                    int number = 0;

                    if (int.TryParse(ImportExcelFile.ReadCell(i, 0), out number))
                    {
                        partCount++;
                    }
                }
            }

        }


        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            timer = timer + 100;
            totalTime = totalTime + 100;
            CommandManager.InvalidateRequerySuggested();
        }

        private void Clock_Timer(object sender, EventArgs e)
        {
            TimeBlock.Text = DateTime.Now.ToString();
            CommandManager.InvalidateRequerySuggested();
        }

        // Convert 1
        private void Create_BOMs()
        {
            timer = 0;
            //timeSpan = new TimeSpan(0, 0, 0);

            print("BOMs", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOMs.Maximum = 100;
                Prog_BOMs.Height = 10;
                Prog_BOMs.Visibility = Visibility.Visible;
                Prog_BOMs.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_BOMs.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_BOMs.IsIndeterminate = true;
                Prog_BOMs.Value = 0;

                ConsoleWindow.Children.Add(Prog_BOMs);
            }));

            print("->\tConverting BOMs.csv", 10, "");

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            // Headers
            xlWorkSheet.Cells[1, 1] = "BOM_ItemCode";
            xlWorkSheet.Cells[1, 2] = "Revision";
            xlWorkSheet.Cells[1, 3] = "Quantity";
            xlWorkSheet.Cells[1, 4] = "Factor";
            xlWorkSheet.Cells[1, 5] = "Yield";
            xlWorkSheet.Cells[1, 6] = "YieldFormula";
            xlWorkSheet.Cells[1, 7] = "YieldItemsFormula";
            xlWorkSheet.Cells[1, 8] = "YieldCoproductsFormula";
            xlWorkSheet.Cells[1, 9] = "YieldScrapsFormula";
            xlWorkSheet.Cells[1, 10] = "Warehouse";
            xlWorkSheet.Cells[1, 11] = "DistRule";
            xlWorkSheet.Cells[1, 12] = "DistRule2";
            xlWorkSheet.Cells[1, 13] = "DistRule3";
            xlWorkSheet.Cells[1, 14] = "DistRule4";
            xlWorkSheet.Cells[1, 15] = "DistRule5";
            xlWorkSheet.Cells[1, 16] = "Project";
            xlWorkSheet.Cells[1, 17] = "BatchSize";
            xlWorkSheet.Cells[1, 18] = "ProdType";
            xlWorkSheet.Cells[1, 19] = "Instructions";

            //BOM_ItemCode
            xlWorkSheet.Cells[2, 1] = "'" + name;
            xlWorkSheet.Cells[3, 1] = "'" + name + "S-MAT";

            // Revision
            xlWorkSheet.Cells[2, 2] = "'00";
            xlWorkSheet.Cells[3, 2] = "'00";

            // Quantity
            xlWorkSheet.Cells[2, 3] = "1";
            xlWorkSheet.Cells[3, 3] = "1";

            // Factor
            xlWorkSheet.Cells[2, 4] = "1";
            xlWorkSheet.Cells[3, 4] = "1";

            // Yield
            xlWorkSheet.Cells[2, 5] = "100";
            xlWorkSheet.Cells[3, 5] = "100";

            //Warehouse
            xlWorkSheet.Cells[2, 10] = "FG";
            xlWorkSheet.Cells[3, 10] = "WIP-SUB";

            //BatchSize
            xlWorkSheet.Cells[2, 17] = "1";
            xlWorkSheet.Cells[3, 17] = "1";

            xlWorkSheet.Cells[2, 18] = "I";
            xlWorkSheet.Cells[3, 18] = "I";


            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "BOMs.csv");

            print("->\tSaving BOMs.csv", 10, "");
            
            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved BOMs.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save BOMs.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOMs.IsIndeterminate = false;
                Prog_BOMs.Visibility = Visibility.Collapsed;
            }));

            //dispatcherTimer.Stop();
        }

        // Convert 2
        private void Create_BOM_Items()
        {
            timer = 0;

            print("BOM_Items", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Items.Maximum = 100;
                Prog_BOM_Items.Height = 10;
                Prog_BOM_Items.Visibility = Visibility.Visible;
                Prog_BOM_Items.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Items.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Items.IsIndeterminate = true;
                Prog_BOM_Items.Value = 0;

                ConsoleWindow.Children.Add(Prog_BOM_Items);
            }));

            print("->\tConverting BOM_Items.csv", 10, "");

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

            xlWorkSheet.Cells[1, 1] = "BOM_ItemCode";
            xlWorkSheet.Cells[1, 2] = "Revision";
            xlWorkSheet.Cells[1, 3] = "Sequence";
            xlWorkSheet.Cells[1, 4] = "ItemCode";
            xlWorkSheet.Cells[1, 5] = "Item_Revision";
            xlWorkSheet.Cells[1, 6] = "Warehouse";
            xlWorkSheet.Cells[1, 7] = "Factor";
            xlWorkSheet.Cells[1, 8] = "FactorDesc";
            xlWorkSheet.Cells[1, 9] = "Quantity";
            xlWorkSheet.Cells[1, 10] = "ScrapPercent";
            xlWorkSheet.Cells[1, 11] = "Yield";
            xlWorkSheet.Cells[1, 12] = "IssueType";
            xlWorkSheet.Cells[1, 13] = "OcrCode";
            xlWorkSheet.Cells[1, 14] = "OcrCode2";
            xlWorkSheet.Cells[1, 15] = "OcrCode3";
            xlWorkSheet.Cells[1, 16] = "OcrCode4";
            xlWorkSheet.Cells[1, 17] = "OcrCode5";
            xlWorkSheet.Cells[1, 18] = "Project";
            xlWorkSheet.Cells[1, 19] = "SubcontractingItem";
            xlWorkSheet.Cells[1, 20] = "Remarks";
            xlWorkSheet.Cells[1, 21] = "Formula";

            int THTSequence = 10;
            int SMTSequence = 10;


            int position = 0;
            for (int i = 0; i < partCount; i++)
            {
                if (ImportExcelFile.ReadCell(i + 1, 7) == "") //THT
                {
                    xlWorkSheet.Cells[position + 2, 1] = "'" + name;
                    xlWorkSheet.Cells[position + 2, 3] = THTSequence.ToString();

                    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[position + 2, 4] = "'" + ImportExcelFile.ReadCell(i + 1, 2);
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[position + 2, 4] = "'" + ImportExcelFile.ReadCell(i + 1, 1);
                    }

                    xlWorkSheet.Cells[position + 2, 2] = "'00";
                    xlWorkSheet.Cells[position + 2, 5] = "'00";
                    xlWorkSheet.Cells[position + 2, 6] = "WIP";
                    xlWorkSheet.Cells[position + 2, 7] = "0";
                    xlWorkSheet.Cells[position + 2, 9] = "'" + ImportExcelFile.ReadCell(i + 1, 3);
                    xlWorkSheet.Cells[position + 2, 10] = "0";
                    xlWorkSheet.Cells[position + 2, 11] = "100";
                    xlWorkSheet.Cells[position + 2, 12] = "M";
                    xlWorkSheet.Cells[position + 2, 19] = "N";
                    xlWorkSheet.Cells[position + 2, 20] = "'" + ImportExcelFile.ReadCell(i + 1, 4);

                    THTSequence = THTSequence + 10;
                    position++;
                }
            }

            xlWorkSheet.Cells[position + 2, 1] = "'" + name;
            xlWorkSheet.Cells[position + 2, 3] = THTSequence.ToString();
            xlWorkSheet.Cells[position + 2, 4] = "'" + name + "S-MAT";
            xlWorkSheet.Cells[position + 2, 2] = "'00";
            xlWorkSheet.Cells[position + 2, 5] = "'00";
            xlWorkSheet.Cells[position + 2, 6] = "WIP-SUB";
            xlWorkSheet.Cells[position + 2, 7] = "0";
            xlWorkSheet.Cells[position + 2, 9] = "1";
            xlWorkSheet.Cells[position + 2, 10] = "0";
            xlWorkSheet.Cells[position + 2, 11] = "100";
            xlWorkSheet.Cells[position + 2, 12] = "M";
            xlWorkSheet.Cells[position + 2, 19] = "N";
            xlWorkSheet.Cells[position + 2, 20] = "";
            position++;



            for (int i = 0; i < partCount; i++)
            {
                if (ImportExcelFile.ReadCell(i + 1, 7) != "") //THT
                {
                    xlWorkSheet.Cells[position + 2, 1] = "'" + name + "S-MAT";
                    xlWorkSheet.Cells[position + 2, 3] = SMTSequence.ToString();

                    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[position + 2, 4] = ImportExcelFile.ReadCell(i + 1, 2);
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[position + 2, 4] = ImportExcelFile.ReadCell(i + 1, 1);
                    }

                    xlWorkSheet.Cells[position + 2, 2] = "'00";
                    xlWorkSheet.Cells[position + 2, 5] = "'00";
                    xlWorkSheet.Cells[position + 2, 6] = "WIP";
                    xlWorkSheet.Cells[position + 2, 7] = "0";
                    xlWorkSheet.Cells[position + 2, 9] = ImportExcelFile.ReadCell(i + 1, 3);
                    xlWorkSheet.Cells[position + 2, 10] = "0";
                    xlWorkSheet.Cells[position + 2, 11] = "100";
                    xlWorkSheet.Cells[position + 2, 12] = "M";
                    xlWorkSheet.Cells[position + 2, 19] = "N";
                    xlWorkSheet.Cells[position + 2, 20] = ImportExcelFile.ReadCell(i + 1, 4);

                    SMTSequence = SMTSequence + 10;
                    position++;
                }
            }






            //for (int i = 0; i < partCount; i++)
            //{
            //    if (ImportExcelFile.ReadCell(i + 1, 7) == "") //THT
            //    {
            //        xlWorkSheet.Cells[i + 2, 1] = "'" + name;
            //        xlWorkSheet.Cells[i + 2, 3] = THTSequence.ToString();

            //        THTSequence = THTSequence + 10;
            //    }
            //    else //SMT
            //    {
            //        xlWorkSheet.Cells[i + 2, 1] = "'" + name + "S-MAT";
            //        xlWorkSheet.Cells[i + 2, 3] = SMTSequence.ToString();

            //        SMTSequence = SMTSequence + 10;
            //    }

            //    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
            //    {
            //        // ItemCode = CPN
            //        xlWorkSheet.Cells[i + 2, 4] = ImportExcelFile.ReadCell(i + 1, 2);
            //    }
            //    else
            //    {
            //        // ItemCode = IPN
            //        xlWorkSheet.Cells[i + 2, 4] = ImportExcelFile.ReadCell(i + 1, 1);
            //    }

            //    xlWorkSheet.Cells[i + 2, 2] = "'00";
            //    xlWorkSheet.Cells[i + 2, 5] = "'00";
            //    xlWorkSheet.Cells[i + 2, 6] = "WIP";
            //    xlWorkSheet.Cells[i + 2, 7] = "0";
            //    xlWorkSheet.Cells[i + 2, 9] = ImportExcelFile.ReadCell(i + 1, 3);
            //    xlWorkSheet.Cells[i + 2, 10] = "0";
            //    xlWorkSheet.Cells[i + 2, 11] = "100";
            //    xlWorkSheet.Cells[i + 2, 12] = "M";
            //    xlWorkSheet.Cells[i + 2, 19] = "N";
            //    xlWorkSheet.Cells[i + 2, 20] = ImportExcelFile.ReadCell(i + 1, 4);
            //}






            // Adds the SMT-THT BOM Linkage
            

            //xlWorkSheet.Sort.SortFields.Add(xlWorkSheet.ra)

            //xlWorkSheet.Sort();

            ////////////
            // Saving //
            ////////////
            print("->\tSaving BOM_Items.csv", 10, "");

            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Items.csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved BOM_Items.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save BOM_Items.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Items.IsIndeterminate = false;
                Prog_BOM_Items.Visibility = Visibility.Collapsed;
            }));
        }

        // Convert 3
        private void Create_BOM_Scraps()
        {
            timer = 0;

            print("BOM_Scraps", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Scraps.Maximum = 100;
                Prog_BOM_Scraps.Height = 10;
                Prog_BOM_Scraps.Visibility = Visibility.Visible;
                Prog_BOM_Scraps.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Scraps.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Scraps.IsIndeterminate = true;
                Prog_BOM_Scraps.Value = 0;

                ConsoleWindow.Children.Add(Prog_BOM_Scraps);
            }));

            print("->\tConverting BOM_Scraps.csv", 10, "");

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

            xlWorkSheet.Cells[1, 1] = "BOM_ItemCode";
            xlWorkSheet.Cells[1, 2] = "Revision";
            xlWorkSheet.Cells[1, 3] = "Sequence";
            xlWorkSheet.Cells[1, 4] = "ItemCode";
            xlWorkSheet.Cells[1, 5] = "Item_Revision";
            xlWorkSheet.Cells[1, 6] = "Warehouse";
            xlWorkSheet.Cells[1, 7] = "Factor";
            xlWorkSheet.Cells[1, 8] = "FactorDesc";
            xlWorkSheet.Cells[1, 9] = "Quantity";
            xlWorkSheet.Cells[1, 10] = "Yield";
            xlWorkSheet.Cells[1, 11] = "Type";
            xlWorkSheet.Cells[1, 12] = "IssueType";
            xlWorkSheet.Cells[1, 13] = "OcrCode";
            xlWorkSheet.Cells[1, 14] = "OcrCode2";
            xlWorkSheet.Cells[1, 15] = "OcrCode3";
            xlWorkSheet.Cells[1, 16] = "OcrCode4";
            xlWorkSheet.Cells[1, 17] = "OcrCode5";
            xlWorkSheet.Cells[1, 18] = "Project";
            xlWorkSheet.Cells[1, 19] = "Remarks";
            xlWorkSheet.Cells[1, 20] = "Formula";

            print("->\tSaving BOM_Scraps.csv", 10, "");
            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Scraps.csv");


            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved BOM_Scraps.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save BOM_Scraps.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Scraps.IsIndeterminate = false;
                Prog_BOM_Scraps.Visibility = Visibility.Collapsed;
            }));
        }

        // Convert 4
        private void Create_BOM_Coproducts()
        {
            timer = 0;

            print("BOM_Coproducts", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Coproducts.Maximum = 100;
                Prog_BOM_Coproducts.Height = 10;
                Prog_BOM_Coproducts.Visibility = Visibility.Visible;
                Prog_BOM_Coproducts.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Coproducts.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_BOM_Coproducts.IsIndeterminate = true;
                Prog_BOM_Coproducts.Value = 0;

                ConsoleWindow.Children.Add(Prog_BOM_Coproducts);
            }));

            print("->\tConverting BOM_Coproducts.csv", 10, "");

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

            xlWorkSheet.Cells[1, 1] = "BOM_ItemCode";
            xlWorkSheet.Cells[1, 2] = "Revision";
            xlWorkSheet.Cells[1, 3] = "Sequence";
            xlWorkSheet.Cells[1, 4] = "ItemCode";
            xlWorkSheet.Cells[1, 5] = "Item_Revision";
            xlWorkSheet.Cells[1, 6] = "Warehouse";
            xlWorkSheet.Cells[1, 7] = "Factor";
            xlWorkSheet.Cells[1, 8] = "FactorDesc";
            xlWorkSheet.Cells[1, 9] = "Quantity";
            xlWorkSheet.Cells[1, 10] = "Yield";
            xlWorkSheet.Cells[1, 11] = "IssueType";
            xlWorkSheet.Cells[1, 12] = "OcrCode";
            xlWorkSheet.Cells[1, 13] = "OcrCode2";
            xlWorkSheet.Cells[1, 14] = "OcrCode3";
            xlWorkSheet.Cells[1, 15] = "OcrCode4";
            xlWorkSheet.Cells[1, 16] = "OcrCode5";
            xlWorkSheet.Cells[1, 17] = "Project";
            xlWorkSheet.Cells[1, 18] = "Remarks";
            xlWorkSheet.Cells[1, 19] = "Formula";


            print("->\tSaving BOM_Coproducts.csv", 10, "");
            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Coproducts.csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved BOM_Coproducts.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save BOM_Coproducts.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_BOM_Coproducts.IsIndeterminate = false;
                Prog_BOM_Coproducts.Visibility = Visibility.Collapsed;
            }));
        }

        // Convert 5
        private void Create_OITM()
        {
            timer = 0;

            print("OITM", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_OITM.Maximum = 100;
                Prog_OITM.Height = 10;
                Prog_OITM.Visibility = Visibility.Visible;
                Prog_OITM.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_OITM.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_OITM.IsIndeterminate = true;
                Prog_OITM.Value = 0;

                ConsoleWindow.Children.Add(Prog_OITM);
            }));

            print("->\tConverting OITM.csv", 10, "");

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

            // FIlls in Headings for OITM
            string line1 = "ItemCode;ItemName;ForeignName;ItemsGroupCode;CustomsGroupCode;SalesVATGroup;BarCode;VatLiable;PurchaseItem;SalesItem;InventoryItem;IncomeAccount;ExemptIncomeAccount;ExpanseAccount;Mainsupplier;SupplierCatalogNo;DesiredInventory;MinInventory;Picture;User_Text;SerialNum;CommissionPercent;CommissionSum;CommissionGroup;TreeType;AssetItem;DataExportCode;Manufacturer;ManageSerialNumbers;ManageBatchNumbers;Valid;ValidFrom;ValidTo;ValidRemarks;Frozen;FrozenFrom;FrozenTo;FrozenRemarks;SalesUnit;SalesItemsPerUnit;SalesPackagingUnit;SalesQtyPerPackUnit;SalesUnitLength;SalesLengthUnit;SalesUnitWidth;SalesWidthUnit;SalesUnitHeight;SalesHeightUnit;SalesUnitVolume;SalesVolumeUnit;SalesUnitWeight;SalesWeightUnit;PurchaseUnit;PurchaseItemsPerUnit;PurchasePackagingUnit;PurchaseQtyPerPackUnit;PurchaseUnitLength;PurchaseLengthUnit;PurchaseUnitWidth;PurchaseWidthUnit;PurchaseUnitHeight;PurchaseHeightUnit;PurchaseUnitVolume;PurchaseVolumeUnit;PurchaseUnitWeight;PurchaseWeightUnit;PurchaseVATGroup;SalesFactor1;SalesFactor2;SalesFactor3;SalesFactor4;PurchaseFactor1;PurchaseFactor2;PurchaseFactor3;PurchaseFactor4;ForeignRevenuesAccount;ECRevenuesAccount;ForeignExpensesAccount;ECExpensesAccount;AvgStdPrice;DefaultWarehouse;ShipType;GLMethod;TaxType;MaxInventory;ManageStockByWarehouse;PurchaseHeightUnit1;PurchaseUnitHeight1;PurchaseLengthUnit1;PurchaseUnitLength1;PurchaseWeightUnit1;PurchaseUnitWeight1;PurchaseWidthUnit1;PurchaseUnitWidth1;SalesHeightUnit1;SalesUnitHeight1;SalesLengthUnit1;SalesUnitLength1;SalesWeightUnit1;SalesUnitWeight1;SalesWidthUnit1;SalesUnitWidth1;ForceSelectionOfSerialNumber;ManageSerialNumbersOnReleaseOnly;WTLiable;CostAccountingMethod;SWW;WarrantyTemplate;IndirectTax;ArTaxCode;ApTaxCode;BaseUnitName;ItemCountryOrg;IssueMethod;SRIAndBatchManageMethod;IsPhantom;InventoryUOM;PlanningSystem;ProcurementMethod;ComponentWarehouse;OrderIntervals;OrderMultiple;LeadTime;MinOrderQuantity;ItemType;ItemClass;OutgoingServiceCode;IncomingServiceCode;ServiceGroup;NCMCode;MaterialType;MaterialGroup;ProductSource;Properties1;Properties2;Properties3;Properties4;Properties5;Properties6;Properties7;Properties8;Properties9;Properties10;Properties11;Properties12;Properties13;Properties14;Properties15;Properties16;Properties17;Properties18;Properties19;Properties20;Properties21;Properties22;Properties23;Properties24;Properties25;Properties26;Properties27;Properties28;Properties29;Properties30;Properties31;Properties32;Properties33;Properties34;Properties35;Properties36;Properties37;Properties38;Properties39;Properties40;Properties41;Properties42;Properties43;Properties44;Properties45;Properties46;Properties47;Properties48;Properties49;Properties50;Properties51;Properties52;Properties53;Properties54;Properties55;Properties56;Properties57;Properties58;Properties59;Properties60;Properties61;Properties62;Properties63;Properties64;AutoCreateSerialNumbersOnRelease;DNFEntry;GTSItemSpec;GTSItemTaxCategory;FuelID;BeverageTableCode;BeverageGroupCode;BeverageCommercialBrandCode;Series;ToleranceDays;TypeOfAdvancedRules;IssuePrimarilyBy;NoDiscounts;AssetClass;AssetGroup;InventoryNumber;Technician;Employee;Location;CapitalizationDate;StatisticalAsset;Cession;DeactivateAfterUsefulLife;UoMGroupEntry;InventoryUoMEntry;DefaultSalesUoMEntry;DefaultPurchasingUoMEntry;DepreciationGroup;AssetSerialNumber;InventoryWeight;InventoryWeightUnit;InventoryWeight1;InventoryWeightUnit1;DefaultCountingUnit;DefaultCountingUoMEntry;Excisable;ChapterID;ScsCode;SpProdType;ProdStdCost;InCostRollup;VirtualAssetItem;EnforceAssetSerialNumbers;AttachmentEntry;GSTRelevnt;SACEntry;GSTTaxCategory;ServiceCategoryEntry;CapitalGoodsOnHoldPercent;CapitalGoodsOnHoldLimit;AssessableValue;AssVal4WTR;SOIExcisable;TNVED;ImportedItem;PricingUnit;U_LicPlate;U_MaxOrdrQty;U_ILeadTime;U_SAAB_IC;U_REUTECH_IC;U_AIRBUS_IC;U_DENEL_IC;U_MARKING_IC;U_ALTERNATIVE_IC;U_MOUSER_IC;U_DIGIKEY_IC;U_CHARACTER_IC;U_PackSize;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode;U_InvLevFromItmDts;U_CTSRSerialization;U_BOY_TB_0";
            string line2 = "ItemCode;ItemName;FrgnName;ItmsGrpCod;CstGrpCode;VatGourpSa;CodeBars;VATLiable;PrchseItem;SellItem;InvntItem;IncomeAcct;ExmptIncom;ExpensAcct;CardCode;SuppCatNum;ReorderQty;MinLevel;PicturName;UserText;SerialNum;CommisPcnt;CommisSum;CommisGrp;TreeType;AssetItem;ExportCode;FirmCode;ManSerNum;ManBtchNum;validFor;validFrom;validTo;ValidComm;frozenFor;frozenFrom;frozenTo;FrozenComm;SalUnitMsr;NumInSale;SalPackMsr;SalPackUn;SLength1;SLen1Unit;SWidth1;SWdth1Unit;SHeight1;SHght1Unit;SVolume;SVolUnit;SWeight1;SWght1Unit;BuyUnitMsr;NumInBuy;PurPackMsr;PurPackUn;BLength1;BLen1Unit;BWidth1;BWdth1Unit;BHeight1;BHght1Unit;BVolume;BVolUnit;BWeight1;BWght1Unit;VatGroupPu;SalFactor1;SalFactor2;SalFactor3;SalFactor4;PurFactor1;PurFactor2;PurFactor3;PurFactor4;FrgnInAcct;ECInAcct;FrgnExpAcc;ECExpAcc;AvgPrice;DfltWH;ShipType;GLMethod;TaxType;MaxLevel;ByWh;BHght2Unit;BHeight2;BLen2Unit;Blength2;BWght2Unit;BWeight2;BWdth2Unit;BWidth2;SHght2Unit;SHeight2;SLen2Unit;Slength2;SWght2Unit;SWeight2;SWdth2Unit;SWidth2;BlockOut;ManOutOnly;WTLiable;EvalSystem;SWW;WarrntTmpl;IndirctTax;TaxCodeAR;TaxCodeAP;BaseUnit;CountryOrg;IssueMthd;MngMethod;Phantom;InvntryUom;PlaningSys;PrcrmntMtd;CompoWH;OrdrIntrvl;OrdrMulti;LeadTime;MinOrdrQty;ItemType;ItemClass;OSvcCode;ISvcCode;ServiceGrp;NCMCode;MatType;MatGrp;ProductSrc;QryGroup1;QryGroup2;QryGroup3;QryGroup4;QryGroup5;QryGroup6;QryGroup7;QryGroup8;QryGroup9;QryGroup10;QryGroup11;QryGroup12;QryGroup13;QryGroup14;QryGroup15;QryGroup16;QryGroup17;QryGroup18;QryGroup19;QryGroup20;QryGroup21;QryGroup22;QryGroup23;QryGroup24;QryGroup25;QryGroup26;QryGroup27;QryGroup28;QryGroup29;QryGroup30;QryGroup31;QryGroup32;QryGroup33;QryGroup34;QryGroup35;QryGroup36;QryGroup37;QryGroup38;QryGroup39;QryGroup40;QryGroup41;QryGroup42;QryGroup43;QryGroup44;QryGroup45;QryGroup46;QryGroup47;QryGroup48;QryGroup49;QryGroup50;QryGroup51;QryGroup52;QryGroup53;QryGroup54;QryGroup55;QryGroup56;QryGroup57;QryGroup58;QryGroup59;QryGroup60;QryGroup61;QryGroup62;QryGroup63;QryGroup64;ManOutOnly;DNFEntry;Spec;TaxCtg;FuelCode;BeverTblC;BeverGrpC;BeverTM;Series;ToleranDay;GLPickMeth;IssuePriBy;NoDiscount;AssetClass;AssetGroup;InventryNo;Technician;Employee;Location;CapitalizationDate;StatisticalAsset;Cession;DeactivateAfterUsefulLife;UoMGroupEntry;InventoryUoMEntry;DefaultSalesUoMEntry;DefaultPurchasingUoMEntry;DepreciationGroup;AssetSerialNumber;InventoryWeight;InventoryWeightUnit;InventoryWeight1;InventoryWeightUnit1;DefaultCountingUnit;DefaultCountingUoMEntry;Excisable;ChapterID;ScsCode;SpProdType;ProdStdCost;InCostRollup;VirtualAssetItem;EnforceAssetSerialNumbers;AttachmentEntry;GSTRelevnt;SACEntry;GSTTaxCategory;ServiceCategoryEntry;CapitalGoodsOnHoldPercent;CapitalGoodsOnHoldLimit;AssessableValue;AssVal4WTR;SOIExcisable;TNVED;ImportedItem;PricingUnit;U_LicPlate;U_MaxOrdrQty;U_ILeadTime;U_SAAB_IC;U_REUTECH_IC;U_AIRBUS_IC;U_DENEL_IC;U_MARKING_IC;U_ALTERNATIVE_IC;U_MOUSER_IC;U_DIGIKEY_IC;U_CHARACTER_IC;U_PackSize;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode;U_InvLevFromItmDts;U_CTSRSerialization;U_BOY_TB_0";

            string[] heading1 = line1.Split(';');
            string[] heading2 = line2.Split(';');

            for (int i = 1; i < heading1.Length+1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i-1];

            }

            for (int i = 1; i < heading2.Length+1; i++)
            {
                xlWorkSheet.Cells[2, i] = heading2[i-1];
            }

            for (int i = 0; i < partCount; i++)
            {
                if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                {
                    // ItemCode = CPN
                    xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell(i + 1, 2);
                }
                else
                {
                    // ItemCode = IPN
                    xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell(i + 1, 1);
                }

                xlWorkSheet.Cells[i + 3, 2] = ImportExcelFile.ReadCell(i + 1, 5);
                xlWorkSheet.Cells[i + 3, 4] = ImportExcelFile.ReadCell(i + 1, 8);
                xlWorkSheet.Cells[i + 3, 16] = ImportExcelFile.ReadCell(i + 1, 6);
                xlWorkSheet.Cells[i + 3, 29] = "N";
                xlWorkSheet.Cells[i + 3, 30] = "Y";
                xlWorkSheet.Cells[i + 3, 81] = "RCV";
                xlWorkSheet.Cells[i + 3, 106] = "B";
                xlWorkSheet.Cells[i + 3, 114] = "M";
                xlWorkSheet.Cells[i + 3, 115] = "A";
                xlWorkSheet.Cells[i + 3, 118] = "M";
            }


            print("->\tSaving OITM.csv", 10, "");

            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "OITM - Incomar (OMAD & OMAE).csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved OITM.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save OITM.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_OITM.IsIndeterminate = false;
                Prog_OITM.Visibility = Visibility.Collapsed;
            }));
        }

        // Convert 
        private void Create_OITW()
        {
            timer = 0;

            print("OITW", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_OITW.Maximum = 100;
                Prog_OITW.Height = 10;
                Prog_OITW.Visibility = Visibility.Visible;
                Prog_OITW.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_OITW.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_OITW.IsIndeterminate = true;
                Prog_OITW.Value = 0;

                ConsoleWindow.Children.Add(Prog_OITW);
            }));

            print("->\tConverting OITM.csv", 10, "");

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

            string line1 = "ParentKey;LineNum;MinimalStock;MaximalStock;MinimalOrder;StandardAveragePrice;Locked;InventoryAccount;CostAccount;TransferAccount;RevenuesAccount;VarienceAccount;DecreasingAccount;IncreasingAccount;ReturningAccount;ExpensesAccount;EURevenuesAccount;EUExpensesAccount;ForeignRevenueAcc;ForeignExpensAcc;ExemptIncomeAcc;PriceDifferenceAcc;WarehouseCode;ExpenseClearingAct;PurchaseCreditAcc;EUPurchaseCreditAcc;ForeignPurchaseCreditAcc;SalesCreditAcc;SalesCreditEUAcc;ExemptedCredits;SalesCreditForeignAcc;ExpenseOffsettingAccount;WipAccount;ExchangeRateDifferencesAcct;GoodsClearingAcct;NegativeInventoryAdjustmentAccount;CostInflationOffsetAccount;GLDecreaseAcct;GLIncreaseAcct;PAReturnAcct;PurchaseAcct;PurchaseOffsetAcct;ShippedGoodsAccount;StockInflationOffsetAccount;StockInflationAdjustAccount;VATInRevenueAccount;WipVarianceAccount;CostInflationAccount;WHIncomingCenvatAccount;WHOutgoingCenvatAccount;StockInTransitAccount;WipOffsetProfitAndLossAccount;InventoryOffsetProfitAndLossAccount;DefaultBin;DefaultBinEnforced;PurchaseBalanceAccount;U_RecAdjSuppAcct;U_TechAcctMC;U_SettCstsBT;U_AcctCstsBT;U_CostsAcct;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode";
            string line2 = "ItemCode;LineNum;MinStock;MaxStock;MinOrder;AvgPrice;Locked;BalInvntAc;SaleCostAc;TransferAc;RevenuesAc;VarianceAc;DecreasAc;IncreasAc;ReturnAc;ExpensesAc;EURevenuAc;EUExpensAc;FrRevenuAc;FrExpensAc;ExmptIncom;PriceDifAc;WhsCode;ExpClrAct;APCMAct;APCMEUAct;APCMFrnAct;ARCMAct;ARCMEUAct;ARCMExpAct;ARCMFrnAct;ExpOfstAct;WipAcct;ExchangeAc;BalanceAcc;NegStckAct;CstOffsAct;DecresGlAc;IncresGlAc;PAReturnAc;PurchaseAc;PurchOfsAc;ShpdGdsAct;StkOffsAct;StokRvlAct;VatRevAct;WipVarAcct;CostRvlAct;WhICenAct;WhOCenAct;StkInTnAct;WipOffset;StockOffst;DftBinAbs;DftBinEnfd;PurBalAct;ItemCode;U_RecAdjSuppAcct;U_TechAcctMC;U_SettCstsBT;U_AcctCstsBT;U_CostsAcct;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5";

            string[] heading1 = line1.Split(';');
            string[] heading2 = line2.Split(';');

            int i;
            for (i = 1; i < heading1.Length + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i - 1];

            }

            for (i = 1; i < heading2.Length + 1; i++)
            {
                xlWorkSheet.Cells[2, i] = heading2[i - 1];
            }

            
            string[] warehouseArray = { "MAIN", "WIP", warehouse, "RCV" };

            i = 0;
            for (int j = 0; j < 4; j++)
            {
                for (int k = 0; k < partCount; k++)
                {
                    if (ImportExcelFile.ReadCell(k + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell((k + 1), 2);
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell((k + 1), 1);
                    }

                    xlWorkSheet.Cells[i + 3, 23] = warehouseArray[j];

                    i++;
                }
            }


            //int startPoint = 0;
            //for (int i = 0; i < ImportExcelFile.rows; i++)
            //{
            //    if (ImportExcelFile.ReadCell(i, 2) == "IPN")
            //    {
            //        startPoint = i;
            //        break;
            //    }
            //}

            //int count = 0;
            //for (int i = 0; i < 4; i++)
            //{
            //    string whs = "";
            //    if (i == 0)
            //    {
            //        whs = "MAIN";
            //    }
            //    else if (i == 1)
            //    {
            //        whs = "WIP";
            //    }
            //    else if (i == 2)
            //    {
            //        whs = "RCV";
            //    }
            //    else
            //    {
            //        whs = "HAL";
            //    }
            //    Console.WriteLine("i: " + i.ToString());

            //    for (int j = 0; j < ImportExcelFile.rows; j++)
            //    {
            //        Console.WriteLine("j: " + j.ToString());
            //        //Console.WriteLine(ImportExcelFile.ReadCell(i, 2) + "|");


            //        //xlWorkSheet.Cells[i + 1 + 2, 0] = ImportExcelFile.ReadCell(i + startPoint + 1, 2).Replace(';', ':');

            //        if (ImportExcelFile.ReadCell(j + startPoint + 1, 1).Replace(';', ':') != "")
            //        {
            //            xlWorkSheet.Cells[count + 1+2, 1] = ImportExcelFile.ReadCell(j + startPoint + 1, 1).Replace(';', ':');
            //            xlWorkSheet.Cells[count + 1+2, 23] = whs;
            //            count++;

            //        }
            //    }
            //}

            print("->\tSaving OITW.csv", 10, "");
            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "OITW - Incomar (OMAD & OMAE).csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved OITW.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");

            }
            catch
            {
                printColor("->\tUnable to Save OITW.csv", 10, "", 255, 0, 0);
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_OITW.IsIndeterminate = false;
                Prog_OITW.Visibility = Visibility.Collapsed;
            }));
        }


        private void writeData(Excel.Worksheet ws, int newCol, string data, int startRow)
        {
            int startPoint = 0;
            for (int i = 0; i < ImportExcelFile.rows; i++)
            {
                if (ImportExcelFile.ReadCell(i , 1) == "POS")
                {
                    startPoint = i;
                    break;
                }
            }


            for (int i = startPoint; i < ImportExcelFile.rows; i++)
            {
                if (ImportExcelFile.ReadCell(i + startPoint + 1, 0).Replace(';', ':') != "")
                {
                    ws.Cells[i - startPoint + startRow + 2, newCol] = data;
                }
            }
        }


        private void copyData(Excel.Worksheet ws, int oldCol, string oldName, int newCol, string newName, int startRow)
        {
            int startPoint = 0;
            for(int i = 0; i< ImportExcelFile.rows; i++)
            {
                if(ImportExcelFile.ReadCell(i, oldCol) == oldName)
                {
                    startPoint = i;
                    break;
                }
            }

            for (int i = 0; i < ImportExcelFile.rows; i++)
            {
                Console.WriteLine(ImportExcelFile.ReadCell(i, oldCol) + "|");

                ws.Cells[i + startRow + 2, newCol] = ImportExcelFile.ReadCell(i+startPoint+1, oldCol).Replace(';', ':');
            }


            //for (int i = 0; i < ImportExcelFile.rows; i++)
            //{
            //    for (int j = 0; j < ImportExcelFile.cols; j++)
            //    {
            //        Console.Write(ImportExcelFile.ReadCell(i, j) + "|");
            //    }
            //    Console.WriteLine();
            //}
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
                tb.TextWrapping = TextWrapping.Wrap; 
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

        [Dispatched]
        private void printColor(string text, double size, string weight, int red, int green, int blue)
        {
            TextBlock tb = new TextBlock();
            tb.Text = text;
            Byte r = ((byte)red);
            Byte g = ((byte)green);
            Byte b = ((byte)blue);
            tb.Foreground = new System.Windows.Media.SolidColorBrush(Color.FromRgb(r, g, b));
            tb.FontSize = size;
            tb.TextWrapping = TextWrapping.Wrap;
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

        private void ScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (e.ExtentHeightChange == 0)
            {   // Content unchanged : user scroll event
                if (ScrollViewer.VerticalOffset == ScrollViewer.ScrollableHeight)
                {   // Scroll bar is in bottom
                    // Set auto-scroll mode
                    AutoScroll = true;
                }
                else
                {   // Scroll bar isn't in bottom
                    // Unset auto-scroll mode
                    AutoScroll = false;
                }
            }

            // Content scroll event : auto-scroll eventually
            if (AutoScroll && e.ExtentHeightChange != 0)
            {   // Content changed and auto-scroll mode set
                // Autoscroll
                ScrollViewer.ScrollToVerticalOffset(ScrollViewer.ExtentHeight);
            }
        }
    }
}
