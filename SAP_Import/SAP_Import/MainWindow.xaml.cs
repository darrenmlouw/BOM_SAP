/*
 * Author:          Darren Louw
 * Position:        Software Engineer
 * Company:         Omnigo (Pty) LTD
 * Date Started:    2022/01/12
 * Date Edited:     2022/02/03
 * 
 * If By Any Chance This Program is not Woring
 * Please Check that the Headings in the BOM Export are the Same as the Headers Sepcified in the CheckHeaders() Functions
 * All the Headers should be in the same order as Specified Below
 * POS - IPN - CPN - Quantity - RefDes - Description - MPNs - CPN Commodity Group - Classification - IPN (ALT)
 * 
 * POS:             Should be a line number (integer) starting at 1 and incrementing by 1 each row              -May NOT be Blank       (1 Num Per Cell)
 * IPN:             Should be the IPN that is Linked to that Part                                               -May be Blank           (1 Ipn Per Cell)
 * CPN:             Generated of Provided CPN                                                                   -May NOT be Blank       (1 CPN per Cell)
 * Quantity:        Quanity that is linked to the BOM                                                           -May NOT be Blank       (1 QTY per Cell)
 * RefDes:          Reference Designators linked to the BOM                                                     -Unknown                (Multiple RefDes per Cell - Seperated by ", ")
 * Description:     Description Associated to the Part/BOM Part                                                 -Unknown                (Should not Contain ";", This Program will change it to ":")
 * MPNs:            Manufacturing Part Numbers linked to the BOM Part                                           -Unknown                (Multiple - Seperated by " $ ")
 * CPN Com Grp:*    Footprint linked to the BOM Part                                                            -May Be Blank           (1 Footprint Per Cell - [Footprint = SMT] - [No Footprint = THT])
 * Classification:  Omnigos Group Codes Linked to the BOM Part                                                  -May NOT be Blank       (1 Class Per Cell)
 */

using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using PostSharp.Patterns.Threading;

namespace SAP_Import
{
    // Class that Holds Information Regarding the File that is Dragged into the Application
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

        //Constructor of the Class
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

        // Fuctions Reads a Cell in the Excel Worksheet at specific index
        public string ReadCell(int i, int j)
        {
            i++;
            j++;

            if (ws.Cells[i, j].Value2 != null)
            {
                return (ws.Cells[i, j].Value2).ToString().Replace(';', ':');
            }
            else
            {
                return "";
            }
        }

        //Releases Object Data
        public void Release()
        {
            Marshal.ReleaseComObject(this.ws);
            Marshal.ReleaseComObject(this.wb);
            Marshal.ReleaseComObject(this.excel);
        }
    }

    /// <summary>
    /// MainWindow that Displays all Information
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        // Variables Storing Relevant Information Regarding the Import File
        ExcelRead ImportExcelFile;
        private string path = "";
        private string filename = "";
        private int partCount = 0;

        // Variables Regarding the Progress/Flow of the Application
        private bool hasFile = false;
        private bool isPos = false;
        private bool isConverted = false;

        // Input Variables
        private string name = "";
        private string warehouse = "";
        
        // Progress Bars for Each Individual File that is Converted
        private ProgressBar Prog_BOMs = new ProgressBar();
        private ProgressBar Prog_BOM_Items = new ProgressBar();
        private ProgressBar Prog_BOM_Scraps = new ProgressBar();
        private ProgressBar Prog_BOM_Coproducts = new ProgressBar();
        private ProgressBar Prog_OITM = new ProgressBar();
        private ProgressBar Prog_OITW = new ProgressBar();
        private ProgressBar Prog_Substitutes = new ProgressBar();
        private ProgressBar Prog_SubstitutesBOMs = new ProgressBar();
        private ProgressBar Prog_SubstitutesRevisions = new ProgressBar();

        // Variables for Execution Time of the Program
        private DispatcherTimer dispatcherTimer;
        private DispatcherTimer dispatcherClock;
        private int timer = 0;
        private int totalTime = 0;

        // Other Variables
        private Boolean AutoScroll = true;

        // Constructor of the MainWindow View
        public MainWindow()
        {
            dispatcherTimer = new DispatcherTimer(DispatcherPriority.Send);
            dispatcherTimer.Tick += new EventHandler(dispatcherTimer_Tick);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, 100);
            dispatcherTimer.Start();

            dispatcherClock = new DispatcherTimer(DispatcherPriority.Send);
            dispatcherClock.Tick += new EventHandler(Clock_Timer);
            dispatcherClock.Interval = new TimeSpan(0, 0, 1);
            dispatcherClock.Start();

            InitializeComponent();
        }

        // Times the Execution Speed of the Application in increments of 100ms
        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            timer = timer + 100;
            totalTime = totalTime + 100;
            CommandManager.InvalidateRequerySuggested();
        }

        // Updates the Clock of the Application every 1s
        private void Clock_Timer(object sender, EventArgs e)
        {
            TimeBlock.Text = DateTime.Now.ToString();
            CommandManager.InvalidateRequerySuggested();
        }

        // Exit Button
        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        // Changes Block Colour a Blue-ish Colour
        private void DragBlock_DragOver(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 200, 255));
        }

        // Changes Block Colour a Grey-ish Colour
        private void DragBlock_DragLeave(object sender, DragEventArgs e)
        {
            DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 200, 200));
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }

            if (e.ClickCount == 2)
            {
                if (this.WindowState == WindowState.Maximized)
                {
                    this.WindowState = WindowState.Normal;
                }
                else
                {
                    this.WindowState = WindowState.Maximized;
                }
            }
        }

        // Checks for the .xlsx File Extensions
        // Changes Colour of Block Depending on the Success of the Drop
        // Stores Information Regaring the Dragged File (Filename, Path, Extension)
        [Background]
        private void DragBlock_Drop(object sender, DragEventArgs e)
        {
            ProgressStart();

            hasFile = false;

            Dispatcher.BeginInvoke((Action)(() =>
            {
                BomName.Text = "";
                Warehouse.Text = "";
                Convert_Back.IsEnabled = false;
                Status.Text = "Status";
            }));

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                filename = System.IO.Path.GetFileName(files[0]);
                path = System.IO.Path.GetFullPath(files[0]);
                string extension = System.IO.Path.GetExtension(files[0]);

                if (extension == ".xlsx")
                {
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
                    printColor("File Does Not Have .xlsx Extension", 12, "", 255, 100, 100);
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

        // Initiates the Conversion Process with Various Checks
        [Background]
        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            totalTime = 0;
            name = "";
            Dispatcher.Invoke((Action)(() =>
            {
                name = BomName.Text;
                warehouse = Warehouse.Text;
                Exit.IsEnabled = false;
            }));

            var regexItem = new Regex("^[A-Z]*$");

            if (hasFile == true && name != "" && warehouse.Length == 3 && regexItem.IsMatch(warehouse))
            {
                
                ImportExcelFile = new ExcelRead(path, 1, filename);

                Dispatcher.BeginInvoke((Action)(() =>
                {
                    ConsoleWindow.Children.Clear();
                    Convert_Back.IsEnabled = false;
                }));

                ProgressStart();

                isConverted = false;
                try
                {
                    ConvertBOM();
                    
                }
                catch
                {
                    isConverted = false;
                }

                ImportExcelFile.wb.Close(true, null, null);
                ImportExcelFile.excel.Quit();
                ImportExcelFile.Release();

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

                    Exit.IsEnabled = true;
                }));

                ProgressEnd();
            }
            else
            {
                if(name == "")
                {
                    printColor("Please Enter BOM Name", 12, "bold", 255, 100, 100);
                }

                if (warehouse.Length != 3)
                {
                    printColor("Please Enter a 3 Letter Warehouse Code", 12, "bold", 255, 100, 100);
                }

                if(!regexItem.IsMatch(warehouse))
                {
                    printColor("No Special Character or Number Allowed", 12, "bold", 255, 100, 100);
                }
            }
        }

        // Helper Function that Initiates the Different Coverting Functions
        private void ConvertBOM()
        {
            isPos = false;
            CheckHeader();

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
                    Create_Substitutes();
                    Create_SubstitutesBOMs();
                    Create_SubstitutesRevisions();
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
        private void CheckHeader()
        {
            partCount = 0;

            // Change the Headings here, in the Same Order as in the Exported BOM File
            if (ImportExcelFile.ReadCell(0, 0) == "POS" &&
                ImportExcelFile.ReadCell(0, 1) == "IPN" &&
                ImportExcelFile.ReadCell(0, 2) == "CPN" &&
                ImportExcelFile.ReadCell(0, 3) == "Quantity" &&
                ImportExcelFile.ReadCell(0, 4) == "RefDes" &&
                ImportExcelFile.ReadCell(0, 5) == "Description" &&
                ImportExcelFile.ReadCell(0, 6) == "MPNs" &&
                ImportExcelFile.ReadCell(0, 7) == "CPN Commodity Group" &&
                ImportExcelFile.ReadCell(0, 8) == "Classification" &&
                ImportExcelFile.ReadCell(0, 9) == "IPN (ALT)")
            {
                isPos = true;

                for (int i = 0; i < ImportExcelFile.rows; i++)
                {
                    int number = 0;

                    if (int.TryParse(ImportExcelFile.ReadCell(i, 0), out number))
                    {
                        partCount++;
                    }
                }
            }
            else
            {
                if (ImportExcelFile.ReadCell(0, 0) != "POS")
                {
                    print("Unable to find \"POS\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 0) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 1) != "IPN")
                {
                    print("Unable to find \"IPN\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 1) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 2) != "CPN")
                {
                    print("Unable to find \"CPN\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 2) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 3) != "Quantity")
                {
                    print("Unable to find \"Quantity\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 3) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 4) != "RefDes")
                {
                    print("Unable to find \"RefDes\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 4) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 5) != "Description")
                {
                    print("Unable to find \"Description\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 5) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 6) != "MPNs")
                {
                    print("Unable to find \"MPNs\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 6) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 7) != "CPN Commodity Group")
                {
                    print("Unable to find \"CPN Commodity Group\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 7) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 8) != "Classification")
                {
                    print("Unable to find \"Classification\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 8) + "\" Instead", 10, "", 255, 0, 0);
                }
                if (ImportExcelFile.ReadCell(0, 9) != "IPN (ALT)")
                {
                    print("Unable to find \"IPN (ALT)\"", 10, "");
                    printColor("Found \"" + ImportExcelFile.ReadCell(0, 9) + "\" Instead", 10, "", 255, 0, 0);
                }
            }
                

        }

        // Conversion File 1
        private void Create_BOMs()
        {
            timer = 0;

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
        }

        // Conversion File 2
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
                    xlWorkSheet.Cells[position + 2, 1] = "'" + name.Replace(';', ':');
                    xlWorkSheet.Cells[position + 2, 3] = THTSequence.ToString();

                    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[position + 2, 4] = "'" + ImportExcelFile.ReadCell(i + 1, 2).Replace(';', ':');
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[position + 2, 4] = "'" + ImportExcelFile.ReadCell(i + 1, 1).Replace(';', ':');
                    }

                    xlWorkSheet.Cells[position + 2, 2] = "'00";
                    xlWorkSheet.Cells[position + 2, 5] = "'00";
                    xlWorkSheet.Cells[position + 2, 6] = "WIP";
                    xlWorkSheet.Cells[position + 2, 7] = "0";
                    xlWorkSheet.Cells[position + 2, 9] = "'" + ImportExcelFile.ReadCell(i + 1, 3).Replace(';', ':');
                    xlWorkSheet.Cells[position + 2, 10] = "0";
                    xlWorkSheet.Cells[position + 2, 11] = "100";
                    xlWorkSheet.Cells[position + 2, 12] = "M";
                    xlWorkSheet.Cells[position + 2, 19] = "N";
                    xlWorkSheet.Cells[position + 2, 20] = "'" + ImportExcelFile.ReadCell(i + 1, 4).Replace(';', ':');

                    THTSequence = THTSequence + 10;
                    position++;
                }
            }

            xlWorkSheet.Cells[position + 2, 1] = "'" + name.Replace(';', ':');
            xlWorkSheet.Cells[position + 2, 3] = THTSequence.ToString();
            xlWorkSheet.Cells[position + 2, 4] = "'" + name.Replace(';', ':') + "S-MAT";
            xlWorkSheet.Cells[position + 2, 2] = "'00";
            xlWorkSheet.Cells[position + 2, 5] = "'00";
            xlWorkSheet.Cells[position + 2, 6] = "WIP-SUB";
            //Wastage 0 for THT
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
                if (ImportExcelFile.ReadCell(i + 1, 7) != "") //SMT
                {
                    xlWorkSheet.Cells[position + 2, 1] = "'" + name.Replace(';', ':') + "S-MAT";
                    xlWorkSheet.Cells[position + 2, 3] = SMTSequence.ToString();

                    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[position + 2, 4] = ImportExcelFile.ReadCell(i + 1, 2).Replace(';', ':');
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[position + 2, 4] = ImportExcelFile.ReadCell(i + 1, 1).Replace(';', ':');
                    }

                    xlWorkSheet.Cells[position + 2, 2] = "'00";
                    xlWorkSheet.Cells[position + 2, 5] = "'00";
                    xlWorkSheet.Cells[position + 2, 6] = "WIP";

                    //Wastage # for SMT
                    string splitWastage = ImportExcelFile.ReadCell(i + 1, 7).Split(' ', 'C')[1];

                    xlWorkSheet.Cells[position + 2, 7] = splitWastage;
                    xlWorkSheet.Cells[position + 2, 9] = ImportExcelFile.ReadCell(i + 1, 3).Replace(';', ':');
                    xlWorkSheet.Cells[position + 2, 10] = "0";
                    xlWorkSheet.Cells[position + 2, 11] = "100";
                    xlWorkSheet.Cells[position + 2, 12] = "M";
                    xlWorkSheet.Cells[position + 2, 19] = "N";
                    xlWorkSheet.Cells[position + 2, 20] = ImportExcelFile.ReadCell(i + 1, 4).Replace(';', ':');

                    SMTSequence = SMTSequence + 10;
                    position++;
                }
            }

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

        // Conversion File 3
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

        // Conversion File 4
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

        // Conversion File 5
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
            string line1 = "ItemCode;ItemName;ForeignName;ItemsGroupCode;CustomsGroupCode;SalesVATGroup;BarCode;VatLiable;PurchaseItem;SalesItem;InventoryItem;IncomeAccount;ExemptIncomeAccount;ExpanseAccount;Mainsupplier;SupplierCatalogNo;DesiredInventory;MinInventory;Picture;User_Text;SerialNum;CommissionPercent;CommissionSum;CommissionGroup;TreeType;AssetItem;DataExportCode;Manufacturer;ManageSerialNumbers;ManageBatchNumbers;Valid;ValidFrom;ValidTo;ValidRemarks;Frozen;FrozenFrom;FrozenTo;FrozenRemarks;SalesUnit;SalesItemsPerUnit;SalesPackagingUnit;SalesQtyPerPackUnit;SalesUnitLength;SalesLengthUnit;SalesUnitWidth;SalesWidthUnit;SalesUnitHeight;SalesHeightUnit;SalesUnitVolume;SalesVolumeUnit;SalesUnitWeight;SalesWeightUnit;PurchaseUnit;PurchaseItemsPerUnit;PurchasePackagingUnit;PurchaseQtyPerPackUnit;PurchaseUnitLength;PurchaseLengthUnit;PurchaseUnitWidth;PurchaseWidthUnit;PurchaseUnitHeight;PurchaseHeightUnit;PurchaseUnitVolume;PurchaseVolumeUnit;PurchaseUnitWeight;PurchaseWeightUnit;PurchaseVATGroup;SalesFactor1;SalesFactor2;SalesFactor3;SalesFactor4;PurchaseFactor1;PurchaseFactor2;PurchaseFactor3;PurchaseFactor4;ForeignRevenuesAccount;ECRevenuesAccount;ForeignExpensesAccount;ECExpensesAccount;AvgStdPrice;DefaultWarehouse;ShipType;GLMethod;TaxType;MaxInventory;ManageStockByWarehouse;PurchaseHeightUnit1;PurchaseUnitHeight1;PurchaseLengthUnit1;PurchaseUnitLength1;PurchaseWeightUnit1;PurchaseUnitWeight1;PurchaseWidthUnit1;PurchaseUnitWidth1;SalesHeightUnit1;SalesUnitHeight1;SalesLengthUnit1;SalesUnitLength1;SalesWeightUnit1;SalesUnitWeight1;SalesWidthUnit1;SalesUnitWidth1;ForceSelectionOfSerialNumber;ManageSerialNumbersOnReleaseOnly;WTLiable;CostAccountingMethod;SWW;WarrantyTemplate;IndirectTax;ArTaxCode;ApTaxCode;BaseUnitName;ItemCountryOrg;IssueMethod;SRIAndBatchManageMethod;IsPhantom;InventoryUOM;PlanningSystem;ProcurementMethod;ComponentWarehouse;OrderIntervals;OrderMultiple;LeadTime;MinOrderQuantity;ItemType;ItemClass;OutgoingServiceCode;IncomingServiceCode;ServiceGroup;NCMCode;MaterialType;MaterialGroup;ProductSource;Properties1;Properties2;Properties3;Properties4;Properties5;Properties6;Properties7;Properties8;Properties9;Properties10;Properties11;Properties12;Properties13;Properties14;Properties15;Properties16;Properties17;Properties18;Properties19;Properties20;Properties21;Properties22;Properties23;Properties24;Properties25;Properties26;Properties27;Properties28;Properties29;Properties30;Properties31;Properties32;Properties33;Properties34;Properties35;Properties36;Properties37;Properties38;Properties39;Properties40;Properties41;Properties42;Properties43;Properties44;Properties45;Properties46;Properties47;Properties48;Properties49;Properties50;Properties51;Properties52;Properties53;Properties54;Properties55;Properties56;Properties57;Properties58;Properties59;Properties60;Properties61;Properties62;Properties63;Properties64;AutoCreateSerialNumbersOnRelease;DNFEntry;GTSItemSpec;GTSItemTaxCategory;FuelID;BeverageTableCode;BeverageGroupCode;BeverageCommercialBrandCode;Series;ToleranceDays;TypeOfAdvancedRules;IssuePrimarilyBy;NoDiscounts;AssetClass;AssetGroup;InventoryNumber;Technician;Employee;Location;CapitalizationDate;StatisticalAsset;Cession;DeactivateAfterUsefulLife;UoMGroupEntry;InventoryUoMEntry;DefaultSalesUoMEntry;DefaultPurchasingUoMEntry;DepreciationGroup;AssetSerialNumber;InventoryWeight;InventoryWeightUnit;InventoryWeight1;InventoryWeightUnit1;DefaultCountingUnit;DefaultCountingUoMEntry;Excisable;ChapterID;ScsCode;SpProdType;ProdStdCost;InCostRollup;VirtualAssetItem;EnforceAssetSerialNumbers;AttachmentEntry;GSTRelevnt;SACEntry;GSTTaxCategory;ServiceCategoryEntry;CapitalGoodsOnHoldPercent;CapitalGoodsOnHoldLimit;AssessableValue;AssVal4WTR;SOIExcisable;TNVED;ImportedItem;PricingUnit;U_LicPlate;U_MaxOrdrQty;U_ILeadTime;U_SAAB_IC;U_REUTECH_IC;U_AIRBUS_IC;U_DENEL_IC;U_MARKING_IC;U_ALTERNATIVE_IC;U_MOUSER_IC;U_DIGIKEY_IC;U_CHARACTER_IC;U_PackSize;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode;U_InvLevFromItmDts;U_CTSRSerialization;U_BOY_TB_0;U_MPNs";
            string line2 = "ItemCode;ItemName;FrgnName;ItmsGrpCod;CstGrpCode;VatGourpSa;CodeBars;VATLiable;PrchseItem;SellItem;InvntItem;IncomeAcct;ExmptIncom;ExpensAcct;CardCode;SuppCatNum;ReorderQty;MinLevel;PicturName;UserText;SerialNum;CommisPcnt;CommisSum;CommisGrp;TreeType;AssetItem;ExportCode;FirmCode;ManSerNum;ManBtchNum;validFor;validFrom;validTo;ValidComm;frozenFor;frozenFrom;frozenTo;FrozenComm;SalUnitMsr;NumInSale;SalPackMsr;SalPackUn;SLength1;SLen1Unit;SWidth1;SWdth1Unit;SHeight1;SHght1Unit;SVolume;SVolUnit;SWeight1;SWght1Unit;BuyUnitMsr;NumInBuy;PurPackMsr;PurPackUn;BLength1;BLen1Unit;BWidth1;BWdth1Unit;BHeight1;BHght1Unit;BVolume;BVolUnit;BWeight1;BWght1Unit;VatGroupPu;SalFactor1;SalFactor2;SalFactor3;SalFactor4;PurFactor1;PurFactor2;PurFactor3;PurFactor4;FrgnInAcct;ECInAcct;FrgnExpAcc;ECExpAcc;AvgPrice;DfltWH;ShipType;GLMethod;TaxType;MaxLevel;ByWh;BHght2Unit;BHeight2;BLen2Unit;Blength2;BWght2Unit;BWeight2;BWdth2Unit;BWidth2;SHght2Unit;SHeight2;SLen2Unit;Slength2;SWght2Unit;SWeight2;SWdth2Unit;SWidth2;BlockOut;ManOutOnly;WTLiable;EvalSystem;SWW;WarrntTmpl;IndirctTax;TaxCodeAR;TaxCodeAP;BaseUnit;CountryOrg;IssueMthd;MngMethod;Phantom;InvntryUom;PlaningSys;PrcrmntMtd;CompoWH;OrdrIntrvl;OrdrMulti;LeadTime;MinOrdrQty;ItemType;ItemClass;OSvcCode;ISvcCode;ServiceGrp;NCMCode;MatType;MatGrp;ProductSrc;QryGroup1;QryGroup2;QryGroup3;QryGroup4;QryGroup5;QryGroup6;QryGroup7;QryGroup8;QryGroup9;QryGroup10;QryGroup11;QryGroup12;QryGroup13;QryGroup14;QryGroup15;QryGroup16;QryGroup17;QryGroup18;QryGroup19;QryGroup20;QryGroup21;QryGroup22;QryGroup23;QryGroup24;QryGroup25;QryGroup26;QryGroup27;QryGroup28;QryGroup29;QryGroup30;QryGroup31;QryGroup32;QryGroup33;QryGroup34;QryGroup35;QryGroup36;QryGroup37;QryGroup38;QryGroup39;QryGroup40;QryGroup41;QryGroup42;QryGroup43;QryGroup44;QryGroup45;QryGroup46;QryGroup47;QryGroup48;QryGroup49;QryGroup50;QryGroup51;QryGroup52;QryGroup53;QryGroup54;QryGroup55;QryGroup56;QryGroup57;QryGroup58;QryGroup59;QryGroup60;QryGroup61;QryGroup62;QryGroup63;QryGroup64;ManOutOnly;DNFEntry;Spec;TaxCtg;FuelCode;BeverTblC;BeverGrpC;BeverTM;Series;ToleranDay;GLPickMeth;IssuePriBy;NoDiscount;AssetClass;AssetGroup;InventryNo;Technician;Employee;Location;CapitalizationDate;StatisticalAsset;Cession;DeactivateAfterUsefulLife;UoMGroupEntry;InventoryUoMEntry;DefaultSalesUoMEntry;DefaultPurchasingUoMEntry;DepreciationGroup;AssetSerialNumber;InventoryWeight;InventoryWeightUnit;InventoryWeight1;InventoryWeightUnit1;DefaultCountingUnit;DefaultCountingUoMEntry;Excisable;ChapterID;ScsCode;SpProdType;ProdStdCost;InCostRollup;VirtualAssetItem;EnforceAssetSerialNumbers;AttachmentEntry;GSTRelevnt;SACEntry;GSTTaxCategory;ServiceCategoryEntry;CapitalGoodsOnHoldPercent;CapitalGoodsOnHoldLimit;AssessableValue;AssVal4WTR;SOIExcisable;TNVED;ImportedItem;PricingUnit;U_LicPlate;U_MaxOrdrQty;U_ILeadTime;U_SAAB_IC;U_REUTECH_IC;U_AIRBUS_IC;U_DENEL_IC;U_MARKING_IC;U_ALTERNATIVE_IC;U_MOUSER_IC;U_DIGIKEY_IC;U_CHARACTER_IC;U_PackSize;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode;U_InvLevFromItmDts;U_CTSRSerialization;U_BOY_TB_0;U_MPNs";

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
                    xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell(i + 1, 2).Replace(';', ':');
                }
                else
                {
                    // ItemCode = IPN
                    xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell(i + 1, 1).Replace(';', ':');
                }


                //ItemName
                xlWorkSheet.Cells[i + 3, 2] = ImportExcelFile.ReadCell(i + 1, 5).Replace(';', ':');
                //ItemGroupCode
                xlWorkSheet.Cells[i + 3, 4] = ImportExcelFile.ReadCell(i + 1, 8).Replace(';', ':');
                //MPN
                xlWorkSheet.Cells[i + 3, heading1.Length] = "$" + ImportExcelFile.ReadCell(i + 1, 6).Replace(';', ':').Replace(" $ ", " $");
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
            string temp = savePath.Replace(ImportExcelFile.filename, "OITM.csv");

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

        // Conversion File 6
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
                        xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell((k + 1), 2).Replace(';', ':');
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[i + 3, 1] = ImportExcelFile.ReadCell((k + 1), 1).Replace(';', ':');
                    }

                    xlWorkSheet.Cells[i + 3, 23] = "'" + warehouseArray[j].Replace(';', ':');

                    i++;
                }
            }

            print("->\tSaving OITW.csv", 10, "");
            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "OITW.csv");

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

        // Conversion File 7
        private void Create_Substitutes()
        {
            timer = 0;

            print("Substitutes", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_Substitutes.Maximum = 100;
                Prog_Substitutes.Height = 10;
                Prog_Substitutes.Visibility = Visibility.Visible;
                Prog_Substitutes.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_Substitutes.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_Substitutes.IsIndeterminate = true;
                Prog_Substitutes.Value = 0;

                ConsoleWindow.Children.Add(Prog_Substitutes);
            }));

            print("->\tConverting Substitutes.csv", 10, "");

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

            // Fill in Headings for Substitutes
            string line1 = "ImportKey1;Code;Remarks";
            string[] heading1 = line1.Split(';');

            for (int i = 1; i < heading1.Length + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i - 1];
            }

            int ImportKey1 = 1;
            for (int i = 0; i < partCount; i++)
            {

                if(ImportExcelFile.ReadCell(i + 1, 9) != "")
                {
                    //ImportKey1
                    xlWorkSheet.Cells[ImportKey1 -1 + 2, 1] = ImportKey1;

                    //Code
                    if (ImportExcelFile.ReadCell(i + 1, 1) == "")
                    {
                        // ItemCode = CPN
                        xlWorkSheet.Cells[ImportKey1 - 1 + 2, 2] = "'" + ImportExcelFile.ReadCell(i + 1, 2).Replace(';', ':');
                    }
                    else
                    {
                        // ItemCode = IPN
                        xlWorkSheet.Cells[ImportKey1 - 1 + 2, 2] = "'" + ImportExcelFile.ReadCell(i + 1, 1).Replace(';', ':');
                    }
                    


                    ImportKey1++;
                }
            }


            print("->\tSaving Substitutes.csv", 10, "");

            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "Substitutes.csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved Substitutes.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save Substitutes.csv", 10, "", 255, 0, 0);
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
                Prog_Substitutes.IsIndeterminate = false;
                Prog_Substitutes.Visibility = Visibility.Collapsed;
            }));
        }

        // Conversion File 8
        private void Create_SubstitutesBOMs()
        {
            timer = 0;

            print("SubstitutesBOMs", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_SubstitutesBOMs.Maximum = 100;
                Prog_SubstitutesBOMs.Height = 10;
                Prog_SubstitutesBOMs.Visibility = Visibility.Visible;
                Prog_SubstitutesBOMs.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_SubstitutesBOMs.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_SubstitutesBOMs.IsIndeterminate = true;
                Prog_SubstitutesBOMs.Value = 0;

                ConsoleWindow.Children.Add(Prog_SubstitutesBOMs);
            }));

            print("->\tConverting SubstitutesBOMs.csv", 10, "");

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

            // Fill in Headings for SubstitutesBOMs
            string line1 = "ImportKey1;ImportKey2;BomCode;BomRevCode;DisableSubs;ValidFrom;ValidTo;BomRemarks;Remarks";
            string[] heading1 = line1.Split(';');

            for (int i = 1; i < heading1.Length + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i - 1];
            }


            print("->\tSaving SubstitutesBOMs.csv", 10, "");

            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "SubstitutesBOMs.csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved SubstitutesBOMs.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save SubstitutesBOMs.csv", 10, "", 255, 0, 0);
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
                Prog_SubstitutesBOMs.IsIndeterminate = false;
                Prog_SubstitutesBOMs.Visibility = Visibility.Collapsed;
            }));
        }

        // Conversion File 9
        private void Create_SubstitutesRevisions()
        {
            timer = 0;

            print("SubstitutesRevisions", 10, "bold");

            Dispatcher.Invoke((Action)(() =>
            {
                Prog_SubstitutesRevisions.Maximum = 100;
                Prog_SubstitutesRevisions.Height = 10;
                Prog_SubstitutesRevisions.Visibility = Visibility.Visible;
                Prog_SubstitutesRevisions.HorizontalAlignment = HorizontalAlignment.Stretch;
                Prog_SubstitutesRevisions.HorizontalContentAlignment = HorizontalAlignment.Stretch;
                Prog_SubstitutesRevisions.IsIndeterminate = true;
                Prog_SubstitutesRevisions.Value = 0;

                ConsoleWindow.Children.Add(Prog_SubstitutesRevisions);
            }));

            print("->\tConverting SubstitutesRevisions.csv", 10, "");

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

            // Fill in Headings for SubstitutesRevisions
            string line1 = "ImportKey1;ImportKey2;Revision;SItemCode;SRevision;Default;ValidFrom;ValidTo;Ratio;RplItm;RplCp;RplSc;Remarks";
            string[] heading1 = line1.Split(';');

            for (int i = 1; i < heading1.Length + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i - 1];
            }


            int ImportKey1 = 1;
            int lineNumber = 1;
            for (int i = 0; i < partCount; i++)
            {

                if (ImportExcelFile.ReadCell(i + 1, 9) != "")
                {
                    for (int j = 0; j < ImportExcelFile.ReadCell(i + 1, 9).Split('$').Length; j++)
                    {
                        //ImportKey1
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 1] = ImportKey1;
                        //ImportKey1      
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 2] = j + 1;
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 3] = "'00";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 4] = ImportExcelFile.ReadCell(i + 1, 9).Split('$')[j].Replace(';', ':');
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 5] = "'00";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 6] = "Y";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 9] = "1";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 10] = "Y";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 11] = "N";
                        xlWorkSheet.Cells[lineNumber - 1 + 2, 12] = "N";


                        lineNumber++;
                    }

                    ImportKey1++;
                }
            }


            print("->\tSaving SubstitutesRevisions.csv", 10, "");

            string savePath = "";
            savePath = ImportExcelFile.path;
            string temp = savePath.Replace(ImportExcelFile.filename, "SubstitutesRevisions.csv");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
                print("->\tSaved SubstitutesRevisions.csv", 10, "");
                print("->\tConversion Time: " + timer.ToString() + "ms", 10, "");
                print(temp, 10, "");
                print("line", 0, "");
            }
            catch
            {
                printColor("->\tUnable to Save SubstitutesRevisions.csv", 10, "", 255, 0, 0);
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
                Prog_SubstitutesRevisions.IsIndeterminate = false;
                Prog_SubstitutesRevisions.Visibility = Visibility.Collapsed;
            }));
        }

        // Function Not in Use
        private void WriteData(Excel.Worksheet ws, int newCol, string data, int startRow)
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

        // Function Not in Use
        private void CopyData(Excel.Worksheet ws, int oldCol, string oldName, int newCol, string newName, int startRow)
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
                ws.Cells[i + startRow + 2, newCol] = ImportExcelFile.ReadCell(i+startPoint+1, oldCol).Replace(';', ':');
            }
        }

        // Print to the In-App Console - No Colour
        [Dispatched]
        private void print(string text, double size, string weight)
        {
            // If Text = "line" then the In-App Console Prints an h-line Across the Console
            if (text == "line")
            {
                Border line = new Border();
                line.Height = 1;
                line.Background = new System.Windows.Media.SolidColorBrush(Color.FromRgb(150, 150, 150));
                line.Margin = new Thickness(0, 1, 0, 1);
                line.CornerRadius = new CornerRadius(0);
                ConsoleWindow.Children.Add(line);
            }
            else // Prints Normal Data onto the In-App Console
            {
                // Creates new TextBlock to add to the In-App Console
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

        //Print to the In-App Console - With Colour
        [Dispatched]
        private void printColor(string text, double size, string weight, int red, int green, int blue)
        {
            // Convert (int) red, green, blue to byte values
            Byte r = ((byte)red);
            Byte g = ((byte)green);
            Byte b = ((byte)blue);

            // Creates new TextBlock to add to the In-App Console
            TextBlock tb = new TextBlock();
            tb.Text = text;
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

        // Allows the In-App Console Window to Automatically Scroll when Data is Added to the Bottom
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
