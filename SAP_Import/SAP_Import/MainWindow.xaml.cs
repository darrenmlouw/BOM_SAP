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

        public MainWindow()
        {
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

        [Background]
        private void Convert_Click(object sender, RoutedEventArgs e)
        {
            name = "";
            Dispatcher.Invoke((Action)(() =>
            {
                name = BomName.Text;
            }));

            Console.WriteLine(name);

            if (hasFile == true && name != "")
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
                        DragBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 233, 233, 233));
                        OutBlock.Background = new System.Windows.Media.SolidColorBrush(Color.FromArgb(153, 200, 255, 200));
                    }
                    else
                    {
                        Status.Text = "Failed";
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
            }
        }

        //[Background]
        private void ConvertBOM()
        {
           

            isPos = false;
            CountPOS();

            if (isPos == true)
            {
                printColor("Correct Header Found", 12, "", 100, 155, 100);

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
                printColor("Unable to Find the Correct Headers", 12, "bold", 255, 100, 100);
            }
        }

        private void CountPOS()
        {
            if(ImportExcelFile.ReadCell(0, 0) == "POS")
            {
                isPos = true;
            }

        }

        //[Background]
        private void Create_BOMs()
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
            xlWorkSheet.Cells[2, 1] = name;
            xlWorkSheet.Cells[3, 1] = name + "S-MAT";

            // Revision
            xlWorkSheet.Cells[2, 2] = "'00";
            xlWorkSheet.Cells[3, 2] = "'00";
            //xlWorkSheet.Cells[4, 2] = "'00";

            // Quantity
            xlWorkSheet.Cells[2, 3] = "1";
            xlWorkSheet.Cells[3, 3] = "1";
            //xlWorkSheet.Cells[4, 3] = "1";

            // Factor
            xlWorkSheet.Cells[2, 4] = "1";
            xlWorkSheet.Cells[3, 4] = "1";
            //xlWorkSheet.Cells[4, 4] = "1";

            // Yield
            xlWorkSheet.Cells[2, 5] = "100";
            xlWorkSheet.Cells[3, 5] = "100";
            //xlWorkSheet.Cells[4, 5] = "100";

            //Warehouse
            xlWorkSheet.Cells[2, 10] = "FG";
            xlWorkSheet.Cells[3, 10] = "FG";
            //xlWorkSheet.Cells[4, 10] = "FG";

            //BatchSize
            xlWorkSheet.Cells[2, 17] = "1";
            xlWorkSheet.Cells[3, 17] = "1";
            //xlWorkSheet.Cells[4, 17] = "1";

            xlWorkSheet.Cells[2, 18] = "I";
            xlWorkSheet.Cells[3, 18] = "I";
            //xlWorkSheet.Cells[4, 18] = "I";






            string savePath = "";

            savePath = ImportExcelFile.path;

            string temp = savePath.Replace(ImportExcelFile.filename, "BOMs.csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
            }
            catch
            {
                MessageBox.Show("Unable to Save File");
            }
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }


        private void Create_BOM_Items()
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

            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            int THTCount = 0;
            int SMTCount = 0;

            //for (int i = 0; i < ImportExcelFile.rows; i++)
            //{

            //    xlWorkSheet.Cells[i, 1] = "";
            //}


            ///Data Section
            // 1
            // ItemCode

            // 2
            writeData(xlWorkSheet, 2, "'00",0);

            // 3
            // Sequence

            //4
            copyData(xlWorkSheet, 1, "IPN", 4, "", 0);
            
            // 5
            writeData(xlWorkSheet, 5, "'00", 0);
            
            // 6
            writeData(xlWorkSheet, 6, "WIP", 0);
            
            // 7
            writeData(xlWorkSheet, 7, "0", 0);
            
            // 8
            // Blank

            // 9
            copyData(xlWorkSheet, 3, "Quantity", 9, "", 0);

            // 10
            writeData(xlWorkSheet, 10, "0", 0);
            
            // 11
            writeData(xlWorkSheet, 11, "100", 0);
            
            // 12
            writeData(xlWorkSheet, 12, "M", 0);
            
            // 13-18
            // Blank

            // 19
            writeData(xlWorkSheet, 19, "N", 0);

            // 20
            copyData(xlWorkSheet, 4, "RefDes", 20, "", 0);

            ///End Data Section


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

            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Items.csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
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


        private void Create_BOM_Scraps()
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



            ///Data Section

            ///End Data Section


            string savePath = "";

            savePath = ImportExcelFile.path;

            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Scraps.csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
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


        private void Create_BOM_Coproducts()
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


            ///End Data Section


            string savePath = "";

            savePath = ImportExcelFile.path;

            string temp = savePath.Replace(ImportExcelFile.filename, "BOM_Coproducts.csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
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


        private void Create_OITM()
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


            copyData(xlWorkSheet, 1, "IPN", 1, "", 1);
            copyData(xlWorkSheet, 5, "Description", 2, "", 1);
            //supplierCatalogNo
            //copyData(xlWorkSheet, 3, "Description", 2, "", 1);
            writeData(xlWorkSheet, 29, "N", 1);
            writeData(xlWorkSheet, 30, "Y", 1);
            writeData(xlWorkSheet, 81, "RCV", 1);
            writeData(xlWorkSheet, 106, "B", 1);
            writeData(xlWorkSheet, 114, "M", 1);
            writeData(xlWorkSheet, 115, "A", 1);
            writeData(xlWorkSheet, 118, "M", 1);


            string savePath = "";

            savePath = ImportExcelFile.path;

            string temp = savePath.Replace(ImportExcelFile.filename, "OITM - Incomar (OMAD & OMAE).csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
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


        private void Create_OITW()
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

            string line1 = "ParentKey;LineNum;MinimalStock;MaximalStock;MinimalOrder;StandardAveragePrice;Locked;InventoryAccount;CostAccount;TransferAccount;RevenuesAccount;VarienceAccount;DecreasingAccount;IncreasingAccount;ReturningAccount;ExpensesAccount;EURevenuesAccount;EUExpensesAccount;ForeignRevenueAcc;ForeignExpensAcc;ExemptIncomeAcc;PriceDifferenceAcc;WarehouseCode;ExpenseClearingAct;PurchaseCreditAcc;EUPurchaseCreditAcc;ForeignPurchaseCreditAcc;SalesCreditAcc;SalesCreditEUAcc;ExemptedCredits;SalesCreditForeignAcc;ExpenseOffsettingAccount;WipAccount;ExchangeRateDifferencesAcct;GoodsClearingAcct;NegativeInventoryAdjustmentAccount;CostInflationOffsetAccount;GLDecreaseAcct;GLIncreaseAcct;PAReturnAcct;PurchaseAcct;PurchaseOffsetAcct;ShippedGoodsAccount;StockInflationOffsetAccount;StockInflationAdjustAccount;VATInRevenueAccount;WipVarianceAccount;CostInflationAccount;WHIncomingCenvatAccount;WHOutgoingCenvatAccount;StockInTransitAccount;WipOffsetProfitAndLossAccount;InventoryOffsetProfitAndLossAccount;DefaultBin;DefaultBinEnforced;PurchaseBalanceAccount;U_RecAdjSuppAcct;U_TechAcctMC;U_SettCstsBT;U_AcctCstsBT;U_CostsAcct;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5;U_ProjectCode";
            string line2 = "ItemCode;LineNum;MinStock;MaxStock;MinOrder;AvgPrice;Locked;BalInvntAc;SaleCostAc;TransferAc;RevenuesAc;VarianceAc;DecreasAc;IncreasAc;ReturnAc;ExpensesAc;EURevenuAc;EUExpensAc;FrRevenuAc;FrExpensAc;ExmptIncom;PriceDifAc;WhsCode;ExpClrAct;APCMAct;APCMEUAct;APCMFrnAct;ARCMAct;ARCMEUAct;ARCMExpAct;ARCMFrnAct;ExpOfstAct;WipAcct;ExchangeAc;BalanceAcc;NegStckAct;CstOffsAct;DecresGlAc;IncresGlAc;PAReturnAc;PurchaseAc;PurchOfsAc;ShpdGdsAct;StkOffsAct;StokRvlAct;VatRevAct;WipVarAcct;CostRvlAct;WhICenAct;WhOCenAct;StkInTnAct;WipOffset;StockOffst;DftBinAbs;DftBinEnfd;PurBalAct;ItemCode;U_RecAdjSuppAcct;U_TechAcctMC;U_SettCstsBT;U_AcctCstsBT;U_CostsAcct;U_OcrCode;U_OcrCode2;U_OcrCode3;U_OcrCode4;U_OcrCode5";

            string[] heading1 = line1.Split(';');
            string[] heading2 = line2.Split(';');

            for (int i = 1; i < heading1.Length + 1; i++)
            {
                xlWorkSheet.Cells[1, i] = heading1[i - 1];

            }

            for (int i = 1; i < heading2.Length + 1; i++)
            {
                xlWorkSheet.Cells[2, i] = heading2[i - 1];
            }



            int startPoint = 0;
            for (int i = 0; i < ImportExcelFile.rows; i++)
            {
                if (ImportExcelFile.ReadCell(i, 2) == "IPN")
                {
                    startPoint = i;
                    break;
                }
            }

            int count = 0;
            for (int i = 0; i < 4; i++)
            {
                string whs = "";
                if (i == 0)
                {
                    whs = "MAIN";
                }
                else if (i == 1)
                {
                    whs = "WIP";
                }
                else if (i == 2)
                {
                    whs = "RCV";
                }
                else
                {
                    whs = "HAL";
                }
                Console.WriteLine("i: " + i.ToString());

                for (int j = 0; j < ImportExcelFile.rows; j++)
                {
                    Console.WriteLine("j: " + j.ToString());
                    //Console.WriteLine(ImportExcelFile.ReadCell(i, 2) + "|");


                    //xlWorkSheet.Cells[i + 1 + 2, 0] = ImportExcelFile.ReadCell(i + startPoint + 1, 2).Replace(';', ':');

                    if (ImportExcelFile.ReadCell(j + startPoint + 1, 1).Replace(';', ':') != "")
                    {
                        xlWorkSheet.Cells[count + 1+2, 1] = ImportExcelFile.ReadCell(j + startPoint + 1, 1).Replace(';', ':');
                        xlWorkSheet.Cells[count + 1+2, 23] = whs;
                        count++;

                    }
                }
            }


            string savePath = "";

            savePath = ImportExcelFile.path;

            string temp = savePath.Replace(ImportExcelFile.filename, "OITW - Incomar (OMAD & OMAE).csv");

            Console.WriteLine(temp);
            print(temp, 10, "");
            print("line", 0, "");

            try
            {
                xlWorkBook.SaveAs(temp, Excel.XlFileFormat.xlCSVWindows, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, true, System.Reflection.Missing.Value, Excel.XlTextVisualLayoutType.xlTextVisualRTL, true);
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
    }
}
