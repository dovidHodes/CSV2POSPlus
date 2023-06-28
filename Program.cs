using System;
using ClosedXML.Excel;


//added:  
namespace shopifyNonSeasonalFormatter
{
    internal class Program
    {
        //static string sourceFilePath = @"C:\Users\User\Desktop\test.xlsx";
        static string sourceFilePath = @"/Users/work/Desktop/test.xlsx";
        static IXLWorksheet? sourceWorksheet;
        static int lastRow;
        static int lastColumn;
        static Column[] columnArrayFromSourceSheet = new Column[15];
        static bool needStoneEdgeSpreadsheet = true;
        static bool needShopifySpreadsheet = true;

        static void Main(string[] args)
        {
            ImportDataFromSourceFile();

            if (needStoneEdgeSpreadsheet)
            {
                XLWorkbook stoneEdgeWorkbook = new XLWorkbook();
                IXLWorksheet stoneEdgeWorksheet = stoneEdgeWorkbook.AddWorksheet();

                FillInStoneEdgeColumnHeaders(stoneEdgeWorksheet);

                PasteRangeToLocation("SKU", ColumnHeadersEnum.sku, stoneEdgeWorksheet, 2, 1);

                AddStoneEdgeItem_Name(stoneEdgeWorksheet);

                PasteRangeToLocation("Supplier SKU", ColumnHeadersEnum.supplier_SKU, stoneEdgeWorksheet, 2, 3);
                PasteRangeToLocation("Barcode",      ColumnHeadersEnum.barcode, stoneEdgeWorksheet, 2, 4);                
                PasteRangeToLocation("Cost",         ColumnHeadersEnum.cost, stoneEdgeWorksheet, 2, 5);
                PasteRangeToLocation("Price",        ColumnHeadersEnum.price, stoneEdgeWorksheet, 2, 6);
                PasteRangeToLocation("Taxable", ColumnHeadersEnum.taxable, stoneEdgeWorksheet, 2, 7);
                PasteRangeToLocation("Price", ColumnHeadersEnum.QOH, stoneEdgeWorksheet, 2, 8);

                PrintSpreadsheet(stoneEdgeWorksheet);

                Console.ReadLine();

            }
            if (needShopifySpreadsheet)
            {

            }
        }

        static bool ColumnHasData(ColumnHeadersEnum columnName)
        {
            return columnArrayFromSourceSheet[(int)columnName] != null;
        }
        static void feedUILabel(string message)
        {

        }
        static Column createNewColumnObject(string columnName, int columnNUmber)
        {
            //
            // Makes the range with the (row, column, row, column) overload
            //
            IXLRange rows = sourceWorksheet.Range(2, columnNUmber, lastRow, columnNUmber);
            Column newColumn = new Column(columnName, rows);

            return newColumn;
        }
        static void showAlert(string bigMessage, string smallMessage)
        {
            Console.WriteLine();
            Console.WriteLine(bigMessage.ToUpper());
            Console.WriteLine(smallMessage);       
            Console.WriteLine();
            Console.ReadLine();
        }
        static void ImportDataFromSourceFile()
        {
            //
            // gets the columns from the sheet and if they're not empty gets the contents of each column and
            // puts the range into the right slot in the column array, by putting it into the slot of that enum number
            //
            var workbook = new XLWorkbook(sourceFilePath);
            sourceWorksheet = workbook.Worksheet(1);
            lastColumn = sourceWorksheet.LastColumnUsed().ColumnNumber();
            lastRow = sourceWorksheet.LastRowUsed().RowNumber();

            for (int columnNumber = 1; columnNumber <= lastColumn; columnNumber++)
            {
                if (!sourceWorksheet.Cell(1, columnNumber).IsEmpty())
                {
                    string columnName = (string)sourceWorksheet.Cell(1, columnNumber).Value;
                    switch (columnName.ToLower())
                    {
                        case "sku":
                            //first makes sure there is no other column with that header name that was already put into a slot
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.sku] == null)
                            {
                                //then for sku column, checks to make sure that all cells are not empty
                                for (int row = 2; row <= lastRow; row++)
                                {
                                    if (sourceWorksheet.Cell(row, columnNumber).IsEmpty())
                                    {
                                        showAlert("missing value from required column", $"Row {row} for SKU column is empty");
                                    }
                                }
                                //puts it in
                                columnArrayFromSourceSheet[(int)ColumnHeadersEnum.sku] = createNewColumnObject(columnName, columnNumber);
                            }
                            else
                            {
                                showAlert("Column Exists", $"there is already a column with name: {columnName}");
                            }
                            break;

                        case "item name" or "name":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.itemName] == null)
                            {
                                for (int row = 2; row <= lastRow; row++)
                                {
                                    if (sourceWorksheet.Cell(row, columnNumber).IsEmpty())
                                    {
                                        showAlert("missing value from required column", $"Row {row} for Item Name column is empty");
                                    }
                                }
                                columnArrayFromSourceSheet[(int)ColumnHeadersEnum.itemName] = createNewColumnObject(columnName, columnNumber);
                            }
                            else
                            {
                                showAlert("Column Exists", $"there is already a column with name: {columnName}");
                            }
                            break;

                        case "supplier sku":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.supplier_SKU] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.supplier_SKU] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "size":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "barcode":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.barcode] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.barcode] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "price":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.price] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.price] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "taxable":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.taxable] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.taxable] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "cost":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.cost] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.cost] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "qoh" or "quantity":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.QOH] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.QOH] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "supplier" or "supplier name" or "supplierName":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.supplierName] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.supplierName] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "gender":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.gender] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.gender] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "color_metafield" or "color metafield":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_metafield] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_metafield] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "color_variant":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;

                        case "extraTags" or "extra tags":
                            if (columnArrayFromSourceSheet[(int)ColumnHeadersEnum.extraTags] == null)
                            { columnArrayFromSourceSheet[(int)ColumnHeadersEnum.extraTags] = createNewColumnObject(columnName, columnNumber); }
                            else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                            break;
                        default:
                            showAlert("Column Name Not Recognized", "no option for column: " + columnName);
                            break;
                    }
                }
                else
                {
                    showAlert($"column {columnNumber} column header is empty", "");
                }
            }
        }
        static void FillInStoneEdgeColumnHeaders(IXLWorksheet stoneEdgeWorksheet)
        {
            string[] stoneEdgeColumnHeaderNames = new string[] { "SKU", "Item Name", "Supplier Sku", "Barcode", "Cost", "price", "taxable", "QOH" };
            for (int row = 1, column = 1; column <= stoneEdgeColumnHeaderNames.Length; column++)
            {
                feedUILabel($"Filling in Stone Edge Header: {stoneEdgeColumnHeaderNames[column - 1]}");
                stoneEdgeWorksheet.Cell(row, column).Value = stoneEdgeColumnHeaderNames[column - 1];
            }
        }
        static void PasteRangeToLocation(string rangeName, ColumnHeadersEnum columnEnum, IXLWorksheet destinationWorksheet, int destinationRow, int destinationColumn)
        {
            if (ColumnHasData(columnEnum))
            {
                IXLRange data = columnArrayFromSourceSheet[(int)columnEnum].rows;
                feedUILabel($"Pasting range to {rangeName}");
                data.CopyTo(destinationWorksheet.Cell(destinationRow, destinationColumn));
            }
        }
        static void AddStoneEdgeItem_Name(IXLWorksheet stoneEdgeWorkSheet)
        {
            bool hasColorVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant] != null;
            bool hasSizeVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size] != null;

            if (hasColorVariants || hasSizeVariants)
            {
                //gets the column info for the title column and variant columns,
                //the null conditional operator only assigns the value if the column object isnt null to avoid null reference exeptions
                IXLRangeColumn titleColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.itemName].rows.Column(1);
                IXLRangeColumn colorColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant]?.rows.Column(1);
                IXLRangeColumn sizeColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size]?.rows.Column(1);

                for (int row = 1; row <= lastRow; row++)
                {
                    //sets the value for the item mame column
                    stoneEdgeWorkSheet.Cell(row + 1, 2).Value = titleColumn.Cell(row).Value;
                    if (hasColorVariants)
                    {
                        //appends color variant to end of name
                        stoneEdgeWorkSheet.Cell(row + 1, 2).Value = stoneEdgeWorkSheet.Cell(row + 1, 2).Value + " " + colorColumn.Cell(row).Value;
                    }
                    if (hasSizeVariants)
                    {
                        //appends size variant to end of name
                        stoneEdgeWorkSheet.Cell(row + 1, 2).Value = stoneEdgeWorkSheet.Cell(row + 1, 2).Value + " " + sizeColumn.Cell(row).Value;
                    }
                }
            }
            else
            {
                PasteRangeToLocation("Title", ColumnHeadersEnum.itemName, stoneEdgeWorkSheet, 2, 2);
            }

        }
        static void PrintSpreadsheet(IXLWorksheet worksheetToPrint)
        {
            Console.WriteLine("------------------------------------------------------------------------------------------------------------------------------------------------");
            for (int row = 1; row <= worksheetToPrint.LastRowUsed().RowNumber(); row++)
            {
                for (int column = 1; column <= worksheetToPrint.LastColumnUsed().ColumnNumber(); column++)
                {
                    string cellValue = $"{worksheetToPrint.Cell(row, column).Value}";
                    if (row == 1)
                    {
                        cellValue = cellValue.ToUpper();
                    }
                    Console.Write($"{cellValue,-20}");
                }
                Console.WriteLine();
                if(row == 1)
                {
                    Console.Write("------------------------------------------------------------------------------------------------------------------------------------------------");
                }
                Console.WriteLine();
            }
            Console.Write("------------------------------------------------------------------------------------------------------------------------------------------------");
        }
    }
    public class Column
    {
        public string columnName;
        public IXLRange rows;

        public Column(string columnName, IXLRange rows)
        {
            this.rows = rows;
            this.columnName = columnName;
        }
    }
    public enum ColumnHeadersEnum
    {
        sku,
        itemName,
        supplier_SKU,
        size,
        barcode,
        price,
        cost,
        QOH,
        taxable,
        supplierName,
        productType,
        gender,
        color_metafield,
        color_variant,
        extraTags
    }
    //NEXT:: maybe put the part of the switch statement taht checks if the slot was already used in the column array, move it to the colmancreator method
    //
    //NEXT: put in a validator that makes sure the neccesary columns exist (name, sku) and a warning for columns that are missing
    //
    //
    //
    //  
    //  user will check a box if there are variants of color or size, if yes, then set the variants bool to yes
    // ,and prompts if it's color varuants or size variants and then proceed to use what's in
    // the color column, (if they choos yes, then that column must not be empty) same for size
    //
    //
    // instead of having the system delete extra columns and have all the formulas, just concantonate any values together from any
    // "taggable" column object  (just have the system know which ones are taggable) duh
    //
    //
    // maybe better to have the program do all the chesboning of sizes and new products insteasd of the the formulas
    // so if there is a size column it will read each cell in title and do, if same as previous then do size column info etc, and if no size column, 
    // then fill in with defaylt title.
    //
    // actually can do mixed with variants and without!! if there is a size column, then if the value of the size column for that row is null
    // then that means that it's a non variant row and fill in default title. also add warning: "some rows in the size column are empty,
    // are you sure this a mixed type, or are you missing values?
    //
    //
    // If price column is empty, set it as zero, matrixifty doesnt allow price of 0
    //
}
