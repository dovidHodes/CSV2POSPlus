﻿using System;
using ClosedXML.Excel;

//added:
namespace shopifyNonSeasonalFormatter
{
    internal class Program
    {
        static IXLWorksheet? sourceWorksheet;
        static int lastRow;
        static int lastColumn;
        static Column[] columnArrayFromSourceSheet = new Column[15];
        static string saveFilepath = @"/Users/work/Desktop/";
        static bool setAsDraftOnShopify = false;

        static bool needStoneEdgeSpreadsheet = false;
        static bool needShopifySpreadsheet = true;
        //static string sourceFilePath = @"C:\Users\User\Desktop\test.xlsx";
        static string sourceFilePath = @"/Users/work/Desktop/test.xlsx";

        static void Main(string[] args)
        {
            ImportDataFromSourceFile(out bool requiredDataPresent);

            if (requiredDataPresent)
            {
                if (needStoneEdgeSpreadsheet)
                {
                    XLWorkbook stoneEdgeWorkbook = new XLWorkbook();
                    IXLWorksheet stoneEdgeWorksheet = stoneEdgeWorkbook.AddWorksheet();
                    FillInStoneEdgeColumnHeaders(stoneEdgeWorksheet);
                    PasteRangeToLocation("SKU", ColumnHeadersEnum.sku, stoneEdgeWorksheet, 2, 1);
                    AddStoneEdgeItem_Name(stoneEdgeWorksheet);
                    PasteRangeToLocation("Supplier SKU", ColumnHeadersEnum.supplier_SKU, stoneEdgeWorksheet, 2, 3);
                    PasteRangeToLocation("Barcode", ColumnHeadersEnum.barcode, stoneEdgeWorksheet, 2, 4);
                    PasteRangeToLocation("Cost", ColumnHeadersEnum.cost, stoneEdgeWorksheet, 2, 5);
                    PasteRangeToLocation("Price", ColumnHeadersEnum.price, stoneEdgeWorksheet, 2, 6);
                    PasteRangeToLocation("Taxable", ColumnHeadersEnum.taxable, stoneEdgeWorksheet, 2, 7);
                    PasteRangeToLocation("QOH", ColumnHeadersEnum.QOH, stoneEdgeWorksheet, 2, 8);

                    PrintSpreadsheet(stoneEdgeWorksheet);
                    SaveFileAs(stoneEdgeWorkbook, "Stone Edge Sheet", true);
                }
                if (needShopifySpreadsheet)
                {
                    XLWorkbook shopifyWorkbook = new XLWorkbook();
                    IXLWorksheet shopifyWorksheet = shopifyWorkbook.AddWorksheet();
                    FillInShopifyColumnHeaders(shopifyWorksheet);
                    bool hasColorVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant] != null;
                    bool hasSizeVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size] != null;
                    bool hasVariants = hasSizeVariants || hasColorVariants;
                    bool[] rowData = new bool[lastRow-1];
                    if (hasVariants) {
                        rowData = getRowVariantData(shopifyWorksheet);
                    }
                    PasteRangeToLocation("title shopify", ColumnHeadersEnum.itemName, shopifyWorksheet, 2, 4);
                    for(int row = 2; row <= lastRow; row++) {
                        shopifyWorksheet.Cell(row, 3).Value = rowData[row-2];
                    }
                    SaveFileAs(shopifyWorkbook, "SHOPIFY", false);
                    if (hasSizeVariants)
                    {

                    }



                }
            }
            else
            {
                showAlert("required data was not present", "no changes were made");
            }
        }

        static bool ColumnHasData(ColumnHeadersEnum columnName)
        {
            return columnArrayFromSourceSheet[(int)columnName] != null;
        }
        static void feedUILabel(string message)
        {

        }
        static bool createNewColumnObject(ColumnHeadersEnum columnEnum, bool isRequired, string columnName, int columnNUmber)
        {
            //checks if the slot is already used
            if (columnArrayFromSourceSheet[(int)columnEnum] == null)
            {
                if (isRequired)
                {
                    //then for required columns, checks to make sure that all cells are not empty
                    for (int row = 2; row <= lastRow; row++)
                    {
                        if (sourceWorksheet.Cell(row, columnNUmber).IsEmpty())
                        {
                            showAlert("missing value from required column", $"Row {row} for {columnName} column is empty");
                            return false;
                        }
                    }
                }
                // Makes the range with the (row, column, row, column) overload, and adds the column to the array
                IXLRange rows = sourceWorksheet.Range(2, columnNUmber, lastRow, columnNUmber);
                columnArrayFromSourceSheet[(int)columnEnum] = new Column(columnName, rows);
            }
            else
            {
                showAlert("Column Exists", $"there is already a column with name: {columnName}");
            }
            //if all required data is present
            return true;
        }
        static void showAlert(string bigMessage, string smallMessage)
        {
            Console.WriteLine();
            Console.WriteLine(bigMessage.ToUpper());
            Console.WriteLine(smallMessage);       
            Console.WriteLine();
            Console.ReadLine();
        }
        static void ImportDataFromSourceFile(out bool requiredDataPresent)
        {
            requiredDataPresent = true;
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
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.sku, true, columnName, columnNumber);                        
                            break;

                        case "item name" or "name":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.itemName, true, columnName, columnNumber);
                            break;

                        case "supplier sku":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.supplier_SKU, false, columnName, columnNumber);
                            break;

                        case "size":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.size, false, columnName, columnNumber);
                            break;

                        case "barcode":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.barcode, false, columnName, columnNumber);
                            break;

                        case "price":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.price, false, columnName, columnNumber);
                            break;

                        case "taxable":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.taxable, false, columnName, columnNumber);
                            break;

                        case "cost":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.cost, false, columnName, columnNumber);
                            break;

                        case "qoh" or "quantity":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.QOH, false, columnName, columnNumber);
                            break;

                        case "supplier" or "supplier name" or "supplierName":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.supplierName, true, columnName, columnNumber);
                            break;

                        case "gender":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.gender, false, columnName, columnNumber);
                            break;

                        case "color_metafield" or "color metafield":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.color_metafield, false, columnName, columnNumber);
                            break;

                        case "color_variant":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.color_variant, false, columnName, columnNumber);
                            break;

                        case "extraTags" or "extra tags":
                            requiredDataPresent = createNewColumnObject(ColumnHeadersEnum.extraTags, false, columnName, columnNumber);
                            break;
                        default:
                            showAlert("Column Name Not Recognized", "no option for column: " + columnName);
                            break;
                    }
                }
                else
                {
                    showAlert($"column {columnNumber} column header is empty", "");
                    requiredDataPresent = false;
                    break;
                }              
            }
            if (requiredDataPresent)
            {
                for (int i = 0; i < 3; i++)
                {
                    if (columnArrayFromSourceSheet[i] == null)
                    {
                        string missingField = i == 0 ? "SKU" : i == 1 ? "Item Name" : "Supplier Name";
                        showAlert("missing required column from spreadsheet", $"missing column {missingField.ToUpper()}, which is a neccessary field");
                        requiredDataPresent = false;
                        break;
                    }
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
            //gets the values from the title and supplier name columns to concat. and set as value for title column
            IXLRangeColumn supplierNameColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.supplierName].rows.Column(1);
            IXLRangeColumn titleColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.itemName].rows.Column(1);
            for (int row = 1; row <= lastRow; row++)
            {
                stoneEdgeWorkSheet.Cell(row + 1, 2).Value = supplierNameColumn.Cell(row).Value + " " + titleColumn.Cell(row).Value;
            }
            //checks if there are color or size variants
            bool hasColorVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant] != null;
            bool hasSizeVariants = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size] != null;
            if (hasColorVariants || hasSizeVariants)
            {
                //gets the column info for the variant columns,
                //the null conditional operator only assigns the value if the column object isnt null to avoid null reference exeptions
                IXLRangeColumn colorColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.color_variant]?.rows.Column(1);
                IXLRangeColumn sizeColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.size]?.rows.Column(1);
                for (int row = 1; row <= lastRow; row++)
                {
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
        static void SaveFileAs(XLWorkbook workbookToSave, string fileName, bool isCSV)
        {
            string filePath = Path.Combine(saveFilepath, fileName + ".xlsx");
            workbookToSave.SaveAs(filePath);
            if (!isCSV)
            {
                showAlert("file saved", "Shopify file successfully saved");
            }
            else
            {
                File.Move(filePath, Path.ChangeExtension(filePath, ".csv"));
                showAlert("file saved", "SE file successfully saved as CSV");
            }
        }
        static void FillInShopifyColumnHeaders(IXLWorksheet shopifyWorksheet)
        {
            string[] ShopifyEdgeColumnHeaderNames = new string[] {"Handle", "Variant SKU", "Is New" ,"Title", "Option 1 Name", "Option 1 Value", "Variant Barcode", "Variant Cost", "Variant Price", "Vendor", "Type", "Metafield: custom.gender [single_line_text_field]", "Metafield: custom.color [single_line_text_field]", "Tags", "Body HTML", "Image Src", "Image Command", "Image Position", "Image Alt Text", "Tags Command", "Status", "Published", "Published Scope", "Gift Card", "Variant Weight", "Variant Weight Unit", "Variant Requires Shipping", "Variant Taxable", "Variant Inventory Tracker", "Variant Inventory Policy", "Variant Fulfillment Service" };
            for (int row = 1, column = 1; column <= ShopifyEdgeColumnHeaderNames.Length; column++)
            {
                feedUILabel($"Filling in Shopify Header: {ShopifyEdgeColumnHeaderNames[column - 1]}");
                shopifyWorksheet.Cell(row, column).Value = ShopifyEdgeColumnHeaderNames[column - 1];
            }
        }
        static bool[] getRowVariantData(IXLWorksheet shopifyWorksheet)
        {
            int nonHeaderRows = lastRow - 1;
            int firstNonHeaderRow = 2;
            bool[] rowData = new bool[nonHeaderRows];
            rowData[0] = true;
            IXLRangeColumn titleColumn = columnArrayFromSourceSheet[(int)ColumnHeadersEnum.itemName]?.rows.Column(1);
            for(int row = 2; row <= nonHeaderRows; row++)
            {
                rowData[row - 1] = !titleColumn.Cell(row).Value.Equals(titleColumn.Cell(row-1).Value);
            }
            return rowData;
        }
        static void FillShopifyTitleColumn(IXLWorksheet shopifyWorksheet, bool hasVariants, bool hasSizeVariants, bool hasColorVariants)
        {
            if (hasSizeVariants) {
                
            }
                
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
        supplierName,
        supplier_SKU,
        size,
        barcode,
        price,
        cost,
        QOH,
        taxable,
        productType,
        gender,
        color_metafield,
        color_variant,
        extraTags
    }
    // rearrange the order of strings im shopify column header array, to order by importance
    //
    // add a bool to set the shopify products as draft
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
