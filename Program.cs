using ClosedXML.Excel;

namespace shopifyNonSeasonalFormatter
{
    internal class Program
    {
        static string sourceFilePath = @"C:\Users\User\Desktop\test.xlsx";
        static IXLWorksheet? sourceWorksheet;
        static int lastRow;
        static int lastColumn;
        static Column[] columnArrayFromSourceSheet = new Column[11];
        static bool needStoneEdgeSpreadsheet = true;
        static bool needShopifySpreadsheet = true;
        static void Main(string[] args)
        {
            importDataFromSourceFile();

            if(needStoneEdgeSpreadsheet)
            {
                XLWorkbook wb = stoneEdgeWorkbook.CreateOrGetTheOneStoneEdgeWorkbookInstance();
                wb.ThisWorksheet.Cell
            }
            if(needShopifySpreadsheet)
            {
                createShopifySpreadsheet();
            }

            foreach (Column column in columnArrayFromSourceSheet)
            {
                if (column != null)
                {
                    Console.WriteLine();
                    Console.WriteLine();
                    Console.WriteLine(column.columnName.ToUpper());
                    Console.WriteLine();
                    foreach (IXLCell cell in column.rows.Cells())
                    {
                        Console.WriteLine(cell.Value);
                    }
                }
            }
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
            Console.WriteLine(smallMessage);
            Console.WriteLine(bigMessage);
            Console.WriteLine();
            Console.ReadLine();
        }

        static void importDataFromSourceFile()

        //
        // gets the columns from the sheet and if they're not empty gets the contents of each column and
        // puts the range into the right slot in the column array, by putting it into the slot of that enum number
        //
        {
            using (var workbook = new XLWorkbook(sourceFilePath))
            {
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
                                if (columnArrayFromSourceSheet[(int)columnHeader.sku] == null)
                                {
                                    //puts it in
                                    columnArrayFromSourceSheet[(int)columnHeader.sku] = createNewColumnObject(columnName, columnNumber);
                                }
                                else
                                {
                                    showAlert("Column Exists", $"there is already a column with name: {columnName}");
                                }
                                break;

                            case "item name" or "name":
                                if (columnArrayFromSourceSheet[(int)columnHeader.itemName] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.itemName] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "size":
                                if (columnArrayFromSourceSheet[(int)columnHeader.size] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.size] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "barcode":
                                if (columnArrayFromSourceSheet[(int)columnHeader.barcode] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.barcode] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "price":
                                if (columnArrayFromSourceSheet[(int)columnHeader.price] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.price] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "supplier" or "supplier name" or "supplierName":
                                if (columnArrayFromSourceSheet[(int)columnHeader.supplierName] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.supplierName] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "gender":
                                if (columnArrayFromSourceSheet[(int)columnHeader.gender] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.gender] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "color_metafield" or "color metafield":
                                if (columnArrayFromSourceSheet[(int)columnHeader.color_metafield] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.color_metafield] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "color_variant":
                                if (columnArrayFromSourceSheet[(int)columnHeader.color_variant] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.color_variant] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "extraTags" or "extra tags":
                                if (columnArrayFromSourceSheet[(int)columnHeader.extraTags] == null)
                                { columnArrayFromSourceSheet[(int)columnHeader.extraTags] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;
                            default:
                                showAlert("Column Name Not Recognized", "no option for column: " + columnName);
                                break;
                        }
                    }
                    else
                    {
                        showAlert($"column {columnNumber} is empty", "");
                    }
                }
            }
        }
        static XLWorkbook createStoneEdgeSpreadsheet()
        {
            
            
            return stoneEdgeWorkbook;
        }
        static XLWorkbook createShopifySpreadsheet()
        {
            XLWorkbook wb = new XLWorkbook();

            return wb;
        }
        static void addStoneEdgeColumnHeaders(XLWorkbook addStoneEdgeHeadersToThisWorkBook)
        {

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
    public enum columnHeader
    {
        sku,
        itemName,
        size,
        barcode,
        price,
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
