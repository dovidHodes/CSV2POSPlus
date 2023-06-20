using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace shopifyNonSeasonalFormatter
{

    //actually. can use the color metafeild for when end user does select product has variants, just cjeck if the size coulumn exists, if not then, we can use color from metafield
    internal class Program
    {
        static string sourceFilePath = @"C:\Users\User\Desktop\test.xlsx";
        static IXLWorksheet? sourceWorksheet;
        static int lastRow;
        static int lastColumn;
        //static int columnCount = 0;
        static void Main(string[] args)
        {
            
            column[] columns = new column[11];

            using (var workbook = new XLWorkbook(sourceFilePath))
            {
                sourceWorksheet = workbook.Worksheet(1);
                lastColumn = sourceWorksheet.LastColumnUsed().ColumnNumber();
                lastRow = sourceWorksheet.LastRowUsed().RowNumber();

                //
                // gets the columns from the sheet and if they're not empty gets the contents of each column and
                // puts the range into the right slot in the column array, by putting it into the slot of that enum number
                //
                for (int columnNumber = 1; columnNumber <= lastColumn; columnNumber++)
                {
                    if(!sourceWorksheet.Cell(1, columnNumber).IsEmpty())
                    {
                        string columnName = (string)sourceWorksheet.Cell(1, columnNumber).Value;
                        switch (columnName.ToLower())
                        {
                            case "sku":
                                //first makes sure there is no other column with that header name that was already put into a slot
                                if(columns[(int)columnHeader.sku] == null)
                                {
                                    //puts it in
                                    columns[(int)columnHeader.sku] = createNewColumnObject(columnName, columnNumber);
                                }
                                else
                                {
                                    showAlert("Column Exists", $"there is already a column with name: {columnName}");
                                }
                                break;

                            case "item name" or "name":
                                if (columns[(int)columnHeader.itemName] == null)
                                { columns[(int)columnHeader.itemName] = createNewColumnObject(columnName, columnNumber); }                          
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }             
                                break;

                            case "size":
                                if (columns[(int)columnHeader.size] == null) 
                                { columns[(int)columnHeader.size] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "barcode":
                                if (columns[(int)columnHeader.barcode] == null)
                                { columns[(int)columnHeader.barcode] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "price":
                                if (columns[(int)columnHeader.price] == null)
                                { columns[(int)columnHeader.price] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "supplier" or "supplier name" or "supplierName":
                                if (columns[(int)columnHeader.supplierName] == null)
                                { columns[(int)columnHeader.supplierName] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "gender":
                                if (columns[(int)columnHeader.gender] == null)
                                { columns[(int)columnHeader.gender] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "color_metafield" or "color metafield":
                                if (columns[(int)columnHeader.color_metafield] == null)
                                { columns[(int)columnHeader.color_metafield] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "color_variant":
                                if (columns[(int)columnHeader.color_variant] == null)
                                { columns[(int)columnHeader.color_variant] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;

                            case "extraTags" or "extra tags":
                                if (columns[(int)columnHeader.extraTags] == null)
                                { columns[(int)columnHeader.extraTags] = createNewColumnObject(columnName, columnNumber); }
                                else { showAlert("Column Exists", $"there is already a column with name: {columnName}"); }
                                break;
                        }
                    }
                    else
                    {
                        showAlert($"column {columnNumber} is empty", "");
                    }
                }

                foreach (column column in columns)
                {
                    if(column!= null)
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
            //  
            static column createNewColumnObject(string columnName, int columnNUmber)
            {
                //
                // Makes the range with the (row, column, row, column) overload
                //
                IXLRange rows = sourceWorksheet.Range(2, columnNUmber, lastRow, columnNUmber);
                column newColumn = new column(columnName, rows);

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
        }
        public class column
        {
            public string columnName;
            public IXLRange rows;

            public column(string columnName, IXLRange rows)
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
        //
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
}