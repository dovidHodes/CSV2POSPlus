using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace shopifyNonSeasonalFormatter
{
   public class stoneEdgeWorkbook
    {
        private static stoneEdgeWorkbook theOneAndOnlyInstanceOfStoneEdgeWorkbook;

        public XLWorkbook Workbook;
        public IXLWorksheet ThisWorksheet;

        //constructor
        private stoneEdgeWorkbook()
        {
            Workbook = new XLWorkbook();
            ThisWorksheet = Workbook.Worksheets.Add();
            FillInHeaderNames();
        }

        // singleton pattern and lazy instatiation public method
        public static XLWorkbook CreateOrGetTheOneStoneEdgeWorkbookInstance()
        {
            if (theOneAndOnlyInstanceOfStoneEdgeWorkbook == null)
            {
                theOneAndOnlyInstanceOfStoneEdgeWorkbook = new stoneEdgeWorkbook();
            }
            return theOneAndOnlyInstanceOfStoneEdgeWorkbook.Workbook;
        }
        private void FillInHeaderNames()
        {
            string[] stoneEdgeColumnHeaderNames = new string[] { "SKU", "Item Name", "Supplier Sku", "Barcode", "Cost", "price", "taxable", "QOH" };
            int row = 1;
            int column = 1;
            foreach (string columnHeaderName in stoneEdgeColumnHeaderNames)
            {
                ThisWorksheet.Cell(row, column).Value = columnHeaderName;
            }
        }
    }
}
    