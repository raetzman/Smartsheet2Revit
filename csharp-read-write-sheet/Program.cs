using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

// Add nuget reference to smartsheet-csharp-sdk (https://www.nuget.org/packages/smartsheet-csharp-sdk/)
using Smartsheet.Api;
using Smartsheet.Api.Models;
using Smartsheet.Api.OAuth;

namespace csharp_read_write_sheet
{
    class Program
    {
        // The API identifies columns by Id, but it's more convenient to refer to column names
        static Dictionary<string, long> columnMap = new Dictionary<string, long>(); // Map from friendly column name to column Id 

        static void Main(string[] args)
        {
            // Get API access token from App.config file or environment
            string accessToken = ConfigurationManager.AppSettings["AccessToken"];
            if (string.IsNullOrEmpty(accessToken))
                accessToken = "Bearer m4u13nfkn1ck7abdap6odkaq9j"; // Environment.GetEnvironmentVariable("SMARTSHEET_ACCESS_TOKEN");
            if (string.IsNullOrEmpty(accessToken))
                throw new Exception("Must set API access token in App.conf file");

            // Initialize client
            SmartsheetClient ss = new SmartsheetBuilder().SetAccessToken(accessToken).Build();

            //Sheet sheet = ss.SheetResources.ImportXlsSheet("../../../Sample Sheet.xlsx", null, 0, null);
            Sheet sheet = ss.SheetResources.GetSheet(4650939152918404, null, null, null, null, null, null, null);
            // Load the entire sheet
            sheet = ss.SheetResources.GetSheet(sheet.Id.Value, null, null, null, null, null, null, null);
            Console.WriteLine("Loaded " + sheet.Rows.Count + " rows from sheet: " + sheet.Name);

            // Build column map for later reference
            foreach (Column column in sheet.Columns)
                columnMap.Add(column.Title, (long)column.Id);

            // Accumulate rows needing update here
            List<Row> rowsToUpdate = new List<Row>();

            foreach (Row row in sheet.Rows)
            {
                Row rowToUpdate = evaluateRowAndBuildUpdates(row);
                if (rowToUpdate != null)
                    rowsToUpdate.Add(rowToUpdate);
            }

            // Finally, write all updated cells back to Smartsheet 
            Console.WriteLine("Writing " + rowsToUpdate.Count + " rows back to sheet id " + sheet.Id);
            ss.SheetResources.RowResources.UpdateRows(sheet.Id.Value, rowsToUpdate);
            Console.WriteLine("Done (Hit enter)");
            Console.ReadLine();
        }


        /*
         * TODO: Replace the body of this loop with your code
         * This *example* looks for rows with a "Status" column marked "Complete" and sets the "Remaining" column to zero
         * 
         * Return a new Row with updated cell values, else null to leave unchanged
         */
        static Row evaluateRowAndBuildUpdates(Row sourceRow)
        {
            Row rowToUpdate = null;

            // Find cell we want to examine
            Cell statusCell = getCellByColumnName(sourceRow, "Status");
            if (statusCell.DisplayValue == "Complete")
            {
                Cell remainingCell = getCellByColumnName(sourceRow, "Remaining");
                if (remainingCell.DisplayValue != "0")                  // Skip if "Remaining" is already zero
                {
                    Console.WriteLine("Need to update row # " + sourceRow.RowNumber);

                    var cellToUpdate = new Cell
                    {
                        ColumnId = columnMap["Remaining"],
                        Value = 0
                    };

                    var cellsToUpdate = new List<Cell>();
                    cellsToUpdate.Add(cellToUpdate);

                    rowToUpdate = new Row
                    {
                        Id = sourceRow.Id,
                        Cells = cellsToUpdate
                    };
                }
            }
            return rowToUpdate;
        }


        // Helper function to find cell in a row
        static Cell getCellByColumnName(Row row, string columnName)
        {
            return row.Cells.First(cell => cell.ColumnId == columnMap[columnName]);
        }
    }
}
//public class SampleFunctions {
//    public static double MultiplyByTwo(double inputNumber) {
//        return inputNumber * 2.0;
//    }

//    public static double MultiplyByThree(double inputNumber) {
//        return inputNumber * 3.0;
//    }

//    public static String HelloWorld(int languageId) {
//        switch (languageId) {
//            case 0:
//                return "Hello World!";
//                break;
//            case 1:
//                return "Hallo Welt!";
//                break;
//            case 2:
//                return "Merhaba Alem!";
//                break;
//            default:
//                return "Hello World!";
//                break;
//        }
//    }

//    private static String accessToken = "m4u13nfkn1ck7abdap6odkaq9j";

//    private static long sheetIdAPITest = 4650939152918404;

//    private static long rowIdXXX = 8247651579783044;

//    private static int rowNumberXXX = 2;

//    private static long columnIdXXX = 3165800489084804;

//    private static int columnNumberXXX = 2;

//    private static SmartsheetClient getSmartsheetClient() {
//        return new SmartsheetBuilder().SetAccessToken(accessToken).Build(); ;
//    }

//    [MultiReturn(new[] { "List of all Sheets" })]
//    public static Dictionary<string, object> getAllSheets() {
//        Dictionary<string, object> dictionaryResult = new Dictionary<string, object>();

//        var listOfAllSheets = new List<String>();

//        SmartsheetClient smartsheetClient = getSmartsheetClient();

//        PaginatedResult<Sheet> paginatedResultSheets = smartsheetClient.SheetResources.ListSheets(null, null, null);

//        for (int i = 0; i < paginatedResultSheets.TotalCount; i++) {
//            listOfAllSheets.Add(paginatedResultSheets.Data[i].Name);
//        }

//        dictionaryResult.Add("List of all Sheets", listOfAllSheets);

//        return dictionaryResult;
//    }

//    public static String getCellValue() {
//        String result = "";

//        SmartsheetClient smartsheetClient = getSmartsheetClient();

//        Sheet sheetAPITest = smartsheetClient.SheetResources.GetSheet(sheetIdAPITest, null, null, null, null, null, null, null);

//        Row row = sheetAPITest.GetRowByRowNumber(rowNumberXXX);

//        Cell cell = row.Cells.First(c => c.ColumnId == columnIdXXX);

//        result = cell.Value.ToString();

//        return result;
//    }

//    public static String getCellValue(int rowIndex, int columnIndex) {
//        String result = "";

//        SmartsheetClient smartsheetClient = getSmartsheetClient();

//        Sheet sheetAPITest = smartsheetClient.SheetResources.GetSheet(sheetIdAPITest, null, null, null, null, null, null, null);

//        Row row = sheetAPITest.GetRowByRowNumber(rowIndex);

//        Column column = sheetAPITest.GetColumnByIndex(columnIndex);

//        Cell cell = row.Cells.First(c => c.ColumnId == column.Id);

//        result = cell.Value.ToString();

//        return result;
//    }

//    public static String getCellDisplayValue() {
//        String result = "";

//        SmartsheetClient smartsheetClient = getSmartsheetClient();

//        Sheet sheetAPITest = smartsheetClient.SheetResources.GetSheet(sheetIdAPITest, null, null, null, null, null, null, null);

//        Row row = sheetAPITest.GetRowByRowNumber(rowNumberXXX);

//        Cell cell = row.Cells.First(c => c.ColumnId == columnIdXXX);

//        result = cell.Value.ToString();

//        return result;
//    }


//    public static void setCellValue(int rowIndex, int columnIndex, String value) {
//        SmartsheetClient smartsheetClient = getSmartsheetClient();

//        Sheet sheetAPITest = smartsheetClient.SheetResources.GetSheet(sheetIdAPITest, null, null, null, null, null, null, null);

//        Row row = sheetAPITest.GetRowByRowNumber(rowIndex);

//        Column column = sheetAPITest.GetColumnByIndex(columnIndex);

//        Cell cell = row.Cells.First(c => c.ColumnId == column.Id);

//        Cell cellToUpdate = new Cell {
//            ColumnId = cell.ColumnId,
//            Value = value
//        };

//        var listOfCellsToUpdate = new List<Cell>();

//        listOfCellsToUpdate.Add(cellToUpdate);

//        Row rowToUpdate = new Row {
//            Id = row.Id,
//            Cells = listOfCellsToUpdate
//        };

//        var listOfRowsToUpdate = new List<Row>();

//        listOfRowsToUpdate.Add(rowToUpdate);

//        smartsheetClient.SheetResources.RowResources.UpdateRows(sheetIdAPITest, listOfRowsToUpdate);
//    }
//}
