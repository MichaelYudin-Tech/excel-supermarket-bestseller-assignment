using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;           
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel; 


namespace Menahel4u
{
    public class GetBestsellers
    {
        public static List<string> totalShoppingList = new List<string>();
        static void Main()
        {

            // List excel shopping files
            var dir = new DirectoryInfo(@"D:\Michael\Projects\menahel4u\excel-supermarket-bestseller-assignment");
            var files = dir.GetFiles("*.xlsx");

            // Create COM object
            Excel.Application xlApp = new Excel.Application();

            // Iterate over excel files
            foreach (var file in files)
            {
                totalShoppingList.AddRange(GetShoppingList(xlApp, file));
            }

            // Analyze bestsellers from totalShoppingList
            List<string> bestSellers = GetBestsellers(totalShoppingList);

            // Output bestseller results to new excel file
            OutputToExcel(xlApp, bestSellers);

            // Cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    
        public static List<string> GetShoppingList(Excel.Application xlApp, FileInfo file)
        {
            List<string> shoppingList = new List<string>();

            Excel.Workbook xlWorkbook = null;
            Excel._Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;

            try
            {
                xlWorkbook = xlApp.Workbooks.Open(file.FullName);
                xlWorksheet = xlWorkbook.Sheets[1];
                xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                // return if no actual data beside the column names
                if (rowCount < 2) return shoppingList;

                int colCount = xlRange.Columns.Count;
                // return if excel file structure is not fitting
                if (colCount < 2) return shoppingList;

                int productColumn = 1;
                int quantityColumn = 2;
                 

                // r = row, c = column
                for (int r = 2; r <= rowCount; r++)
                {
                    // if the value of column "Quantity" an int, add the product to the shopping list x times it was purchased
                    if (xlRange.Cells[r, quantityColumn] == null || xlRange.Cells[r, quantityColumn].Value2 == null)
                    {
                        continue;
                    }

                    object quantityCellValue = xlRange.Cells[r, quantityColumn].Value2;

                    // Check that product cell exists and is not null 
                    if (xlRange.Cells[r, productColumn] == null || xlRange.Cells[r, productColumn].Value2 == null) 
                    {
                        continue;
                    }

                    if (quantityCellValue is int || (quantityCellValue is double && (double)quantityCellValue % 1 == 0))
                    {
                        int quantity = Convert.ToInt32(quantityCellValue);

                        for (int i = 0; i < quantity; i++) 
                        {
                            shoppingList.Add(xlRange.Cells[r, 1].Value2.ToString());
                        }
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                if (xlRange != null) Marshal.ReleaseComObject(xlRange);
                if (xlWorksheet != null) Marshal.ReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);
                }
            }

            return shoppingList;
        }

        public static List<string> GetBestsellers(List<string> totalShoppingListToCount)
        {
            return (from i in totalShoppingListToCount
                group i by i into g
                orderby g.Count() descending
                select g.Key).Take(10).ToList();
        }

        public static void OutputToExcel(Excel.Application xlApp, List<string> bestSellers)
        {
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            xlWorksheet.Name = "Best Sellers";

            xlWorksheet.Cells[1, 1] = "Rank";
            xlWorksheet.Cells[1, 2] = "Product";

            for (int i = 0; i < bestSellers.Count; i++)
            {
                xlWorksheet.Cells[i + 2, 1] = i + 1;
                xlWorksheet.Cells[i + 2, 2] = bestSellers[i];
            }

            string filePath = @"D:\Michael\Projects\menahel4u\BestSellers.xlsx";
            xlWorkbook.SaveAs(filePath);

            // Cleanup
            xlWorkbook.Close();
            xlApp.Quit();

            // Release COM objects
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
        }
    }
}   