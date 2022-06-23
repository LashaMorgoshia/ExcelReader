using IronXL;
using System;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelReader
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // declare excel file path
            var filePath = "file.xlsx";

            if (args.Length > 0)
                filePath = args[0];

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                Console.WriteLine($"File not found, please move your Excel file into specific path: {filePath}");
                Console.WriteLine("Press Enter to exit ..");
                Console.ReadLine();
                return;
            }

            // read excel
            Read(filePath);

            // generate sql script
            GenerateSQLScripts(filePath);

            // display complation message
            Console.WriteLine("done");
            Console.ReadLine();
        }

        static void Read(string filePath)
        {
            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workbook = WorkBook.Load(filePath);
            WorkSheet sheet = workbook.WorkSheets.First();

            //Select cells easily in Excel notation and return the calculated value
            var cellValue = sheet["H2"].StringValue;
            var rowCount = sheet.RowCount;

            // insert new value into another cell
            sheet.SetCellValue(0, 13, "Converted_BD");

            // Read from Ranges of cells elegantly.
            int i = 1;
            foreach (var cell in sheet[$"H2:H{rowCount}"])
            {
                Console.WriteLine($"{i} Cell {cell.AddressString} has value '{cell.Text}'");

                // convert original value
                var convertedValue = TransformOriginalValueIntoCorrected(cell.Text);
               
                // insert new value into another cell
                sheet.SetCellValue(i, 13, convertedValue);

                i++;
            }

            sheet.SaveAs(@"corrected_file.xlsx");
        }

        static string TransformOriginalValueIntoCorrected(string cellValue)
        {
            // declare new variable for conversion
            var newValue = default(DateTime);

            // convert num to month
            cellValue = ConvertNumToMonth(cellValue);

            // replace GEO names and Symbols
            cellValue = ConvertGeo2En(cellValue);

            // try to convert into dt
            DateTime.TryParse(cellValue, out newValue);

            // try to convert into another format
            if (newValue == DateTime.MinValue && !string.IsNullOrEmpty(cellValue))
                DateTime.TryParseExact(cellValue, "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out newValue);

            // return new value
            return newValue.ToLongDateString();
        }

        static string ConvertGeo2En(string cellValue)
        {
            return cellValue
                .Replace("იან", "Jan")
                .Replace("Ian", "Jan")
                .Replace("თებ", "Feb")
                .Replace("მარ", "Mar")
                .Replace("აპრ", "Apr")
                .Replace("მაი", "May")
                .Replace("ივნ", "Jun")
                .Replace("ივლ", "Jul")
                .Replace("აგვ", "Aug")
                .Replace("სექ", "Sep")
                .Replace("ოქტ", "Oct")
                .Replace("ნოე", "Nov")
                .Replace("დეკ", "Dec")
                .Replace(",", "/")
                .Replace(".", "/")
                .Trim();
        }

        static string ConvertNumToMonth(string cellValue)
        {
            var monthNames = new string[] { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Nov", "Dec" };
            var parts = cellValue.Split('-');

            if (parts != null && parts.Length == 3)
            {
                var monthNum = -1;
                int.TryParse(parts[1], out monthNum);
                if (monthNum > -1)
                {
                    parts[1] = monthNames[monthNum];
                    cellValue = string.Join("-", parts);
                }
            }

            return cellValue;
        }

        static void GenerateSQLScripts(string filePath)
        {
            var script = $@"
USE [master] 
GO

sp_configure 'show advanced options', 1;
RECONFIGURE;
GO
sp_configure 'Ad Hoc Distributed Queries', 1;
RECONFIGURE;
GO

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1 
GO 

EXEC master.dbo.sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
GO 

/****** Object:  Database [ImportFromExcel]    Script Date: 6/23/2022 3:16:46 PM ******/
CREATE DATABASE [ImportFromExcel]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'ImportFromExcel', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\ImportFromExcel.mdf' , SIZE = 8192KB , MAXSIZE = UNLIMITED, FILEGROWTH = 65536KB )
 LOG ON 
( NAME = N'ImportFromExcel_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL15.MSSQLSERVER\MSSQL\DATA\ImportFromExcel_log.ldf' , SIZE = 8192KB , MAXSIZE = 2048GB , FILEGROWTH = 65536KB )
 WITH CATALOG_COLLATION = DATABASE_DEFAULT
GO


USE ImportFromExcel;
GO
SELECT * INTO Imported_Items
FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0', 'Excel 12.0; Database={$@"{Environment.CurrentDirectory}\corrected_file.xlsx"}', [Sheet1$]);
GO";

            File.WriteAllText("SQL_SCRIPT.sql", script);
        }
    }
}
