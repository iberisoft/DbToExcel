using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace DbToExcel
{
    static class Program
    {
        static IConfiguration m_Config;

        static void Main()
        {
            m_Config = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build();
            Convert();
        }

        private static void Convert()
        {
            var filePath = "Test.xlsx";
            var bindingFilePath = "Binding.xml";

            using (var connection = new SqlConnection(m_Config["ConnectionString"]))
            {
                connection.Open();
                using (var command = connection.CreateCommand())
                {
                    command.CommandText = m_Config["QueryText"];
                    using (var reader = command.ExecuteReader())
                    {
                        var rows = ReadRows(reader).ToArray();
                        Console.WriteLine($"{rows.Length} rows read");
                        SaveToExcelFile(rows, filePath, bindingFilePath);
                        Console.WriteLine($"Saved to file '{filePath}' by using binding file '{bindingFilePath}'");
                    }
                }
            }

            var info = new ProcessStartInfo(filePath);
            info.UseShellExecute = true;
            Process.Start(info);
        }

        private static IEnumerable<Dictionary<string, object>> ReadRows(SqlDataReader reader)
        {
            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (var i = 0; i < reader.FieldCount; ++i)
                {
                    row.Add(reader.GetName(i), reader.GetValue(i));
                }
                yield return row;
            }
        }

        private static void SaveToExcelFile(Dictionary<string, object>[] rows, string filePath, string bindingFilePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var workbookBinding = new WorkbookBinding(XDocument.Load(bindingFilePath));
                PopulateWorkbook(rows, package.Workbook, workbookBinding);
                package.Save();
            }
        }

        private static void PopulateWorkbook(Dictionary<string, object>[] rows, ExcelWorkbook workbook, WorkbookBinding workbookBinding)
        {
            foreach (var worksheetBinding in workbookBinding.Worksheets)
            {
                var worksheet = workbook.Worksheets[worksheetBinding.Name];
                if (worksheet != null)
                {
                    PopulateWorksheet(rows, worksheet, worksheetBinding);
                }
            }
        }

        private static void PopulateWorksheet(Dictionary<string, object>[] rows, ExcelWorksheet worksheet, WorksheetBinding worksheetBinding)
        {
            foreach (var cellBinding in worksheetBinding.Cells)
            {
                ExcelRangeBase cell = worksheet.Cells[cellBinding.Name];
                foreach (var row in rows)
                {
                    var value = row[cellBinding.Source];
                    cell.Value = cellBinding.Format != null && value is IFormattable formattable ? formattable.ToString(cellBinding.Format, null) : value.ToString();
                    cell = cell.Offset(1, 0);
                }
            }
        }
    }
}
