using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace SalesCalculator
{
    public static class ProcessExcelService
    {

        public static void ProcessExcel(string xlsPath, string configPath, List<Dictionary<string, string>> configTable)
        {
            var config = Config.LoadConfig(configPath);

            using (var workbook = new XLWorkbook(xlsPath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var lastRow = worksheet.LastRowUsed().RowNumber();

                    PositionCheckingService.CheckPosition(configPath, config, worksheet, lastRow);

                    TableCreationService.CreatTables(config, worksheet, lastRow, configTable);
                }

                string directory = Path.GetDirectoryName(xlsPath) ?? "";
                string fileName = Path.GetFileNameWithoutExtension(xlsPath);
                string extension = Path.GetExtension(xlsPath);

                string newPath = Path.Combine(directory, $"{fileName}_Report{extension}");

                workbook.SaveAs(newPath);

                Console.WriteLine("\nИзменения сохранены в новый (_Report) файл.");
            }
        }
    }
}
