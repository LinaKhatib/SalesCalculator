using ClosedXML.Excel;

namespace SalesCalculator
{
    public static class TableCreationService
    {
        private readonly static int[] intColumns = { 7, 8, 11 };
        private readonly static int[] arrCol = { 4, 1, 1 };

        private static int posCatRow;
        private static int posCatCol;

        private static Dictionary<string, int[]> catigoryDictionary;
        private static List<string> listBigCat;

        public static void CreatTables(Dictionary<string, string> config, IXLWorksheet worksheet, int lastRow, List<Dictionary<string, string>> configTable)
        {
            catigoryDictionary = new Dictionary<string, int[]>();
            listBigCat = new List<string>();

            DrawTable(configTable, lastRow, worksheet);

            string colLetter = "";
            string posCat;

            for (int row = 2; row <= lastRow; row++)
            {
                string pos = worksheet.Cell(row, 2).GetValue<string>();

                if (!string.IsNullOrEmpty(pos) && config.ContainsKey(pos))
                {

                    RemoveSpaces(worksheet, row);

                    posCatRow = catigoryDictionary[config[pos]][0];
                    posCatCol = catigoryDictionary[config[pos]][1];

                    for (int i = 0; i < arrCol.Length; i++)
                    {
                        posCatCol += arrCol[i];
                        colLetter = XLHelper.GetColumnLetterFromNumber(posCatCol);
                        posCat = worksheet.Cell(posCatRow, posCatCol).GetValue<string>();

                        if (string.IsNullOrEmpty(posCat))
                        {
                            worksheet.Cell(posCatRow, posCatCol).FormulaA1 = $"={colLetter}{row}";
                        }

                        else
                            worksheet.Cell(posCatRow, posCatCol).FormulaA1 = $"{worksheet.Cell(posCatRow, posCatCol).FormulaA1}+{colLetter}{row}";

                    }
                }
            }

            CalculateTheTable(worksheet, lastRow);
        }

        private static void CalculateTheTable(IXLWorksheet worksheet, int lastRow)
        {

            var listSum = new List<string>();

            int lastSumCellRow = worksheet.LastRowUsed().RowNumber();

            int firstPosCatRow;
            int lastPosCatRow;

            string colLetter = "";

            for (int i = 0; i < listBigCat.Count; i++)
            {
                posCatRow = catigoryDictionary[listBigCat[i]][0];
                posCatCol = catigoryDictionary[listBigCat[i]][1] + 6;

                firstPosCatRow = posCatRow + 1;
                if (i < listBigCat.Count - 1)
                {
                    lastPosCatRow = catigoryDictionary[listBigCat[i + 1]][0] - 1;
                }
                else
                {
                    lastPosCatRow = lastSumCellRow;
                }

                colLetter = XLHelper.GetColumnLetterFromNumber(posCatCol);
                worksheet.Cell(posCatRow, posCatCol).FormulaA1 = $"=SUM({colLetter}{firstPosCatRow}:{colLetter}{lastPosCatRow})";

                string posSumCell = worksheet.Cell(lastSumCellRow + 1, posCatCol).GetValue<string>();

                if (string.IsNullOrEmpty(posSumCell))
                {
                    worksheet.Cell(lastSumCellRow + 1, posCatCol).FormulaA1 = $"={colLetter}{posCatRow}";
                }

                else
                    worksheet.Cell(lastSumCellRow + 1, posCatCol).FormulaA1 = $"{worksheet.Cell(lastSumCellRow + 1, posCatCol).FormulaA1}+{colLetter}{posCatRow}";

            }

            worksheet.Cell(lastRow + 1, posCatCol).FormulaA1 = $"=SUM({colLetter}2:{colLetter}{lastRow})";
            worksheet.Cell(lastSumCellRow + 2, posCatCol).FormulaA1 = $"={colLetter}{lastSumCellRow + 1} - {colLetter}{lastRow + 1}";
        }

        private static void RemoveSpaces(IXLWorksheet worksheet, int row)
        {
            for (int i = 0; i < intColumns.Length; i++)
            {
                var cell = worksheet.Cell(row, intColumns[i]);
                string rawValue = cell.GetString();

                string cleanValue = new string(rawValue.Where(c => !char.IsWhiteSpace(c)).ToArray());

                if (double.TryParse(cleanValue.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result))
                {
                    cell.Value = result;
                    cell.Style.NumberFormat.Format = "0.00";
                }
            }

        }

        private static void DrawTable(List<Dictionary<string, string>> configTable, int lastRow, IXLWorksheet worksheet)
        {

            if (configTable.Any())
            {
                int summaryStartRow = lastRow + 2;

                foreach (var conf in configTable)
                {
                    var cell = worksheet.Cell(int.Parse(conf["Row"]) + summaryStartRow, conf["Column"]);

                    if (conf.ContainsKey("Text") && !string.IsNullOrEmpty(conf["Text"]))
                    {
                        cell.Value = conf["Text"];

                        if (cell.Address.ColumnNumber == 2)
                        {
                            catigoryDictionary.Add(conf["Text"].ToString(), new int[] { cell.Address.RowNumber, cell.Address.ColumnNumber });
                        }
                    }

                    if (conf.TryGetValue("RgbColor", out string rgbStr) && long.TryParse(rgbStr, out long rgb))
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromColor(System.Drawing.Color.FromArgb((int)rgb));
                    }

                    if (conf.TryGetValue("IsBold", out string isBold) && bool.Parse(isBold))
                    {
                        cell.Style.Font.Bold = true;
                    }

                    if (conf.TryGetValue("HAlign", out string hAlignStr))
                    {
                        switch (hAlignStr)
                        {
                            case "-4152":
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                                break;
                            case "-4108":
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                break;
                            case "-4131":
                                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                listBigCat.Add(conf["Text"].ToString());
                                break;
                        }
                    }

                    if (conf.TryGetValue("AddBorders", out string addBorders) && bool.Parse(addBorders))
                    {
                        cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    }

                    if (conf.TryGetValue("IsEmpty", out string isEmpty) && bool.Parse(isEmpty))
                    {
                        cell.Clear(XLClearOptions.Contents);
                    }
                }
            }

        }
    }
}
