// Ignore Spelling: Xls

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.Configuration;

namespace Xls
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var data = new ConfigurationBuilder()
                .AddJsonFile("C:\\Users\\lucas\\source\\repos\\Xls\\Xls\\appsettings.json", true, true).Build();
            Dictionary<string, List<string>> values = GetRowValues(data["filePath"]!, data["bookName"]!, data, data["row"]!.Split(','));
            foreach(KeyValuePair<string, List<string>> pair in values)
            {
                Console.WriteLine();
                Console.WriteLine(pair.Key);
                foreach(string s in pair.Value)
                {
                    Console.WriteLine(s);
                }
                Console.WriteLine();
            }
        }

        static Sheets? GetAllWorkSheets(string fileName)
        {
            Sheets? theSheets = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                theSheets = document?.WorkbookPart?.Workbook.Sheets;
            }

            return theSheets;
        }

        static Dictionary<string, List<string>> GetRowValues(string fileName, string sheetName, IConfiguration data ,params string[] rowNames)
        {
            Dictionary<string, List<string>> dicValues = new Dictionary<string, List<string>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart? wbPart = document.WorkbookPart;
                Sheet? theSheet = wbPart?.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                if (theSheet is null || theSheet.Id is null)
                {
                    throw new ArgumentException("Spreadsheet not found, check file path");
                }
                for (int i = 0; i < rowNames.Length; i++)
                {
                    WorksheetPart? wsPart = (WorksheetPart)wbPart!.GetPartById(theSheet.Id!);
                    Row? theRow = wsPart.Worksheet?.Descendants<Row>()?.Where(r => r.RowIndex == rowNames[i]).FirstOrDefault();

                    List<Cell>? theCells = theRow?.Descendants<Cell>().Where(c => c.CellReference!.ToString()!.Contains(rowNames[i])).ToList();
                    List<string> theValues = new();
                    foreach (Cell c in theCells!)
                    {
                        string value = c.InnerText;
                        if (c.DataType?.Value == null)
                        {
                            if (c.StyleIndex != null)
                            {
                                int styleIndex = (int)c.StyleIndex!.Value;
                                CellFormat cellFormat = (CellFormat)wbPart.WorkbookStylesPart!.Stylesheet.CellFormats!.ElementAt(styleIndex);
                                uint formatId = cellFormat.NumberFormatId!.Value;

                                if (formatId == (uint)DataTypes.DateShort || formatId == (uint)DataTypes.DateLong)
                                {
                                    if (double.TryParse(c.InnerText, out double oaDate))
                                    {
                                        theValues.Add(DateTime.FromOADate(oaDate).ToShortDateString());
                                    }
                                }
                                else if (formatId == (uint)DataTypes.Percentage)
                                {
                                    theValues.Add(Math.Round(double.Parse(c.InnerText), 4, MidpointRounding.AwayFromZero) * 100 + "%");
                                }
                                else if (formatId == (uint)DataTypes.Currency)
                                {
                                    theValues.Add("£" + double.Parse(c.InnerText));
                                }
                                else
                                {
                                    theValues.Add(c.InnerText);
                                }
                            }
                            else if(data["showBlankCells"] == "true")
                            {
                                theValues.Add("Blank Cell");
                            }
                        }
                        else if (c.DataType?.Value == CellValues.SharedString)
                        {
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (stringTable is not null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                theValues.Add(value);
                            }
                        }
                        else if (c.DataType?.Value == CellValues.Boolean)
                        {
                            switch (value)
                            {
                                case "0":
                                    theValues.Add("FALSE");
                                    break;
                                default:
                                    theValues.Add("TRUE");
                                    break;
                            }
                        }
                    }
                    dicValues.Add($"Row {rowNames[i]}:", theValues);
                }
            }
            return dicValues;
        }
    }
}