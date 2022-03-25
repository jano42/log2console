using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Log2Console
{
    public static class utils
    {
        public static void Export2Excel(ListView listView, string FileName)
        {
            CreateSpreadsheetWorkbook(listView, FileName, true);
        }

        private static SheetData ListViewToExcelSheet(ListView listView)
        {
            var sheetData = new SheetData();

            var row = new Row { RowIndex = 1 };
            int column = 1;
            int timeColumnIndex = -1;
            foreach (ColumnHeader ch in listView.Columns)
            {
                row.Append(new Cell { CellReference = GetExcelColumnName(column++) + row.RowIndex, DataType = CellValues.String, CellValue = new CellValue(ch.Text) });
                if (ch.Text == "Time")
                    timeColumnIndex = column - 1;//-1 because of previous ++
            }
            sheetData.Append(row);

            foreach (ListViewItem lvi in listView.Items)
            {
                column = 1;
                row = new Row { RowIndex = row.RowIndex + 1 };
                foreach (ListViewItem.ListViewSubItem lvs in lvi.SubItems)
                {
                    if (column == timeColumnIndex)
                    {
                        DateTime t = DateTime.Now;
                        if (DateTime.TryParse(lvs.Text, out t))
                        {
                            row.Append(new Cell { CellReference = GetExcelColumnName(column++) + row.RowIndex, DataType = CellValues.Date, CellValue = new CellValue(t), StyleIndex = 1});
                        }
                        else
                        {
                            row.Append(new Cell { CellReference = GetExcelColumnName(column++) + row.RowIndex, DataType = CellValues.String, CellValue = new CellValue(lvs.Text) });
                        }
                    }
                    else
                    {
                        row.Append(new Cell { CellReference = GetExcelColumnName(column++) + row.RowIndex, DataType = CellValues.String, CellValue = new CellValue(lvs.Text) });
                    }
                }
                sheetData.Append(row);
            }
            return sheetData;
        }


        private static Sheet CreateSheet(string name, UInt32Value id, SheetData data, WorkbookPart workbookPart)
        {
            // Add a WorksheetPart to the WorkbookPart.
            var informationWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            informationWorksheetPart.Worksheet = new Worksheet(data);
            
            // Append a new worksheet and associate it with the workbook.
            var informationSheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(informationWorksheetPart),
                SheetId = id,
                Name = name

            };

            return informationSheet;
        }

        public static void CreateSpreadsheetWorkbook(ListView listView, string filename, bool open = true)
        {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            var spreadsheetDocument = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);

            // Add a WorkbookPart to the document.
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            //IEnumerable<Row> rows = worksheetPart.Worksheet.Descendants<Row>();



            if (spreadsheetDocument.WorkbookPart != null)
            {
                var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet
                {
                    // blank font list
                    Fonts = new Fonts
                    {
                        Count = 1
                    }
                };

                stylesPart.Stylesheet.Fonts.AppendChild(new Font());

                // create fills
                stylesPart.Stylesheet.Fills = new Fills();

                // create a solid red fill
                var solidRed = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }, // red fill
                    BackgroundColor = new BackgroundColor { Indexed = 64 }
                };

                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } }); // required, reserved by Excel
                stylesPart.Stylesheet.Fills.AppendChild(new Fill { PatternFill = solidRed });
                stylesPart.Stylesheet.Fills.Count = 3;

                // blank border list
                stylesPart.Stylesheet.Borders = new Borders
                {
                    Count = 1
                };
                stylesPart.Stylesheet.Borders.AppendChild(new Border());

                // blank cell format list
                stylesPart.Stylesheet.CellStyleFormats = new CellStyleFormats
                {
                    Count = 1
                };
                stylesPart.Stylesheet.CellStyleFormats.AppendChild(new CellFormat());

                stylesPart.Stylesheet.NumberingFormats = new NumberingFormats();
                var formatDateTime = new NumberingFormat
                {
                    NumberFormatId = 2,
                    FormatCode = StringValue.FromString($"dd/mm/yyyy hh:mm:ss{System.Threading.Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator}000")
                }; 
                stylesPart.Stylesheet.NumberingFormats.AppendChild(formatDateTime);
                stylesPart.Stylesheet.NumberingFormats.Count = 1;



                // cell format list
                stylesPart.Stylesheet.CellFormats = new CellFormats();
                // empty one for index 0, seems to be required
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat());
                // cell format references style format 0, font 0, border 0, fill 2 and applies the fill
                stylesPart.Stylesheet.CellFormats.AppendChild(new CellFormat { FormatId = 0, ApplyNumberFormat = true, NumberFormatId = 2, FontId = 0, BorderId = 0, FillId = 0, ApplyFill = false }).AppendChild(new Alignment { Horizontal = HorizontalAlignmentValues.Center });
                stylesPart.Stylesheet.CellFormats.Count = 2;

                stylesPart.Stylesheet.Save();

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                sheets.Append(CreateSheet("Logs", 1, ListViewToExcelSheet(listView), workbookPart));
            }

            workbookPart.Workbook.Save();

            // Close the document.
            spreadsheetDocument.Close();


            if (open)
            {
                //open it
                Process.Start(new ProcessStartInfo { Arguments = filename, FileName = "explorer.exe" });
            }
        }

        private static string GetExcelColumnName(int columnNumber)
        {
            var columnName = "";

            while (columnNumber > 0)
            {
                var modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

    }
}