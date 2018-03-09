using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace openxmltest
{
    class Program
    {
        static void Main(string[] args)
        {
            Read();

            BuildWorkbook("openxml2.xlsx");

            Console.WriteLine("\nPress 任意键退出");
            Console.ReadLine();
        }

        private static void Read()
        {
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open("read.xlsx", false))
            {
                WorkbookPart workbookpart = doc.WorkbookPart;

                //获取SharedStringTable
                SharedStringTablePart sharedStringTablePart = workbookpart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

                //获取worksheet
                WorksheetPart worksheetPart = workbookpart.WorksheetParts.First();
                Worksheet worksheet = worksheetPart.Worksheet;

                var cells = worksheet.Descendants<Cell>();
                var rows = worksheet.Descendants<Row>();


                Console.WriteLine("One way: go through each cell in the sheet");
                foreach (Cell cell in cells)
                {
                    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                    {
                        int ssid = int.Parse(cell.CellValue.Text);
                        string str = sharedStringTable.ChildElements[ssid].InnerText;
                        Console.WriteLine("({0}){1}\t", ssid, str);
                    }
                    else if (cell.CellValue != null)
                    {
                        Console.WriteLine("{0}\t", cell.CellValue.Text);
                    }
                }

                Console.WriteLine("Or... via each row");
                foreach (Row row in rows)
                {
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(c.CellValue.Text);
                            string str = sharedStringTable.ChildElements[ssid].InnerText;
                            Console.Write("({0}){1}\t", ssid, str);
                        }
                        else if (c.CellValue != null)
                        {
                            Console.Write("{0}\t", c.CellValue.Text);
                        }
                    }
                    Console.Write(Environment.NewLine);
                }
            }
        }

        private static void BuildWorkbook(string filename)
        {
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookpart = doc.AddWorkbookPart();
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    WorkbookStylesPart workbookStylesPart = workbookpart.AddNewPart<WorkbookStylesPart>();

                    workbookpart.Workbook = new Workbook();
                    worksheetPart.Worksheet = new Worksheet();
                    workbookStylesPart.Stylesheet = CreateStylesheet();

                    Sheets sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Columns columns = worksheetPart.Worksheet.AppendChild(new Columns());
                    //顺序在columns之后
                    SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                    #region sheets

                    var sheet = sheets.AppendChild(new Sheet
                    {
                        Name = "产品",
                        SheetId = 1,
                        Id = doc.WorkbookPart.GetIdOfPart(worksheetPart),
                    });

                    #endregion

                    #region cols

                    columns.Append(new Column()
                    {
                        Min = 1,
                        Max = 1,
                        Width = 38.22d,
                        CustomWidth = true
                    });
                    columns.Append(CreateColumnData(2, 4, 20));
                    columns.Append(CreateColumnData(6, 6, 6.5703125));

                    #endregion

                    #region autofilter

                    worksheetPart.Worksheet.AppendChild(new AutoFilter
                    {
                        Reference = "B1"
                    });

                    #endregion autofilter

                    #region data

                    Row r;
                    Cell c;

                    // header
                    r = new Row();
                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "A1";
                    c.CellValue = new CellValue("产品ID");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "B1";
                    c.CellValue = new CellValue("产品说明");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "C1";
                    c.CellValue = new CellValue("订单说明");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "D1";
                    c.CellValue = new CellValue("折扣说明");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "E1";
                    c.CellValue = new CellValue("货币");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "F1";
                    c.CellValue = new CellValue("费用");
                    r.Append(c);
                    sheetData.Append(r);

                    // content
                    r = new Row();
                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "A2";
                    c.CellValue = new CellValue("15D05473-742F-4691-8BF4-5124F2D66176");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "B2";
                    c.CellValue = new CellValue("Iced Lemon Tea");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "C2";
                    c.CellValue = new CellValue("Special Iced Lemon Tea");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "D2";
                    c.CellValue = new CellValue("Iced Lemon Tea (50% off)");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.String;
                    c.CellReference = "E2";
                    c.CellValue = new CellValue("USD");
                    r.Append(c);

                    c = new Cell();
                    c.StyleIndex = 3;
                    c.DataType = CellValues.Number;
                    c.CellReference = "F2";
                    c.CellValue = new CellValue("5.95");
                    r.Append(c);

                    c = new Cell();
                    c.DataType = CellValues.Boolean;
                    c.CellReference = "G2";
                    c.CellValue = new CellValue("0");
                    r.Append(c);
                    sheetData.Append(r);

                    #endregion


                    doc.Save();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadLine();
            }
        }

        private static Column CreateColumnData(UInt32 StartColumnIndex, UInt32 EndColumnIndex, double ColumnWidth)
        {
            Column column;
            column = new Column();
            column.Min = StartColumnIndex;
            column.Max = EndColumnIndex;
            column.Width = ColumnWidth;
            column.CustomWidth = true;
            return column;
        }

        private static Stylesheet CreateStylesheet()
        {
            Stylesheet ss = new Stylesheet();

            Fonts fts = new Fonts();
            DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
            FontName ftn = new FontName();
            ftn.Val = "Calibri";
            FontSize ftsz = new FontSize();
            ftsz.Val = 11;
            ft.FontName = ftn;
            ft.FontSize = ftsz;
            fts.Append(ft);
            fts.Count = (uint)fts.ChildElements.Count;

            Fills fills = new Fills();
            Fill fill;
            PatternFill patternFill;
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.None;
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fill = new Fill();
            patternFill = new PatternFill();
            patternFill.PatternType = PatternValues.Gray125;
            fill.PatternFill = patternFill;
            fills.Append(fill);
            fills.Count = (uint)fills.ChildElements.Count;

            Borders borders = new Borders();
            Border border = new Border();
            border.LeftBorder = new LeftBorder();
            border.RightBorder = new RightBorder();
            border.TopBorder = new TopBorder();
            border.BottomBorder = new BottomBorder();
            border.DiagonalBorder = new DiagonalBorder();
            borders.Append(border);
            borders.Count = (uint)borders.ChildElements.Count;

            CellStyleFormats csfs = new CellStyleFormats();
            CellFormat cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            csfs.Append(cf);
            csfs.Count = (uint)csfs.ChildElements.Count;

            uint iExcelIndex = 164;
            NumberingFormats nfs = new NumberingFormats();
            CellFormats cfs = new CellFormats();

            cf = new CellFormat();
            cf.NumberFormatId = 0;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cfs.Append(cf);

            NumberingFormat nf;
            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = "dd/mm/yyyy hh:mm:ss";
            nfs.Append(nf);
            cf = new CellFormat();
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = "#,##0.0000";
            nfs.Append(nf);
            cf = new CellFormat();
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            // #,##0.00 is also Excel style index 4
            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = "#,##0.00";
            nfs.Append(nf);
            cf = new CellFormat();
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            // @ is also Excel style index 49
            nf = new NumberingFormat();
            nf.NumberFormatId = iExcelIndex++;
            nf.FormatCode = "@";
            nfs.Append(nf);
            cf = new CellFormat();
            cf.NumberFormatId = nf.NumberFormatId;
            cf.FontId = 0;
            cf.FillId = 0;
            cf.BorderId = 0;
            cf.FormatId = 0;
            cf.ApplyNumberFormat = true;
            cfs.Append(cf);

            nfs.Count = (uint)nfs.ChildElements.Count;
            cfs.Count = (uint)cfs.ChildElements.Count;

            CellStyles css = new CellStyles();
            CellStyle cs = new CellStyle();
            cs.Name = "Normal";
            cs.FormatId = 0;
            cs.BuiltinId = 0;
            css.Append(cs);
            css.Count = (uint)css.ChildElements.Count;

            DifferentialFormats dfs = new DifferentialFormats();
            dfs.Count = 0;

            TableStyles tss = new TableStyles();
            tss.Count = 0;
            tss.DefaultTableStyle = "TableStyleMedium9";
            tss.DefaultPivotStyle = "PivotStyleLight16";

            ss.Append(nfs);
            ss.Append(fts);
            ss.Append(fills);
            ss.Append(borders);
            ss.Append(csfs);
            ss.Append(cfs);
            ss.Append(css);
            ss.Append(dfs);

            return ss;
        }
    }
}
