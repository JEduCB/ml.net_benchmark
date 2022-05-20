using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace oneDalTest.Helpers
{
    internal class ExcelDocument
    {
        public SpreadsheetDocument Document { get; set; }
        public SheetData SheetData{ get; set; }
        public WorkbookPart WorkbookPart { get; set; }
        public WorksheetPart WorksheetPart { get; set; }
    }

    internal static class ExcelExporter
    {
        public static ExcelDocument CreateExcelDocument(string xlsxFile, string task, string dataset)
        {
            SpreadsheetDocument document = SpreadsheetDocument.Create(xlsxFile, SpreadsheetDocumentType.Workbook);
            WorkbookPart workbookPart = document.AddWorkbookPart();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            WorkbookStylesPart stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            var objects = GenerateStyleSheet(worksheetPart);
            stylesPart.Stylesheet = objects.Item1;
            stylesPart.Stylesheet.Save();

            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Append(objects.Item2);

            workbookPart.Workbook = new Workbook();

            Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = $"{task} - {dataset}" };

            workbookPart.Workbook.CalculationProperties = new CalculationProperties()
            {
                ForceFullCalculation = true,
                FullCalculationOnLoad = true,
                FullPrecision = true,
            };

            sheets.Append(sheet);

            return new ExcelDocument { Document = document, SheetData = sheetData, WorkbookPart = workbookPart, WorksheetPart = worksheetPart };
        }

        static int _rowIndex = 1;
        static List<string> _cellFormulas = null;
        public static void Export(ExcelDocument excelDoc, List<string> rows, bool calculateSpeedUp)
        {
            int from = 0;
            int to = 0;

            foreach (var row in rows)
            {
                if (to > 0 && !string.IsNullOrEmpty(row))
                {
                    ++to;
                }

                if (row.StartsWith("1,"))
                {
                    from = _rowIndex;
                    to = _rowIndex;
                }

                AppendRow(excelDoc.SheetData, row, _rowIndex++, row.StartsWith("Run,") ? 8U : 0U);
            }

            string[] formulasRow =
            {
                ",,,,Min",
                ",,,,Max",
                ",,,,Median",
                ",,,,Mean",
                ",,,,StDev",
            };

            UInt32Value[] fillColor = { 2U, 3U, 4U, 5U, 6U };

            string[] formulas =
            {
                ",=Min({0}:{1})",
                ",=Max({0}:{1})",
                ",=Median({0}:{1})",
                ",=Average({0}:{1})",
                ",=STDEVP({0}:{1})",
            };

            bool saveCellFormulas = _cellFormulas == null;
            if (saveCellFormulas)
            {
                _cellFormulas = new List<string>();
            }

            int formulaIndex = 0;
            foreach (var row in formulasRow)
            {
                var formula = row;

                for (int columnIndex = 6; columnIndex <= 10; columnIndex++)
                {
                    formula += string.Format(formulas[formulaIndex], GetCellReference(columnIndex, from),
                        GetCellReference(columnIndex, to));

                    if (saveCellFormulas)
                    {
                        _cellFormulas.Add(GetCellReference(columnIndex, _rowIndex));
                    }
                }

                AppendRow(excelDoc.SheetData, formula, _rowIndex++, fillColor[formulaIndex++], true);
            }

            if (calculateSpeedUp)
            {
                List<string> _conditionalFormatCells = new List<string>();

                for (int i = 0; i < formulasRow.Count() - 1; i++)
                {
                    string speedUp = ",,,,Speedup";

                    for (int columnIndex = 6; columnIndex <= 10; columnIndex++)
                    {
                        var cellReference = GetCellReference(columnIndex, _rowIndex - formulasRow.Count());
                        speedUp += ",=" + _cellFormulas[0] + "/" + cellReference;
                        _cellFormulas.RemoveRange(0, 1);
                        _conditionalFormatCells.Add(GetCellReference(columnIndex, _rowIndex));
                    }

                    AppendRow(excelDoc.SheetData, speedUp, _rowIndex++, 7U, true);
                }

                AddConditionalFormat(excelDoc, _conditionalFormatCells);
            }

            AppendRow(excelDoc.SheetData, "", _rowIndex++, 0U);
        }

        private static void AppendRow(SheetData sheetData, string value, int rowIndex, UInt32Value styleIndex,
            bool isFormula = false)
        {
            string[] columns = value.Split(",");

            Row row = new Row();
            int columnIndex = 1;

            foreach (var column in columns)
            {
                Cell cell = CreateCell(column, columnIndex++, rowIndex, styleIndex, isFormula);
                row.AppendChild(cell);
            }

            sheetData.AppendChild(row);
        }

        private static Cell CreateCell(string column, int columnIndex, int rowIndex, UInt32Value styleIndex,
            bool isFormula)
        {
            CellValues cellValueType = CellValues.Number;
            CellFormula formula = null;
            CellValue cellValue = null;

            if (double.TryParse(column, out double dValue) && column.Contains('.'))
            {
                cellValue = new CellValue(dValue.ToString("0.###"));
            }
            else if (int.TryParse(column, out int iValue))
            {
                cellValue = new CellValue(column);
            }
            else if (isFormula && column.StartsWith("="))
            {
                column = column.Replace("=", "");
                formula = new CellFormula(column);
                cellValueType = CellValues.Number;
                cellValue = null;
            }
            else
            {
                cellValue = new CellValue(column);
                cellValueType = CellValues.String;

                if (isFormula && string.IsNullOrEmpty(column))
                {
                    styleIndex = 1U;
                }
            }

            Cell cell = new Cell
            {
                CellReference = GetCellReference(columnIndex, rowIndex),
                DataType = cellValueType,
                CellValue = cellValue,
                CellFormula = formula,
                StyleIndex = styleIndex,
            };

            return cell;
        }

        private static string GetCellReference(int columnIndex, int rowIndex)
        {
            string columnName = "";

            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }

            return columnName + rowIndex;
        }


        public static (Stylesheet, ConditionalFormatting) GenerateStyleSheet(WorksheetPart worksheetPart)
        {
            Stylesheet stylesheet = new Stylesheet()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" }
            };

            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //FontId 0 = Default
            Font font0 = new Font();
            FontSize fontSize0 = new FontSize() { Val = 11D };
            Color color0 = new Color() { Theme = 1U };
            FontName fontName0 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering0 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme0 = new FontScheme() { Val = FontSchemeValues.Minor };

            font0.Append(fontSize0);
            font0.Append(color0);
            font0.Append(fontName0);
            font0.Append(fontFamilyNumbering0);
            font0.Append(fontScheme0);

            //FontId 1 = Header/Yellow/Bold
            Font font1 = new Font(new Bold());
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Fonts fonts = new Fonts() { Count = 2U, KnownFonts = true };
            fonts.Append(font0);
            fonts.Append(font1);

            // FillId = 0
            PatternFill patternFill0 = new PatternFill() { PatternType = PatternValues.None };
            Fill fill0 = new Fill();
            fill0.Append(patternFill0);

            // FillId = 1
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.Gray125 };
            Fill fill1 = new Fill();
            fill1.Append(patternFill1);

            // FillId = 2, Gold Accent 4 Lighter 80%
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFFF2CC" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = 64U };
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid };
            patternFill2.Append(foregroundColor2);
            patternFill2.Append(backgroundColor2);
            Fill fill2 = new Fill();
            fill2.Append(patternFill2);

            // FillId = 3, Gold Accent 4 Lighter 60%
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFFFE699" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = 64U };
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            patternFill3.Append(foregroundColor3);
            patternFill3.Append(backgroundColor3);
            Fill fill3 = new Fill();
            fill3.Append(patternFill3);

            // FillId = 4, Gold Accent 4 Lighter 40%
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FFFFD966" };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = 64U };
            patternFill4.Append(foregroundColor4);
            patternFill4.Append(backgroundColor4);
            Fill fill4 = new Fill();
            fill4.Append(patternFill4);

            // FillId = 5, Gold Accent 4
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "FFFFC000" };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = 64U };
            patternFill5.Append(foregroundColor5);
            patternFill5.Append(backgroundColor5);
            Fill fill5 = new Fill();
            fill5.Append(patternFill5);

            // FillId = 6, Gold Accent 4 Darker 25%
            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFBF8F00" };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = 64U };
            patternFill6.Append(foregroundColor6);
            patternFill6.Append(backgroundColor6);
            Fill fill6 = new Fill();
            fill6.Append(patternFill6);

            // FillId = 7, Orange Accent 2 Darker 25%
            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor7 = new ForegroundColor() { Rgb = "FFC65911" };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = 64U };
            patternFill7.Append(foregroundColor7);
            patternFill7.Append(backgroundColor7);
            Fill fill7 = new Fill();
            fill7.Append(patternFill7);

            // FillId = 8, Black
            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor8 = new ForegroundColor() { Rgb = "00000000" };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = 64U };
            patternFill8.Append(foregroundColor8);
            patternFill8.Append(backgroundColor8);
            Fill fill8 = new Fill();
            fill8.Append(patternFill8);

            Fills fills = new Fills() { Count = 9U };
            fills.Append(fill0);
            fills.Append(fill1);
            fills.Append(fill2);
            fills.Append(fill3);
            fills.Append(fill4);
            fills.Append(fill5);
            fills.Append(fill6);
            fills.Append(fill7);
            fills.Append(fill8);

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Borders borders = new Borders() { Count = 1U };
            borders.Append(border1);

            CellFormat cellFormat = new CellFormat()
            {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U
            };

            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = 1U };
            cellStyleFormats.Append(cellFormat);

            CellFormat cellFormat0 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 0U, BorderId = 0U, FormatId = 0U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 1U, BorderId = 0U, FormatId = 0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 2U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 3U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 4U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 5U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 6U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = 0U, FontId = 0U, FillId = 7U, BorderId = 0U, FormatId = 0U, ApplyFill = true };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = 0U, FontId = 1U, FillId = 8U, BorderId = 0U, FormatId = 0U, ApplyFill = true, ApplyFont = true };

            CellFormats cellFormats = new CellFormats() { Count = 9U };
            cellFormats.Append(cellFormat0);
            cellFormats.Append(cellFormat1);
            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);
            cellFormats.Append(cellFormat6);
            cellFormats.Append(cellFormat7);
            cellFormats.Append(cellFormat8);

            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = 0U, BuiltinId = 0U };
            CellStyles cellStyles = new CellStyles() { Count = 1U };
            cellStyles.Append(cellStyle1);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);

            //Conditional Formatting Styles
            //Conditional Formatting FontId 0 = Green
            Font cfFont0 = new Font();
            Color cfColor0 = new Color() { Rgb = "FF006100" };
            cfFont0.Append(cfColor0);

            // Conditional Formatting FillId = 0, Light Green
            PatternFill cfPatternFill0 = new PatternFill() { PatternType = PatternValues.Solid };
            BackgroundColor cfBackgroundColor0 = new BackgroundColor() { Rgb = "FFC6EFCE" };
            cfPatternFill0.Append(cfBackgroundColor0);
            Fill cfFill0 = new Fill();
            cfFill0.Append(cfPatternFill0);

            //Conditional Formatting FontId 1 = Red
            Font cfFont1 = new Font();
            Color cfColor1 = new Color() { Rgb = "FF9C0006" };
            cfFont1.Append(cfColor1);

            // Conditional Formatting FillId = 1, Light Red
            PatternFill cfPatternFill1 = new PatternFill() { PatternType = PatternValues.Solid };
            BackgroundColor cfBackgroundColor1 = new BackgroundColor() { Rgb = "FFFFC7CE" };
            cfPatternFill1.Append(cfBackgroundColor1);
            Fill cfFill1 = new Fill();
            cfFill1.Append(cfPatternFill1);

            //Conditional Formatting FontId 2 = Brown
            Font cfFont2 = new Font();
            Color cfColor2 = new Color() { Rgb = "FF9C5700" };
            cfFont2.Append(cfColor2);

            // Conditional Formatting FillId = 0, Light Yellow
            PatternFill cfPatternFill2 = new PatternFill() { PatternType = PatternValues.Solid };
            BackgroundColor cfBackgroundColor2 = new BackgroundColor() { Rgb = "FFFFEB9C" };
            cfPatternFill2.Append(cfBackgroundColor2);
            Fill cfFill2 = new Fill();
            cfFill2.Append(cfPatternFill2);

            DifferentialFormat differentialFormat0 = new DifferentialFormat();
            differentialFormat0.Append(cfFont0);
            differentialFormat0.Append(cfFill0);

            DifferentialFormat differentialFormat1 = new DifferentialFormat();
            differentialFormat1.Append(cfFont1);
            differentialFormat1.Append(cfFill1);

            DifferentialFormat differentialFormat2 = new DifferentialFormat();
            differentialFormat2.Append(cfFont2);
            differentialFormat2.Append(cfFill2);

            //grab the differential formats part so we can add the style to apply
            DifferentialFormats differentialFormats = stylesheet.GetFirstChild<DifferentialFormats>();
            if (differentialFormats == null)
            {
                differentialFormats = new DifferentialFormats() { Count = 3U };
            }

            differentialFormats.Append(differentialFormat0);
            differentialFormats.Append(differentialFormat1);
            differentialFormats.Append(differentialFormat2);

            //add the differential formats to the stylesheet 
            stylesheet.Append(differentialFormats);

            //create the formulas
            Formula formula0 = new Formula() { Text = "1.000" };
            Formula formula1 = new Formula() { Text = "1.000" };
            Formula formula2 = new Formula() { Text = "1.000" };

            //create the conditional formatting rules with a type of Expression
            ConditionalFormattingRule conditionalFormattingRule0 = new ConditionalFormattingRule()
            {
                Type = ConditionalFormatValues.CellIs,
                FormatId = 0U,
                Priority = 1,
                Operator = ConditionalFormattingOperatorValues.GreaterThan
            };

            ConditionalFormattingRule conditionalFormattingRule1 = new ConditionalFormattingRule()
            {
                Type = ConditionalFormatValues.CellIs,
                FormatId = 1U,
                Priority = 2,
                Operator = ConditionalFormattingOperatorValues.LessThan
            };

            ConditionalFormattingRule conditionalFormattingRule2 = new ConditionalFormattingRule()
            {
                Type = ConditionalFormatValues.CellIs,
                FormatId = 2U,
                Priority = 3,
                Operator = ConditionalFormattingOperatorValues.Equal
            };

            //append the formulas to the rules
            conditionalFormattingRule0.Append(formula0);
            conditionalFormattingRule1.Append(formula1);
            conditionalFormattingRule2.Append(formula2);

            //create the conditional format reference
            ConditionalFormatting conditionalFormatting = new ConditionalFormatting()
            {
                SequenceOfReferences = new ListValue<StringValue>()
                {
                    InnerText = "F38:J41"
                }
            };

            //append the formatting rules to the formatting collection
            conditionalFormatting.Append(conditionalFormattingRule0);
            conditionalFormatting.Append(conditionalFormattingRule1);
            conditionalFormatting.Append(conditionalFormattingRule2);

            return (stylesheet, conditionalFormatting);
        }

        private static void AddConditionalFormat(ExcelDocument excelDoc, List<string> conditionalFormatCells)
        {
            var conditionalFormattingRange =
                $"{conditionalFormatCells[0]}:{conditionalFormatCells[conditionalFormatCells.Count() - 1]}";

            var SequenceOfReferences = new ListValue<StringValue>()
            {
                InnerText = conditionalFormattingRange
            };

            excelDoc.WorksheetPart.Worksheet.GetFirstChild<ConditionalFormatting>().SequenceOfReferences = SequenceOfReferences;
        }

    }
}
