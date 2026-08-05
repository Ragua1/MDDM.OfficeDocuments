using System.Collections.Generic;
using System.Globalization;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Styles;
using OfficeDocuments.Excel.VerificationTests.Properties;
using OfficeDocuments.Excel.TestKit;
using OfficeDocuments.Excel.TestKit.Validation;
using Color = System.Drawing.Color;

namespace OfficeDocuments.Excel.VerificationTests;

public class RealisticWorkbookTests : SpreadsheetTestBase
{
    public static readonly Random Rnd = new Random();

    [Fact]
    public void MinimalWorkbook_IsValidAndHasOneSheet()
    {
        var filePath = GetFilepath("minimal.xlsx");
        using (var w = CreateNewSpreadsheet(filePath))
        {
            w.AddWorksheet();
        }

        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        Assert.Single(reopened.GetWorksheetsName());
    }

    [Fact]
    public void MultiSheetStyledReport_IsValidAndReadable()
    {
        var filePath = GetFilepath("multi-sheet-styled-report.xlsx");
        using var w = CreateNewSpreadsheet(filePath);
        var s = w.CreateStyle(
            new Font { FontSize = 10, Color = Color.Black, FontName = FontNameValues.Arial },
            new Fill(System.Drawing.ColorTranslator.FromHtml("#FFFF99"))
        );
        var s_mediumBorder_all = w.CreateStyle(border: new Border(BorderStyleValues.Medium));
        var s_mediumBorder_rl = w.CreateStyle(
            border: new Border
            {
                Left = BorderStyleValues.Medium,
                Right = BorderStyleValues.Medium
            }
        );
        var s_fill_white = w.CreateStyle(
            fill: new Fill(Color.White)
        );


        // INFO sheet
        var ws_info = w.AddWorksheet("INFO", s);

        var s_bolt = w.CreateStyle(
            new Font { Bold = true }
        );
        var s_font_blue = w.CreateStyle(
            new Font { ArgbHexColor = "#2A66FF" }
        );

        var s_specBorder = w.CreateStyle(
            border: new Border
            {
                Top = BorderStyleValues.Thin,
                Left = BorderStyleValues.Medium,
                Bottom = BorderStyleValues.Thin,
                Right = BorderStyleValues.Medium
            }
        );

        var s_undeline = w.CreateStyle(new Font { Underline = UnderlineValues.Single });
        var timeFormat = "d.M.yyyy";

        ws_info.AddCell("Datum", s_bolt.CreateMergedStyle(s_mediumBorder_all));
        ws_info.AddCell(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture), s_mediumBorder_all);

        ws_info.AddStyle(s_font_blue);
        CreateRow(ws_info, "Verzia", s_bolt, 2, s_specBorder);
        CreateRow(ws_info, "E-Mail", s_bolt, "jan.novak@example.com", s_font_blue.CreateMergedStyle(s_specBorder).CreateMergedStyle(s_undeline));
        CreateRow(ws_info, "Meno", s_bolt, "Jan Novák", s_specBorder);
        CreateRow(ws_info, "Telefon", s_bolt, 555123456, s_specBorder);
        CreateRow(ws_info, "Uzivatel", s_bolt, "jan.novak", s_specBorder);
        CreateRow(ws_info, "Typ Nominacie", s_bolt, "TYZDENNA", s_specBorder);

        var s_mediumBorder_top = w.CreateStyle(border: new Border { Top = BorderStyleValues.Medium });
        ws_info.AddRow(s_fill_white);

        for (uint i = 1; i <= ws_info.GetRow(1).Cells.Count; i++)
        {
            ws_info.AddCell(s_mediumBorder_top);
            ws_info.SetColumnWidth(i, 22);
        }

        //TYDENNA sheet
        var s_font_red = w.CreateStyle(
            new Font { Color = Color.Red }
        );
        var s_border_spec = w.CreateStyle(
            border: new Border
            {
                Top = BorderStyleValues.Medium,
                Right = BorderStyleValues.Thin,
                Bottom = BorderStyleValues.Medium,
                Left = BorderStyleValues.Thin
            }
        );

        var ws_tydenna = w.AddWorksheet("TYDENNA", s);
        ws_tydenna.AddRow(s_mediumBorder_all);
        ws_tydenna.AddCell("TYZDENNA", s_font_red.CreateMergedStyle(s_border_spec));
        ws_tydenna.AddCell("Datum", s_border_spec);
        ws_tydenna.AddCell(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
        ws_tydenna.AddCell(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
        ws_tydenna.AddCell(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
        ws_tydenna.AddCell(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));

        var vals = new List<List<object>>
        {
            new List<object> {"", "Upstream kod" , "N_VNGSK2", "N_RWE2", "N_SPPRWE7", "P-CEZ14-2F" },
            new List<object> {"", "Downstream kod", "D_EEU11", "D_EEU11", "D_EEU11", "D_EEU11" },
            new List<object> {"", "Vstupny Bod", "DOMACI BOD", "ZASOBNIK_NAFTA", "ZASOBNIK_POZAGAS", "TAZOBNA SIET" },
            new List<object> {"", "Uzivatel", "jan.novak", "jan.novak", "jan.novak", "jan.novak" },
            new List<object> {"", "Verzia", 2, 2, 2, 2 },
            new List<object> {"", "Typ nominacie", "DENNA", "DENNA", "DENNA", "DENNA" }
        };

        foreach (var val in vals)
        {
            ws_tydenna.AddRow(s_font_blue);
            ws_tydenna.AddCell();
            ws_tydenna.AddCell(val[1], s_bolt);
            for (var i = 2; i < val.Count; i++)
            {
                ws_tydenna.AddCell(val[i], s_mediumBorder_rl);
            }
        }

        var s_thousandSpace = w.CreateStyle(
            numberFormat: new NumberingFormat("#,##0")
        );


        ws_tydenna.AddRow(s_fill_white.CreateMergedStyle(s_mediumBorder_all));
        ws_tydenna.AddCell("Množstvo", s_bolt.CreateMergedStyle(s_border_spec));
        ws_tydenna.AddCell("kWh", s_bolt.CreateMergedStyle(s_border_spec));
        ws_tydenna.AddCellWithFormula("Sum(C9:C15)", s_thousandSpace);
        ws_tydenna.AddCellWithFormula("Sum(D9:D15)", s_thousandSpace);
        ws_tydenna.AddCellWithFormula("Sum(E9:E15)", s_thousandSpace);
        ws_tydenna.AddCellWithFormula("Sum(F9:F15)", s_thousandSpace);

        timeFormat = "dd.MM.yyyy h:mm";
        var list = new List<int> { 100000, 200000, 300000, 400000 };
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);
        CreateRow2(ws_tydenna.AddRow(), DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture),
            DateTime.ParseExact("23.03.2015 6:00", timeFormat, CultureInfo.InvariantCulture), list);

        ws_tydenna.AddRow(s_fill_white);

        for (uint i = 1; i <= ws_tydenna.GetRow(1).Cells.Count; i++)
        {
            ws_tydenna.AddCell(s_mediumBorder_top);
            ws_tydenna.SetColumnWidth(i, 20);
        }

        w.Close();
        OpenXmlValidation.AssertValid(filePath);

        using var reopened = OpenExistingSpreadsheet(filePath);
        Assert.Equal(["INFO", "TYDENNA"], reopened.GetWorksheetsName());

        var info = reopened.GetWorksheet("INFO");
        Assert.NotNull(info);
        Assert.Equal("Datum", info.GetCellByReference("A1")?.GetStringValue());
        Assert.Equal("Verzia", info.GetCellByReference("A2")?.GetStringValue());
        Assert.Equal("E-Mail", info.GetCellByReference("A3")?.GetStringValue());
        Assert.Equal("jan.novak@example.com", info.GetCellByReference("B3")?.GetStringValue());
        Assert.Equal("Jan Novák", info.GetCellByReference("B4")?.GetStringValue());
        Assert.Equal("Typ Nominacie", info.GetCellByReference("A7")?.GetStringValue());

        var tydenna = reopened.GetWorksheet("TYDENNA");
        Assert.NotNull(tydenna);
        Assert.Equal("TYZDENNA", tydenna.GetCellByReference("A1")?.GetStringValue());
        Assert.Equal("Upstream kod", tydenna.GetCellByReference("B2")?.GetStringValue());
        Assert.Equal("Množstvo", tydenna.GetCellByReference("A8")?.GetStringValue());
        Assert.Equal("Sum(C9:C15)", tydenna.GetCellByReference("C8")?.GetFormula());
    }

    [Fact]
    public void LargeStyledSheet_IsValidAndReadable()
    {
        var filepath = GetFilepath("large-styled-sheet.xlsx");

        var headers = new List<string> { "p.č.", "Id města", "Hodnota 1", "Hodnota 2" };

        using var w = CreateNewSpreadsheet(filepath);
        var ws = w.AddWorksheet("MySheet - 1");

        var s = w.CreateStyle(new Font { FontSize = 20, Color = Color.Blue, FontName = FontNameValues.Tahoma });

        var c = ws.AddCellOnRange(3, 6, 2, s);
        c.SetValue("Testing data for my code");

        var s3 = w.CreateStyle(
            font: new Font { FontSize = 12, Color = Color.AliceBlue, FontName = FontNameValues.Calibri },
            fill: new Fill(Color.BlueViolet),
            numberFormat: new NumberingFormat("dd/mm/yyyy")
        );

        c = ws.AddCellOnIndex(3, 3, s3);
        c.SetValue(DateTime.UtcNow);

        ws.AddCell();
        var s4 = w.CreateStyle(
            new Font { Color = Color.Chartreuse },
            new Fill(Color.Black)
        );
        c = ws.AddCell("Alabama", s4);
        ws.AddCellOnRange(c.ColumnIndex, c.ColumnIndex, c.RowIndex, c.RowIndex + 1);

        var r = ws.AddRow(5, w.CreateStyle(
            new Font { FontSize = 13, Color = Color.AliceBlue, FontName = FontNameValues.Tahoma },
            new Fill(Color.DarkBlue))
        );

        for (var i = 0; i < headers.Count; i++)
        {
            var h = headers[i];
            c = r.AddCell();
            c.SetValue(h);
            ws.SetColumnWidth(Convert.ToUInt32(i + 1), 12);
        }

        ws.AddRow();
        var s1 = w.CreateStyle(
            new Font { FontSize = 12, Color = Color.Red, FontName = FontNameValues.Calibri },
            new Fill(Utils.ArgbHexConverter(Color.Aqua)),
            new Border(BorderStyleValues.Thin)
        );

        var s2 = w.CreateStyle(
            font: new Font { ArgbHexColor = Utils.ArgbHexConverter(Color.Blue) },
            numberFormat: new NumberingFormat("#,##0.00#"),
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Left }
        );


        for (var i = 1; i <= 1000; i++)
        {
            r = ws.AddRow(s1);


            var values = GetValue(i).ToList();
            for (var j = 0; j < values.Count; j++)
            {
                c = r.AddCell();
                switch (j)
                {
                    case 0:
                    case 1:
                        c.SetValue(Convert.ToInt32(values[j]));
                        break;
                    default:
                        c.SetValue(values[j]);

                        if (i % 2 == 0)
                        {
                            c.AddStyle(s2);
                        }

                        break;
                }
            }
        }

        w.Close();
        OpenXmlValidation.AssertValid(filepath);

        using var reopened = OpenExistingSpreadsheet(filepath);
        var sheet = reopened.GetWorksheet("MySheet - 1");

        Assert.NotNull(sheet);
        Assert.Equal("Testing data for my code", sheet.GetCellByReference("C2")?.GetStringValue());
        Assert.Equal("p.č.", sheet.GetCellByReference("A5")?.GetStringValue());
        Assert.Equal("Hodnota 2", sheet.GetCellByReference("D5")?.GetStringValue());
        // 1000 data rows starting after the header block.
        Assert.Equal(1, sheet.GetCellByReference("A7")?.GetIntValue());
        Assert.Equal(1000, sheet.GetCellByReference("A1006")?.GetIntValue());
    }

    [Fact]
    public void SheetWithTable_IsValidAndReadable()
    {
        var filepath = GetFilepath("sheet-with-table.xlsx");

        var headers = new List<string> { "p.č.", "Id města", "Hodnota 1", "Hodnota 2" };

        using var w = CreateNewSpreadsheet(filepath);
        var sheetName = "MySheet - 1";
        var ws = w.AddWorksheet(sheetName);
        ICell startCell, endCell;


        var s = w.CreateStyle(new Font { FontSize = 20, Color = Color.Blue, FontName = FontNameValues.Tahoma });

        var c = ws.AddCellOnRange(3, 6, 2, s);
        c.SetValue("Testing data for my code");

        var r = ws.AddRow(5, w.CreateStyle(
            new Font { FontSize = 13, Color = Color.AliceBlue, FontName = FontNameValues.Tahoma },
            new Fill(Color.DarkBlue))
        );

        for (var i = 0; i < headers.Count; i++)
        {
            var h = headers[i];
            c = r.AddCell();
            c.SetValue(h);
            ws.SetColumnWidth(Convert.ToUInt32(i + 1), 12);
        }

        startCell = r.Cells.First();

        //ws.AddRow();
        var s1 = w.CreateStyle(
            new Font { FontSize = 12, Color = Color.Red, FontName = FontNameValues.Calibri },
            new Fill(Utils.ArgbHexConverter(Color.Aqua)),
            new Border(BorderStyleValues.Thin)
        );

        var s2 = w.CreateStyle(
            font: new Font { ArgbHexColor = Utils.ArgbHexConverter(Color.Blue) },
            numberFormat: new NumberingFormat("#,##0.00#"),
            alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Left }
        );


        for (var i = 1; i <= 10; i++)
        {
            r = ws.AddRow(s1);


            var values = GetValue(i).ToList();
            for (var j = 0; j < values.Count; j++)
            {
                c = r.AddCell();
                switch (j)
                {
                    case 0:
                    case 1:
                        c.SetValue(Convert.ToInt32(values[j]));
                        break;
                    default:
                        c.SetValue(values[j]);

                        if (i % 2 == 0)
                        {
                            c.AddStyle(s2);
                        }

                        break;
                }
            }
        }

        endCell = r.Cells.Last();

        w.AddTable(sheetName, startCell, endCell, headers);

        w.Close();
        OpenXmlValidation.AssertValid(filepath);

        using var reopened = OpenExistingSpreadsheet(filepath);
        var table = reopened.GetTables(sheetName).SingleOrDefault();

        Assert.NotNull(table);
        Assert.Equal(sheetName, table.WorksheetName);
        Assert.Equal(headers, table.ColumnNames);
        Assert.Equal("p.č.", reopened.GetWorksheet(sheetName)?.GetCellByReference("A5")?.GetStringValue());
    }

    [Fact]
    public void ReopenAndAppendRows_StaysValid()
    {
        var filepath = GetFilepath("reopen-and-append.xlsx");
        using (var writer = CreateNewSpreadsheet(filepath))
        {
            var s = writer.CreateStyle(
                font: new Font { FontName = FontNameValues.Arial, FontSize = 20, Bold = true, Color = Color.DarkBlue, Underline = UnderlineValues.Double },
                alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            );
            var s1 = writer.CreateStyle(
                font: new Font { FontSize = 14, Bold = true, Color = Color.Coral },
                fill: new Fill(Color.MediumAquamarine),
                border: new Border { SetBorderStyle = BorderStyleValues.Thin, Bottom = BorderStyleValues.Medium }
            );
            var s2 = writer.CreateStyle(
                font: new Font { FontSize = 12, Bold = false, Color = Color.LightCoral },
                fill: new Fill(Color.Aquamarine),
                alignment: new Alignment { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center }
            );

            var sheet = writer.AddWorksheet("Mushrooms");

            var c = sheet.AddCellOnRange(2, 6, s);
            c.SetValue((object)"List of favorite mushrooms");

            var r = sheet.AddRow(s1);
            var headers = new[] { "ID", "Group ID", "Name", "Type", "Color", "Rate", "Place" };
            foreach (var h in headers)
            {
                r.AddCell(h);
                sheet.SetColumnWidth(12);
            }

            var r2 = sheet.AddRow(s2);
            for (var i = 0; i < r.Cells.Count; i++)
            {
                r2.AddCell();
            }
        }

        using (var writer = OpenExistingSpreadsheet(filepath))
        {
            var sheet = writer.GetWorksheet(writer.GetWorksheetsName().First());
            Assert.NotNull(sheet);

            var row = sheet.Rows.LastOrDefault();
            Assert.NotNull(row);

            for (var i = 0; i < 1000; i++)
            {
                var r = sheet.AddRow(row.Style); //style from existed rows cannot be loaded
                var data = CreateRow3();
                foreach (var cellData in data)
                {
                    var c = r.AddCell(cellData);
                    c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                }
            }
        }

        OpenXmlValidation.AssertValid(filepath);
    }

    [Fact]
    public void ExtendExcelAuthoredWorkbook_LargeAppend_StaysValid()
    {
        var filepath = GetFilepath("extend-excel-authored-large.xlsx");
        using var fileStream = File.Create(filepath, Resources.Example_1.Length);
        fileStream.Write(Resources.Example_1, 0, Resources.Example_1.Length);

        using var writer = OpenExistingSpreadsheet(fileStream);
        var sheet = writer.GetWorksheet(writer.GetWorksheetsName().First());
        Assert.NotNull(sheet);

        for (var i = 0; i < 100; i++)
        {
            ICell c;
            var r = sheet.AddRow();
            var data = CreateRow4();
            foreach (var cellData in data)
            {
                c = r.AddCell(cellData);
                c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
            }
            c = r.AddCellWithFormula($"Sum(B{i + 4}:F{i + 4})");
            c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
        }

        writer.Close();
        fileStream.Flush();
        // Example_1.xlsx itself carries pageSetup/@verticalDpi="0", which the schema forbids.
        OpenXmlValidation.AssertValid(fileStream, "verticalDpi");
    }

    [Fact]
    public void ExtendExcelAuthoredWorkbook_SmallAppend_StaysValid()
    {
        var filepath = GetFilepath("extend-excel-authored-small.xlsx");
        using var fileStream = File.Create(filepath, Resources.Example_1.Length);
        fileStream.Write(Resources.Example_1, 0, Resources.Example_1.Length);

        using var writer = OpenExistingSpreadsheet(fileStream);
        var sheet = writer.GetWorksheet(writer.GetWorksheetsName().First());
        Assert.NotNull(sheet);

        for (var i = 0; i < 10; i++)
        {
            var r = sheet.AddRow();
            var data = CreateRow4();
            ICell c;
            foreach (var cellData in data)
            {
                c = r.AddCell(cellData);
                c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
            }
            c = r.AddCellWithFormula($"Sum(B{i + 4}:F{i + 4})");
            c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
        }

        writer.Close();
        fileStream.Flush();
        // Example_1.xlsx itself carries pageSetup/@verticalDpi="0", which the schema forbids.
        OpenXmlValidation.AssertValid(fileStream, "verticalDpi");
    }

    [Fact]
    public void InMemoryWorkbook_RoundTripsThroughStream()
    {
        var memory = new MemoryStream();
        uint cellIndex;
        var textValue = "12300";
        using (var writer = CreateNewSpreadsheet(memory))
        {
            var sheet = writer.AddWorksheet();
            var cell = sheet.AddCell(textValue);
            cellIndex = cell.ColumnIndex;
        }

        OpenXmlValidation.AssertValid(memory);

        Assert.True(cellIndex >= 1);

        using (var writer = OpenExistingSpreadsheet(memory))
        {
            var sheet = writer.GetWorksheet(writer.GetWorksheetsName().First());
            var cell = sheet?.GetCell(cellIndex);
            Assert.NotNull(cell);
            Assert.Equal(textValue, cell.Value);
        }
    }

    private static void CreateRow(IWorksheet sheet, string header, IStyle headerStyle, object value, IStyle valueStyle)
    {
        sheet.AddRow();
        sheet.AddCell(header, headerStyle);
        sheet.AddCell(value, valueStyle);
    }

    private static void CreateRow2(IRow row, DateTime from, DateTime to, List<int> valList)
    {
        var w = row.Worksheet.Spreadsheet;
        var s_fill_green = w.CreateStyle(
            fill: new Fill(System.Drawing.ColorTranslator.FromHtml("#CCFFCC")),
            border: new Border { Right = BorderStyleValues.Medium, Left = BorderStyleValues.Medium }
        );
        var s_fill_grey = w.CreateStyle(
            fill: new Fill(Color.Gray),
            numberFormat: new NumberingFormat("dd.mm.yyyy h:mm"),
            border: new Border { Right = BorderStyleValues.Thin, Left = BorderStyleValues.Thin }
        );

        row.AddCell(from, s_fill_grey);
        row.AddCell(to, s_fill_grey);

        foreach (var val in valList)
        {
            row.AddCell(val, s_fill_green);
        }
    }

    private static IEnumerable<object> CreateRow3()
    {
        return new List<object> { Rnd.Next(1, 1000), Rnd.Next(1, 10000), "Name", "Type", "Color", Rnd.NextDouble() * 1000, "Place" };
    }

    private static IEnumerable<object> CreateRow4()
    {
        return new List<object> { Rnd.Next(1, 1000), Rnd.Next(1, 6), Rnd.Next(1, 6), Rnd.Next(1, 6), Rnd.Next(1, 6), Rnd.Next(1, 6) };
    }

    private static IEnumerable<double> GetValue(int pos)
    {
        return new[] { pos, Rnd.Next(1, 10000), Rnd.NextDouble() * 1000, Rnd.NextDouble() * 1000 };
    }
}