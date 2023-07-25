using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeDocumentsApi.Excel.Enums;
using OfficeDocumentsApi.Excel.Interfaces;
using OfficeDocumentsApi.Excel.Styles;
using OfficeDocumentsApi.Excel.Test.Properties;
using Color = System.Drawing.Color;

namespace OfficeDocumentsApi.Excel.Test
{
    [TestClass]
    public class CreationTest : SpreadsheetTestBase
    {
        public static readonly Random Rnd = new Random();

        [TestMethod]
        public void BasicFile()
        {
            var filePath = GetFilepath("doc1.xlsx");
            using (var w = CreateTestee(filePath))
            {
                ;
            }
        }

        [TestMethod]
        public void CustomFile1()
        {
            var filePath = GetFilepath("doc2.xlsx");
            using (var w = CreateTestee(filePath))
            {
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

                ws_info.AddCellWithValue("Datum", s_bolt.CreateMergedStyle(s_mediumBorder_all));
                ws_info.AddCellWithValue(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture), s_mediumBorder_all);

                ws_info.AddStyle(s_font_blue);
                CreateRow(ws_info, "Verzia", s_bolt, 2, s_specBorder);
                CreateRow(ws_info, "E-Mail", s_bolt, "matej.zabsky@tollnet.cz", s_font_blue.CreateMergedStyle(s_specBorder).CreateMergedStyle(s_undeline));
                CreateRow(ws_info, "Meno", s_bolt, "Matìj Zábský", s_specBorder);
                CreateRow(ws_info, "Telefon", s_bolt, 608267556, s_specBorder);
                CreateRow(ws_info, "Uzivatel", s_bolt, "matej.zabsky", s_specBorder);
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
                ws_tydenna.AddCellWithValue("TYZDENNA", s_font_red.CreateMergedStyle(s_border_spec));
                ws_tydenna.AddCellWithValue("Datum", s_border_spec);
                ws_tydenna.AddCellWithValue(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
                ws_tydenna.AddCellWithValue(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
                ws_tydenna.AddCellWithValue(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));
                ws_tydenna.AddCellWithValue(DateTime.ParseExact("23.4.2014", timeFormat, CultureInfo.InvariantCulture));

                var vals = new List<List<object>>
                {
                    new List<object> {"", "Upstream kod" , "N_VNGSK2", "N_RWE2", "N_SPPRWE7", "P-CEZ14-2F" },
                    new List<object> {"", "Downstream kod", "D_EEU11", "D_EEU11", "D_EEU11", "D_EEU11" },
                    new List<object> {"", "Vstupny Bod", "DOMACI BOD", "ZASOBNIK_NAFTA", "ZASOBNIK_POZAGAS", "TAZOBNA SIET" },
                    new List<object> {"", "Uzivatel", "matej.zabsky", "matej.zabsky", "matej.zabsky", "matej.zabsky" },
                    new List<object> {"", "Verzia", 2, 2, 2, 2 },
                    new List<object> {"", "Typ nominacie", "DENNA", "DENNA", "DENNA", "DENNA" }
                };

                foreach (var val in vals)
                {
                    ws_tydenna.AddRow(s_font_blue);
                    ws_tydenna.AddCell();
                    ws_tydenna.AddCellWithValue(val[1], s_bolt);
                    for (var i = 2; i < val.Count; i++)
                    {
                        ws_tydenna.AddCellWithValue(val[i], s_mediumBorder_rl);
                    }
                }

                var s_thousandSpace = w.CreateStyle(
                    numberFormat: new NumberingFormat("#,##0")
                );


                ws_tydenna.AddRow(s_fill_white.CreateMergedStyle(s_mediumBorder_all));
                ws_tydenna.AddCellWithValue("Množstvo", s_bolt.CreateMergedStyle(s_border_spec));
                ws_tydenna.AddCellWithValue("kWh", s_bolt.CreateMergedStyle(s_border_spec));
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
            }
        }

        [TestMethod]
        public void CustomFile2()
        {
            var filepath = GetFilepath("doc3.xlsx");

            var headers = new List<string> { "p.è.", "Id místa", "Hodnota 1", "Hodnota 2" };

            using (var w = CreateTestee(filepath))
            {
                var ws = w.AddWorksheet("MySheet - 1");

                var s = w.CreateStyle(new Font { FontSize = 20, Color = Color.Blue, FontName = FontNameValues.Tahoma });

                var c = ws.AddCellOnRange(3, 6, 2, s);
                c.SetValue("Testing data for my code");

                var s3 = w.CreateStyle(
                    font: new Font { FontSize = 12, Color = Color.AliceBlue, FontName = FontNameValues.Calibri },
                    fill: new Fill(Color.BlueViolet),
                    numberFormat: new NumberingFormat("dd/mm/yyyy")
                );

                c = ws.AddCell(3, 3, s3);
                c.SetValue(DateTime.UtcNow);

                ws.AddCell();
                var s4 = w.CreateStyle(
                    new Font { Color = Color.Chartreuse },
                    new Fill(Color.Black)
                );
                c = ws.AddCellWithValue("Alabama", s4);
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
            }
        }

        [TestMethod]
        public void CustomFile3()
        {
            var filepath = GetFilepath("doc3.xlsx");

            var headers = new List<string> { "p.è.", "Id místa", "Hodnota 1", "Hodnota 2" };

            using (var w = CreateTestee(filepath))
            {
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
            }
        }

        [TestMethod]
        public void OpenAndAdjustCustomFile1()
        {
            var filepath = GetFilepath("doc4.xlsx");
            using (var writer = CreateTestee(filepath))
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
                    r.AddCellWithValue(h);
                    sheet.SetColumnWidth(12);
                }

                var r2 = sheet.AddRow(s2);
                for (var i = 0; i < r.Cells.Count; i++)
                {
                    r2.AddCell();
                }
            }

            using (var writer = CreateOpenTestee(filepath))
            {
                var sheet = writer.Worksheets.FirstOrDefault();
                Assert.IsNotNull(sheet);

                var row = sheet.Rows.LastOrDefault();
                Assert.IsNotNull(row);

                for (var i = 0; i < 1000; i++)
                {
                    var r = sheet.AddRow(row.Style); //style from existed rows cannot be loaded
                    var data = CreateRow3();
                    foreach (var cellData in data)
                    {
                        var c = r.AddCellWithValue(cellData);
                        c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                    }
                }
            }
        }

        [TestMethod]
        public void OpenAndAdjustCustomFile2()
        {
            var filepath = GetFilepath("doc5.xlsx");
            using (var fileStream = File.Create(filepath, Resources.Example_1.Length))
            {
                fileStream.Write(Resources.Example_1, 0, Resources.Example_1.Length);

                using (var writer = CreateOpenTestee(fileStream))
                {
                    var sheet = writer.Worksheets.FirstOrDefault();
                    Assert.IsNotNull(sheet);

                    for (var i = 0; i < 100; i++)
                    {
                        ICell c;
                        var r = sheet.AddRow();
                        var data = CreateRow4();
                        foreach (var cellData in data)
                        {
                            c = r.AddCellWithValue(cellData);
                            c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                        }
                        c = r.AddCellWithFormula($"Sum(B{i + 4}:F{i + 4})");
                        c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                    }
                }
            }
        }

        [TestMethod]
        public void OpenAndAdjustCustomFile3()
        {
            var filepath = GetFilepath("doc6.xlsx");
            using (var fileStream = File.Create(filepath, Resources.Example_1.Length))
            {
                fileStream.Write(Resources.Example_1, 0, Resources.Example_1.Length);

                using (var writer = CreateOpenTestee(fileStream))
                {
                    var sheet = writer.Worksheets.FirstOrDefault();
                    Assert.IsNotNull(sheet);

                    for (var i = 0; i < 10; i++)
                    {
                        var r = sheet.AddRow();
                        var data = CreateRow4();
                        ICell c;
                        foreach (var cellData in data)
                        {
                            c = r.AddCellWithValue(cellData);
                            c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                        }
                        c = r.AddCellWithFormula($"Sum(B{i + 4}:F{i + 4})");
                        c.AddStyle(sheet.GetCell(c.ColumnIndex, c.RowIndex - 1)?.Style);
                    }
                }
            }
        }

        [TestMethod]
        public void CreateInMemoryStream()
        {
            var memory = new MemoryStream();
            uint columnIndex = 0;
            var textValue = "12300";
            using (var writer = CreateTestee(memory))
            {
                var sheet = writer.AddWorksheet();
                var cell = sheet.AddCellWithValue(textValue);
                columnIndex = cell.ColumnIndex;
            }

            Assert.IsTrue(columnIndex > 0);

            using (var writer = CreateOpenTestee(memory))
            {
                var sheet = writer.Worksheets.FirstOrDefault();
                var cell = sheet.GetCell(columnIndex);
                Console.WriteLine(cell.CellReference);
                Assert.AreEqual(cell.Value, textValue);
            }
        }

        [TestMethod]
        public void FormulaInMemoryStream()
        {
            int num1 = 1;
            int num2 = 2;

            uint columnIndex = 0;

            var memory = new MemoryStream();
            using (var writer = CreateTestee(memory))
            {
                var sheet = writer.AddWorksheet();
                var cell1 = sheet.AddCellWithValue(num1);
                var cell2 = sheet.AddCellWithValue(num2);
                var sumCell = sheet.AddCellWithFormula("SUM(A1:B1)");

                columnIndex = sumCell.ColumnIndex;

            }

                Assert.IsTrue(columnIndex > 0);

            using (var writer = CreateOpenTestee(memory))
            {
                var sheet = writer.Worksheets.FirstOrDefault();
                var cell = sheet.GetCell(columnIndex);
                var formulaValue = cell.GetFormulaValue();
                Console.WriteLine(formulaValue);
                Assert.AreEqual(formulaValue, num1 + num2);
            }
        }

        private static void CreateRow(IWorksheet sheet, string header, IStyle headerStyle, object value, IStyle valueStyle)
        {
            sheet.AddRow();
            sheet.AddCellWithValue(header, headerStyle);
            sheet.AddCellWithValue(value, valueStyle);
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

            row.AddCellWithValue(from, s_fill_grey);
            row.AddCellWithValue(to, s_fill_grey);

            foreach (var val in valList)
            {
                row.AddCellWithValue(val, s_fill_green);
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
            //return new[] { pos.ToString(), Rnd.Next(1, 10000).ToString(), (Rnd.NextDouble() * 1000).ToString("N6"), (Rnd.NextDouble() * 1000).ToString("N6") };
        }
    }
}