using DocumentFormat.OpenXml;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Interfaces;
using OfficeDocuments.Excel.Options;
using Color = System.Drawing.Color;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Writes conditional-formatting rules onto a worksheet. Differential formats are resolved
/// through the injected delegate so this writer stays decoupled from the workbook stylesheet.
/// </summary>
internal sealed class ConditionalFormattingWriter(
    SpreadsheetLib.Worksheet worksheetElement,
    WorksheetElementOrderer orderer,
    Func<IStyle, uint> getOrCreateDifferentialFormat)
{
    public void Add(string reference, ConditionalFormattingOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var conditionalFormatting = new SpreadsheetLib.ConditionalFormatting
        {
            SequenceOfReferences = new ListValue<StringValue> { InnerText = reference }
        };

        var rule = new SpreadsheetLib.ConditionalFormattingRule
        {
            Priority = Convert.ToInt32(GetNextPriority())
        };

        switch (options.Type)
        {
            case ConditionalFormattingType.GreaterThan:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.CellIs;
                rule.Operator = SpreadsheetLib.ConditionalFormattingOperatorValues.GreaterThan;
                rule.FormatId = getOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula(options.Formula ?? throw new InvalidOperationException("Conditional formatting formula is required.")));
                break;
            case ConditionalFormattingType.LessThan:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.CellIs;
                rule.Operator = SpreadsheetLib.ConditionalFormattingOperatorValues.LessThan;
                rule.FormatId = getOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula(options.Formula ?? throw new InvalidOperationException("Conditional formatting formula is required.")));
                break;
            case ConditionalFormattingType.ContainsText:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.ContainsText;
                rule.Text = options.Text;
                rule.FormatId = getOrCreateDifferentialFormat(options.Style!);
                rule.Append(new SpreadsheetLib.Formula($"NOT(ISERROR(SEARCH(\"{EscapeFormulaString(options.Text)}\",{reference.Split(':')[0]})))"));
                break;
            case ConditionalFormattingType.DuplicateValues:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.DuplicateValues;
                rule.FormatId = getOrCreateDifferentialFormat(options.Style!);
                break;
            case ConditionalFormattingType.TwoColorScale:
                rule.Type = SpreadsheetLib.ConditionalFormatValues.ColorScale;
                rule.Append(
                    new SpreadsheetLib.ColorScale(
                        new SpreadsheetLib.ConditionalFormatValueObject { Type = SpreadsheetLib.ConditionalFormatValueObjectValues.Min },
                        new SpreadsheetLib.ConditionalFormatValueObject { Type = SpreadsheetLib.ConditionalFormatValueObjectValues.Max },
                        new SpreadsheetLib.Color { Rgb = Utils.ArgbHexConverter(options.MinimumColor ?? Color.LightGreen) },
                        new SpreadsheetLib.Color { Rgb = Utils.ArgbHexConverter(options.MaximumColor ?? Color.DarkGreen) }
                    )
                );
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(options));
        }

        conditionalFormatting.Append(rule);
        orderer.InsertConditionalFormatting(conditionalFormatting);
    }

    private uint GetNextPriority()
    {
        var priorities = worksheetElement.Elements<SpreadsheetLib.ConditionalFormatting>()
            .SelectMany(item => item.Elements<SpreadsheetLib.ConditionalFormattingRule>())
            .Select(item => item.Priority?.Value ?? 0);

        return Convert.ToUInt32(priorities.DefaultIfEmpty().Max() + 1);
    }

    private static string EscapeFormulaString(string? value)
    {
        return (value ?? string.Empty).Replace("\"", "\"\"");
    }
}
