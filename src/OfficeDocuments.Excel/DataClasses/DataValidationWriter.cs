using DocumentFormat.OpenXml;
using OfficeDocuments.Excel.Enums;
using OfficeDocuments.Excel.Options;
using SpreadsheetLib = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeDocuments.Excel.DataClasses;

/// <summary>
/// Writes data-validation rules onto a worksheet, keeping the shared
/// <c>dataValidations</c> element in the correct CT_Worksheet position.
/// </summary>
internal sealed class DataValidationWriter(SpreadsheetLib.Worksheet worksheetElement, WorksheetElementOrderer orderer)
{
    public void Add(string reference, DataValidationOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);

        var validations = worksheetElement.GetFirstChild<SpreadsheetLib.DataValidations>();
        if (validations == null)
        {
            validations = new SpreadsheetLib.DataValidations();
            orderer.InsertDataValidations(validations);
        }

        var validation = new SpreadsheetLib.DataValidation
        {
            AllowBlank = options.AllowBlank,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = reference }
        };

        validation.Type = options.Type switch
        {
            DataValidationType.List => SpreadsheetLib.DataValidationValues.List,
            DataValidationType.Whole => SpreadsheetLib.DataValidationValues.Whole,
            DataValidationType.Decimal => SpreadsheetLib.DataValidationValues.Decimal,
            DataValidationType.Date => SpreadsheetLib.DataValidationValues.Date,
            DataValidationType.Custom => SpreadsheetLib.DataValidationValues.Custom,
            _ => throw new ArgumentOutOfRangeException(nameof(options))
        };

        if (options.Operator.HasValue)
        {
            validation.Operator = options.Operator.Value switch
            {
                DataValidationOperator.Between => SpreadsheetLib.DataValidationOperatorValues.Between,
                DataValidationOperator.NotBetween => SpreadsheetLib.DataValidationOperatorValues.NotBetween,
                DataValidationOperator.Equal => SpreadsheetLib.DataValidationOperatorValues.Equal,
                DataValidationOperator.NotEqual => SpreadsheetLib.DataValidationOperatorValues.NotEqual,
                DataValidationOperator.GreaterThan => SpreadsheetLib.DataValidationOperatorValues.GreaterThan,
                DataValidationOperator.LessThan => SpreadsheetLib.DataValidationOperatorValues.LessThan,
                DataValidationOperator.GreaterThanOrEqual => SpreadsheetLib.DataValidationOperatorValues.GreaterThanOrEqual,
                DataValidationOperator.LessThanOrEqual => SpreadsheetLib.DataValidationOperatorValues.LessThanOrEqual,
                _ => throw new ArgumentOutOfRangeException(nameof(options))
            };
        }

        if (!string.IsNullOrWhiteSpace(options.PromptTitle))
        {
            validation.PromptTitle = options.PromptTitle;
        }

        if (!string.IsNullOrWhiteSpace(options.Prompt))
        {
            validation.Prompt = options.Prompt;
        }

        if (!string.IsNullOrWhiteSpace(options.ErrorTitle))
        {
            validation.ErrorTitle = options.ErrorTitle;
        }

        if (!string.IsNullOrWhiteSpace(options.Error))
        {
            validation.Error = options.Error;
        }

        validation.Append(new SpreadsheetLib.Formula1(options.Formula1));
        if (!string.IsNullOrWhiteSpace(options.Formula2))
        {
            validation.Append(new SpreadsheetLib.Formula2(options.Formula2));
        }

        validations.Append(validation);
        validations.Count = Convert.ToUInt32(validations.Count());
    }
}
