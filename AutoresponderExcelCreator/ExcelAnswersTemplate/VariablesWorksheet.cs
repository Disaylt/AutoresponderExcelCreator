using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    internal class VariablesWorksheet : ExcelWorksheetTemplate
    {
        private readonly List<(string key, string value)> _variablesAndValue = new List<(string, string)>
        {
            ("$buy_item$", "(Стандартный, неизменяемое значение) Вставляет название купленого товара"),
            ("$ref_item$", "(Стандартный, неизменяемое значение) Вставляет название рекомендуемого товара"),
            ("$ref_id$", "(Стандартный, неизменяемое значение) Вставляет название рекомендуемого товара"),
            ("$buyer_name$", "(Стандартный, неизменяемое значение) Вставляет имя покупателя, если оно доступно. (В некоторых случаях сайт отдает вместо имени - пользователь/user)"),
            ("$my_variable$", "Вставьте сюда необходимое слово или предложение"),
            ("$приветствие$", "Привет"),
            ("$приветствие$", "Здравствуй (Так как 2 одинаковых, то выберется одна из 2 переменных. Но количество переменных ограничено только возможностями экселя)")
        };

        private enum Titles
        {
            ColumnWithVariable = 1,
            ColumnWithValue
        }

        internal override string WorksheetName => "Variables";

        internal override void FillWorksheet(IXLWorksheet xLWorksheet)
        {
            xLWorksheet.Cell(1, (int)Titles.ColumnWithVariable).Value = "Переменная";
            xLWorksheet.Cell(1, (int)Titles.ColumnWithValue).Value = "Значение";
            int numRow = 2;
            foreach (var variable in _variablesAndValue)
            {
                xLWorksheet.Cell(numRow, (int)Titles.ColumnWithVariable).Value = variable.key;
                xLWorksheet.Cell(numRow, (int)Titles.ColumnWithValue).Value = variable.value;
                numRow += 1;
            }
        }
    }
}
