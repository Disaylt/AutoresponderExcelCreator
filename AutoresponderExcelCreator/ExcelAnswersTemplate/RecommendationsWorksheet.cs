using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    internal class RecommendationsWorksheet : ExcelWorksheetTemplate
    {
        private List<(string buyId, string buyItemName, string refId, string refItemName)> _valuesTemlate = new List<(string buyId, string buyItemName, string refId, string RefItemName)>
        {
            ("Купленый id","Наименование купленого товара","Рекомендуемый id","Рекомендуемое наименование"),
            ("В этой колнке необходимо записать купленый id","Здесь пишите как должен называться к ответе этот товар","(Может быть пустым) Если необходимо оставляете id товара, которые будете рекомендовать","В этой колонке пишите название товара, которй необходимо рекомендовать"),
            ("123","Наименование 1","234","товар 1"),
            ("123 (Будет выбрана 1 из 2 рекомендаций, можно добавить больше вариантов)","Наименование 2","345","товар 2"),
            ("234","Наименование 3","123","товар 1")
        };

        private enum Titles
        {
            BuyId = 1,
            BuyItemName,
            RefId,
            RefItemName
        }

        internal override string WorksheetName => "Recommendations";

        internal override void FillWorksheet(IXLWorksheet xLWorksheet)
        {
            int numRow = 1;
            foreach (var values in _valuesTemlate)
            {
                xLWorksheet.Cell(numRow, (int)Titles.BuyId).Value = values.buyId;
                xLWorksheet.Cell(numRow, (int)Titles.BuyItemName).Value = values.buyItemName;
                xLWorksheet.Cell(numRow, (int)Titles.RefId).Value = values.refId;
                xLWorksheet.Cell(numRow, (int)Titles.RefItemName).Value = values.refItemName;
                numRow += 1;
            }
        }
    }
}
