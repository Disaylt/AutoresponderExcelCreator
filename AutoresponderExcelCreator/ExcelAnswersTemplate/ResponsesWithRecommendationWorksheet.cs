﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    internal class ResponsesWithRecommendationWorksheet : IExcelWorksheetTemplate
    {
        public string WorksheetName => "ResponsesWithRecommendation";

        private void FillWorksheet(IXLWorksheet xLWorksheet)
        {
            for (int numColumn = 1; numColumn <= 4; numColumn++)
            {
                xLWorksheet.Column(numColumn).Style.Alignment.WrapText = true;
                xLWorksheet.Cell(1, numColumn).Value = $"Название заголовка {numColumn}";
                for (int numRow = 2; numRow <= 4; numRow++)
                {
                    xLWorksheet.Cell(numRow, numColumn).Value = $"Текст {numRow - 1} (Необходимо удалить/заменить все заполненые стандартным текстом ячейки)";
                }
            }
            xLWorksheet.Cell(1, 5).Value = $"Продолжите заголовки либо удалите лишние(В том числе этот заголовок)";
            xLWorksheet.Cell(2, 5).Value = $"В тексте можете использовать стандартные ключи для рекомендаций. Посмотреть их можно во вкладке Variables";
        }

        public void AddNewWorksheetTamplate(IXLWorkbook xLWorkbook)
        {
            try
            {
                xLWorkbook.AddWorksheet(WorksheetName);
                IXLWorksheet xLWorksheet = xLWorkbook.Worksheet(WorksheetName);
                FillWorksheet(xLWorksheet);
            }
            catch
            {

            }
        }
    }
}
