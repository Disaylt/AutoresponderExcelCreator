using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    public abstract class ExcelWorksheetTemplate
    {
        internal abstract string WorksheetName { get; }

        internal abstract void FillWorksheet(IXLWorksheet xLWorksheet);

        internal virtual void UpdateSettingsWorksheet(IXLWorksheet xLWorksheet)
        {
            xLWorksheet.Style.Alignment.WrapText = true;
            foreach(var column in xLWorksheet.ColumnsUsed())
            {
                column.Width = 50;
                column.DataType = XLDataType.Text;
            }
        }

        public void AddNewWorksheet(IXLWorkbook xLWorkbook)
        {
            xLWorkbook.AddWorksheet(WorksheetName);
            IXLWorksheet xLWorksheet = xLWorkbook.Worksheet(WorksheetName);
            FillWorksheet(xLWorksheet);
            UpdateSettingsWorksheet(xLWorksheet);
        }
    }
}
