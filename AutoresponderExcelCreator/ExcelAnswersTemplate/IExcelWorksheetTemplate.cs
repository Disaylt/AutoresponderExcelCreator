using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    public interface IExcelWorksheetTemplate
    {
        public string WorksheetName { get; }
        public void AddNewWorksheetTamplate(IXLWorkbook xLWorkbook);
    }
}
