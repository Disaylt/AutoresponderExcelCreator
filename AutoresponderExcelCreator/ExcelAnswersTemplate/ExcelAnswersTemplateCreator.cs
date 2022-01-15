using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    public static class ExcelAnswersTemplateCreator
    {
        private static List<IExcelWorksheetTemplate>? _excelWorksheetsTemplate;
        private readonly static List<IExcelWorksheetTemplate>? _StandardExcelWorksheetsTemplate = new List<IExcelWorksheetTemplate>
        {

        };

        public static List<IExcelWorksheetTemplate> ExcelWorksheetsTemplate
        {
            get
            {
                if (_excelWorksheetsTemplate == null)
                {
                    _excelWorksheetsTemplate = new List<IExcelWorksheetTemplate>();
                }
                return _excelWorksheetsTemplate;
            }
        }
        public static void CreateExcelAnswersTemplate(string path)
        {
            IXLWorkbook workbook = new XLWorkbook();
            foreach(var worksheetName in ExcelWorksheetsTemplate)
            {
                worksheetName.AddNewWorksheetTamplate(workbook);
            }
            workbook.SaveAs($@"{path}\AnswersTemplate.xlsx");
        }
    }
}