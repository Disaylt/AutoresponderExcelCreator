using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator.ExcelAnswersTemplate
{
    public static class ExcelAnswersTemplateCreator
    {
        private static List<ExcelWorksheetTemplate>? _excelWorksheetsTemplate;
        private readonly static List<ExcelWorksheetTemplate> _standardExcelWorksheetsTemplate = new List<ExcelWorksheetTemplate>
        {
            new RecommendationsWorksheet(),
            new ResponsesWithRecommendationWorksheet(),
            new ResponsesWorksheet(),
            new VariablesWorksheet()
        };
        /// <summary>
        /// Contains custom worksheets 
        /// </summary>
        public static List<ExcelWorksheetTemplate> ExcelWorksheetsTemplate
        {
            get
            {
                if (_excelWorksheetsTemplate == null)
                {
                    _excelWorksheetsTemplate = new List<ExcelWorksheetTemplate>();
                }
                return _excelWorksheetsTemplate;
            }
        }

        /// <summary>
        /// Shows last error in class 
        /// </summary>
        public static string? ExcelLastError { get; set; }

        /// <summary>
        /// Creates a template at the specified path with the name AnswersTemplate
        /// </summary>
        /// <param name="path">Path to create a template</param>
        public static void Create(string path)
        {
            try
            {
                IXLWorkbook xLWorkbook = new XLWorkbook();
                List<ExcelWorksheetTemplate> combainWorksheetsTempate = _standardExcelWorksheetsTemplate
                    .Concat(ExcelWorksheetsTemplate)
                    .ToList();

                foreach (var worksheetName in combainWorksheetsTempate)
                {
                    worksheetName.AddNewWorksheet(xLWorkbook);
                }
                xLWorkbook.SaveAs($@"{path}\AnswersTemplate.xlsx");
            }
            catch (Exception ex)
            {
                ExcelLastError = ex.Message;
            }
        }
    }
}