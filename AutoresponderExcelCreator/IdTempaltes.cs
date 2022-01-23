using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator
{
    internal class IdTempaltes
    {
        private IXLWorkbook xLWorkbook { get; set; }

        internal IdTempaltes(string pathToIdTemplate)
        {
            if(File.Exists(pathToIdTemplate))
            {
                xLWorkbook = new XLWorkbook(pathToIdTemplate);
            }
            else
            {
                xLWorkbook = new XLWorkbook();
            }
        }

        public string? GetIdTemplate(string? idProduct)
        {
            IXLWorksheet xLWorksheet = xLWorkbook.Worksheet(1);
            string? idTemplate = xLWorksheet?
                .RowsUsed()?
                .Skip(1)
                .Where(x => x.Cell(1).GetValue<string>() == idProduct)?
                .FirstOrDefault()?
                .Cell(2)
                .GetValue<string>();
            return idTemplate;
        }
    }
}   
