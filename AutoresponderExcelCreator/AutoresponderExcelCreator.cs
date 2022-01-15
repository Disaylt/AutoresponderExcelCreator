namespace AutoresponderExcelCreator
{
    public class AutoresponderExcelCreator
    {
        private readonly string _excelFolserPath;
        private string _fileName;
        public AutoresponderExcelCreator(string excelFolserPath)
        {
            _excelFolserPath = excelFolserPath;
        }

        public string GetResponseText(string feedbackText, string username = "")
        {
            return string.Empty;
        }
    }
}