namespace AutoresponderExcelCreator
{
    public class AutoresponderExcel
    {
        private readonly BlackListHandler _blackListHandler;
        private readonly IdTempaltes _idTempaltes;
        private readonly string _directoryBrandExcelTemplates;
        private readonly string _pathToStandardExcelTemplate;
        private string _lastUseBrand;
        private Random _random;
        private IXLWorkbook? workbook { get; set; }
        private Dictionary<string, List<string>> variablesKeyAndValues { get; set; }
        public List<string> ExceptionMessages { get; }
        

        public AutoresponderExcel(string pathToStandardExcelTemplate, string directoryBrandExcelTemplates, string pathToBlackList, string pathToIdTemplate)
        {
            _idTempaltes = new IdTempaltes(pathToIdTemplate);
            _blackListHandler = new BlackListHandler(pathToBlackList);
            _lastUseBrand = string.Empty;
            _directoryBrandExcelTemplates = directoryBrandExcelTemplates;
            _pathToStandardExcelTemplate = pathToStandardExcelTemplate;
            _random = new Random();
            variablesKeyAndValues = new Dictionary<string, List<string>>();
            ExceptionMessages = new List<string>();
        }

        private void UpdateVariables()
        {
            string sheetName = "Variables";
            variablesKeyAndValues = new Dictionary<string, List<string>>();
            if (workbook.TryGetWorksheet(sheetName, out var sheet))
            {
                var rows = sheet.RowsUsed().Skip(1);
                foreach (var row in rows)
                {
                    string key = row.Cell(1).Value?.ToString()?.Trim() ?? string.Empty;
                    if(!key.StartsWith('$') || !key.EndsWith('$')) { key = string.Empty; }
                    string value = row.Cell(2).Value?.ToString()?.Trim() ?? string.Empty;
                    if(!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(key))
                    {
                        if(variablesKeyAndValues.ContainsKey(key))
                        {
                            variablesKeyAndValues[key].Add(value);
                        }
                        else
                        {
                            List<string> values = new List<string> { value };
                            variablesKeyAndValues.Add(key, values);
                        }
                    }
                }
            }
        }

        private void UpdateExcel(string? brand, string? productId)
        {
            string? templateName;
            string? idTemplate = _idTempaltes.GetIdTemplate(productId);
            if (!string.IsNullOrEmpty(idTemplate))
            {
                templateName = idTemplate;
            }
            else
            {
                templateName = brand;
            }

            if (workbook == null || _lastUseBrand != templateName)
            {
                string[] availableBrands = Directory.GetFiles(_directoryBrandExcelTemplates, "*.xlsx")
                    .Select(x => Path.GetFileName(x))
                    .ToArray();

                if (availableBrands.Contains($"{templateName}.xlsx"))
                {
                    workbook = new XLWorkbook($@"{_directoryBrandExcelTemplates}\{templateName}.xlsx");
                }
                else
                {
                    workbook = new XLWorkbook(_pathToStandardExcelTemplate);
                }
                UpdateVariables();
                _lastUseBrand = templateName ?? string.Empty;
            }
        }

        private RecommendationProductInfo? GetGetRecommendationUserInfo(string? userName, string? producId = "")
        {
            string sheetName = "UserRecommendations";
            if(!string.IsNullOrEmpty(userName) && workbook.TryGetWorksheet(sheetName, out var sheet))
            {
                List<RecommendationProductInfo> productsInfo = sheet
                    .RowsUsed()
                    .Where(x => x.Cell(1).GetString().ToLower() == userName.ToLower() && (x.Cell(2).GetString() == string.Empty || x.Cell(2).GetString() == producId))
                    .Select(x => new RecommendationProductInfo
                    {
                        BuyProductName = x.Cell(3).GetString(),
                        RecommendationId = x.Cell(4).GetString(),
                        RecommendationName = x.Cell(5).GetString()
                    })
                    .ToList();
                if(productsInfo.Count != 0)
                {
                    var productInfo = productsInfo.ElementAt(_random.Next(0, productsInfo.Count));
                    return productInfo;
                }
            }

            return null;
        }

        private RecommendationProductInfo? GetRecommendationInfo(string? productId)
        {
            string sheetName = "Recommendations";
            if (workbook.TryGetWorksheet(sheetName, out var sheet))
            {
                RecommendationProductInfo[] productsInfo = sheet
                    .Column(1)
                    .CellsUsed()
                    .Where(x => x.Value?.ToString()?.Trim() == productId)
                    .Select(x => new RecommendationProductInfo
                    {
                        BuyProductName = sheet.Cell(x.Address.RowNumber, 2).Value.ToString() ?? string.Empty,
                        RecommendationId = sheet.Cell(x.Address.RowNumber, 3).Value.ToString() ?? string.Empty,
                        RecommendationName = sheet.Cell(x.Address.RowNumber, 4).Value.ToString() ?? string.Empty
                    })
                    .Where(x => !string.IsNullOrEmpty(x.BuyProductName))
                    .ToArray();

                if(productsInfo.Count() != 0)
                {
                    var productInfo = productsInfo.ElementAt(_random.Next(0, productsInfo.Count()));
                    return productInfo;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        private string GetPieceText(IXLColumn xLColumn)
        {
            string pieceText = string.Empty;
            int numRows = xLColumn.CellsUsed().Count();
            if (numRows > 1)
            {
                pieceText = xLColumn?
                    .CellsUsed()?
                    .Skip(1)?
                    .ElementAtOrDefault(_random.Next(0, numRows - 1))?
                    .Value?
                    .ToString() ?? string.Empty;
            }
            return pieceText;
        }

        private string GetAnswerText(string sheetName)
        {
            string answerText = string.Empty;
            if(workbook.TryGetWorksheet(sheetName, out var sheet))
            {
                foreach(var cell in sheet.Row(1).CellsUsed())
                {
                    var column = sheet.Column(cell.Address.ColumnNumber);
                    string pieceText = GetPieceText(column);
                    if(!string.IsNullOrEmpty(pieceText))
                    {
                        answerText += $"{pieceText} ";
                    }
                }
            }
            return answerText;
        }

        private string ReplaceRecommendationProductInfo(string answerText, RecommendationProductInfo? recommendationProductInfo)
        {
            answerText = answerText.Replace("$buy_item$", recommendationProductInfo?.BuyProductName ?? string.Empty);
            answerText = answerText.Replace("$ref_item$", recommendationProductInfo?.RecommendationName ?? string.Empty);
            answerText = answerText.Replace("$ref_id$", recommendationProductInfo?.RecommendationId ?? string.Empty);
            return answerText;
        }

        private string ReplaceUserName(string answerText, string? userName)
        {
            answerText = answerText.Replace("$buyer_name$", userName);
            return answerText;
        }

        private string ReplaceCustomVariables(string answerText)
        {
            foreach(var variables in variablesKeyAndValues)
            {
                string value = variables.Value[_random.Next(0, variables.Value.Count)];
                answerText = answerText.Replace(variables.Key, value);
            }
            return answerText;
        }

        private bool CheckTextForVariables(string text)
        {
            if(text.Contains('$'))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string? GetResponseText(string? feedbackText, string? brand, string? productId, string? username = "")
        {
            try
            {
                string answerText;
                if (_blackListHandler.CheckBanWords(feedbackText)) { return null; }
                UpdateExcel(brand, productId);
                RecommendationProductInfo? recommendationProductInfo = GetRecommendationInfo(productId);
                RecommendationProductInfo? userRecommendationProductInfo = GetGetRecommendationUserInfo(username, productId);

                if(userRecommendationProductInfo != null)
                {
                    answerText = GetAnswerText("ResponsesWithUserRecommendation");
                    answerText = ReplaceRecommendationProductInfo(answerText, userRecommendationProductInfo);
                }
                else if (recommendationProductInfo != null)
                { 
                    answerText = GetAnswerText("ResponsesWithRecommendation");
                    answerText = ReplaceRecommendationProductInfo(answerText, recommendationProductInfo);
                }
                else
                {
                    answerText = GetAnswerText("Responses");
                }
                answerText = ReplaceUserName(answerText, username);
                answerText = ReplaceCustomVariables(answerText);

                if (CheckTextForVariables(answerText)) { return null; }

                return answerText;
            }
            catch (Exception ex)
            {
                ExceptionMessages.Add($"FeedbackText: {feedbackText}; Id: {productId}; Error: {ex.Message}");
                return null;
            }
        }
    }
}