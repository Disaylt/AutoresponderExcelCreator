using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoresponderExcelCreator
{
    internal class BlackListHandler
    {
        private readonly string _pathToBlackList;
        private List<string>? _banWords;
        private List<string> banWords 
        { 
            get
            {
                if(_banWords == null)
                {
                    _banWords = File
                        .ReadAllLines(_pathToBlackList, Encoding.UTF8)
                        .Select(x => x.ToLower().Trim())
                        .ToList();
                }
                return _banWords;
            } 
        }

        internal BlackListHandler(string path)
        {
            _pathToBlackList = path;
        }

        internal bool CheckBanWords(string? text)
        {
            bool isBanWords = false;
            if (!string.IsNullOrEmpty(text))
            {
                string lowerText = text.ToLower();
                foreach (string banWord in banWords)
                {
                    if (lowerText.Contains(banWord))
                    {
                        isBanWords = true;
                        break;
                    }
                }
            }
            return isBanWords;
        }
    }
}
