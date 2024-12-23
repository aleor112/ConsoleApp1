using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestProject1;

namespace ConsoleApp1
{
    public class Helper
    {
        public static bool IsGroupFromAdesso(string group)
        {
            if (String.IsNullOrWhiteSpace(group)) return false;
            var bRes = Const.adessoAssignemntGroupKeyWords.Any(g => group.ToLower().Contains(g));
            return bRes;
        }

        public static string ClearRedundantPartOfString(string text)
        {
            var startIndex = text.IndexOf("\r\n") + 1;
            var endIndex = text.IndexOf("was");
            var length = endIndex == -1 ? text.Length - startIndex : endIndex - startIndex;
            var result = text.Substring(startIndex, length).Trim();
            return result;
        }
    }
}
