using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SKToolsAddins.Utils
{
    public static class TextUtils
    {
        public static string RemoveJunkChar(this string str)//Remove special characters"(;,' in string
        {

            if (!string.IsNullOrEmpty(str))
            {
                str = str.Replace("(", "");
                str = str.Replace(")", "");
                str = str.Replace("'", "");
                str = str.Replace(";", "");
                str = str.Replace(" ", "");
            }
            return str;
        }
    }
}
