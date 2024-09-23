namespace SKRevitAddins.Utils
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
