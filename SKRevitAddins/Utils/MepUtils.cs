using System.Collections.Generic;
using System.Linq;

namespace SKRevitAddins.Utils
{
    public static class MepUtils
    {
        public static int SelectPipeSize (this int fu, Dictionary<int, int> sizeDict)
        {
            int i = 0;
            var ele = sizeDict.ElementAt(i);
            var key = ele.Key;
            var value = ele.Value;
            while (fu > key)
            {
                i++;
                ele = sizeDict.ElementAt(i);
                key = ele.Key;
                value = ele.Value;
            }
            return value;
        }
    }
}
