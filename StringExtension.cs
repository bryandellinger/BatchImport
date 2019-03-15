using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ADBatchImportForInterProd
{
    public static class StringExtension
    {
        public static string Truncate(this string s, int length)
        {
            if (s.Length > length)
            {
                s = s.Substring(0, length);
            }

            return s;
        }
    }
}
