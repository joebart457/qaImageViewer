using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace qaImageViewer.Service
{
    static class ValidatorService
    {

        public static bool ValidateSingleColumn(string text, int maxColumnAliasLength = 3)
        {
            if (text.Length > maxColumnAliasLength || text.Length == 0) return false;
            foreach (Char c in text)
            {
                if (!Char.IsLetter(c)) { return false; }
            }
            return true;
        }

        public static bool ValidateSingleColumnOrRowIdOption(string text, int maxColumnAliasLength = 3)
        {
            if (text == ExcelAppHelperService.ROWID_OPTION) return true;
            if (text.Length > maxColumnAliasLength || text.Length == 0) return false;
            foreach (Char c in text)
            {
                if (!Char.IsLetter(c)) { return false; }
            }
            return true;
        }

        public static bool ValidateColumnFormat(string str)
        {
            string[] columnRanges = str.Split(',');
            foreach (string colRange in columnRanges)
            {
                string[] cols = colRange.Split(':');
                if (cols.Length == 2)
                {
                    foreach (string column in cols)
                    {
                        foreach (char c in column)
                        {
                            if (Char.IsLetter(c))
                            {
                                continue;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
                else
                {
                    return false;
                }
            }
            return true;
        }
    }
}
