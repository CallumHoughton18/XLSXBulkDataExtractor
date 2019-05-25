using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXBulkDataExtractor.Common
{
    public static class ExcelUtilities
    {
        public static int GetColumnNumbeFromColumnName(string columnName)
        {
            string columnNameUpper = columnName.ToUpperInvariant();
            int number = 0;
            int pow = 1;
            for (int i = columnNameUpper.Length - 1; i >= 0; i--)
            {
                number += (columnNameUpper[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            return number;
        }
    }
}
