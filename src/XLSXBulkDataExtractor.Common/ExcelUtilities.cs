using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXBulkDataExtractor.Common
{
    public static class ExcelUtilities
    {
        /// <summary>
        /// Gets the number of the column from the name of the column in Excel column naming format, ie A, B, C, AA, BB, CC.
        /// </summary>
        /// <param name="columnName">Name of the column.</param>
        /// <returns>System.Int32.</returns>
        public static int GetColumnNumberFromColumnName(string columnName)
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
