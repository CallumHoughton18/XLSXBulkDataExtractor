using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.Common;

namespace XLSXBulkDataExtractor.WPFLogic.Interfaces
{
    public interface IXLIOService
    {
        ReturnMessage SaveWorkbook(string path, XLWorkbook workbook);
    }
}
