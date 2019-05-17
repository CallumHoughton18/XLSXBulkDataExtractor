using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.Common;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;

namespace XLSXBulkDataExtractor.Service_Implementations
{
    class XLIOService : IXLIOService
    {
        public ReturnMessage SaveWorkbook(string path, XLWorkbook xlWorkbook)
        {
            try
            {
                xlWorkbook.SaveAs(path);
                return new ReturnMessage(true, $"Successfully saved to: {path}");
            }
            catch(Exception e)
            {
                return new ReturnMessage(false, e.Message);
            }
        }
    }
}
