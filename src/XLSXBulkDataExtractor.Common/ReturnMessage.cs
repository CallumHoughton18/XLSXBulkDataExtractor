using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXBulkDataExtractor.Common
{
    public class ReturnMessage
    {
        public bool Success { get; private set; }
        public string Message { get; private set; }

        public ReturnMessage(bool success, string message)
        {
            Success = success;
            Message = message;
        }
    }
}
