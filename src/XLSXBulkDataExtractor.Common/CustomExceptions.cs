using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXBulkDataExtractor.Common
{
    public class CollectionEmptyException : Exception { }
    public class NoDataOutputtedException : Exception
    {
        public string ExceptionTitle { get; private set; }
        public NoDataOutputtedException(string message, string exceptionTitle):
            base(message)
        {
            ExceptionTitle = exceptionTitle;
        }
    }

}
