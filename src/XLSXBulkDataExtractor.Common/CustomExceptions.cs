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
        public NoDataOutputtedException(string message, string exceptionTitle) :
            base(message)
        {
            ExceptionTitle = exceptionTitle;
        }
    }
    
    public class ExtractionFailedException : Exception
    {
        public string ExceptionTitle { get; private set; }
        public IEnumerable<string> FailedExtractionMessages { get; private set; }
        public ExtractionFailedException(string message,string exceptionTitle, IEnumerable<string> failedExtractionMessages) :
            base(message)
        {
            FailedExtractionMessages = failedExtractionMessages;
            ExceptionTitle = exceptionTitle;
        }
    }

}
