using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Models;

namespace XLSXBulkDataExtractor.WPFLogicTests
{
    public static class WPFLogicTestsHelper
    {
        public static IEnumerable<DataRetrievalRequest> GenerateMockRequests()
        {
            for (int i = 0; i < 3; i++)
            {
                yield return new DataRetrievalRequest { ColumnName = i.ToString(), Row = i, FieldName = $"MockFieldName{i}" };
            }
        }
    }
}
