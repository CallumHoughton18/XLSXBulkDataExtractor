using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXBulkDataExtractor.BL.Models
{
    public class DataRetrievalRequest
    {
        public int Column { get; private set; }
        public int Row { get; private set; }
        public string FieldName { get; private set; }

        public DataRetrievalRequest(int column, int row, string fieldName)
        {
            Column = column;
            Row = row;
            FieldName = fieldName;
        }
    }
}
