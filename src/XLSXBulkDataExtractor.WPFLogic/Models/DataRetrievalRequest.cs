using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XLSXBulkDataExtractor.Common;
using XLSXBulkDataExtractor.WPFLogic.ViewModels;

namespace XLSXBulkDataExtractor.WPFLogic.Models
{
    public class DataRetrievalRequest : NotifyPropertyChangedBase
    {
        private int _columnNumber = 1;
        public int ColumnNumber
        {
            get
            {
                return _columnNumber;
            }
        }
        private string _columnName = "A";
        /// <summary>
        /// Gets or sets the name of the column. Also sets the private field <seealso cref="_columnNumber"/> which is the integer equivalent of the column name.
        /// </summary>
        /// <value>The name of the column.</value>
        public string ColumnName
        {
            get
            {
                return _columnName;
            }
            set
            {
                if (_columnName != value)
                {
                    int columnNumber;
                    bool successfulParsing = false;

                    if (!int.TryParse(value, out columnNumber))
                    {
                        bool allLetters = value.All(char.IsLetter);

                        if (allLetters)
                        {
                            columnNumber = ExcelUtilities.GetColumnNumberFromColumnName(value);
                            successfulParsing = true;
                        }
                    }
                    else successfulParsing = true;

                    if (successfulParsing)
                    {
                        _columnName = value;
                        _columnNumber = columnNumber;
                    }
                    

                    OnPropertyChanged(nameof(ColumnName));
                }
            }
        }
        private int _row = 1;
        public int Row
        {
            get
            {
                return _row;
            }
            set
            {
                if (_row != value)
                {
                    _row = value;
                    OnPropertyChanged(nameof(Row));
                }
            }
        }
        private string _fieldName;
        public string FieldName
        {
            get
            {
                return _fieldName;
            }
            set
            {
                if (_fieldName != value)
                {
                    _fieldName = value;
                    OnPropertyChanged(nameof(FieldName));
                }
            }
        }
    }
}
