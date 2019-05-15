using System;
using System.Collections.Generic;
using System.Text;
using XLSXBulkDataExtractor.BL.ViewModels;

namespace XLSXBulkDataExtractor.BL.Models
{
    public class DataRetrievalRequest : NotifyPropertyChangedBase
    {
        private int _column;
        public int Column
        {
            get
            {
                return _column;
            }
            set
            {
                if (_column != value)
                {
                    _column = value;
                    OnPropertyChanged(nameof(Column));
                }
            }
        }
        private int _row;
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
