using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXDataExtractor;

namespace XLSXBulkDataExtractor.WPFLogic
{
    class NewDataExtractor : DataExtractor, IDisposable
    {
        bool disposed;

        public NewDataExtractor(string path):
            base(path)
        {

        }


        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    //dispose managed resources
                }
            }
            //dispose unmanaged resources
            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
