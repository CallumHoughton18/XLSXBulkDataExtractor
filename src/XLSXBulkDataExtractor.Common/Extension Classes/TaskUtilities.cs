using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXBulkDataExtractor.Common.Extension_Classes
{
    public static class TaskUtilities
    {
        /// <summary>
        /// fires the task unsafely and as an asynchronous operation. Any exceptions are rethrown.
        /// </summary>
        /// <param name="task">The task.</param>
        public static async void FireUnsafeAsync(this Task task)
        {
            try
            {
                await task;
            }
            catch (Exception)
            {
                throw;
            }
        }
    }
}
