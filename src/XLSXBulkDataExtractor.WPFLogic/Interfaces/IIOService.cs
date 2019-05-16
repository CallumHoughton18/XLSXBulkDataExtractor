using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.Common;

namespace XLSXBulkDataExtractor.WPFLogic.Interfaces
{
    public interface IIOService
    {
        string ChooseFolderDialog();

        ReturnMessage SaveText(string text, string path);
    }
}
