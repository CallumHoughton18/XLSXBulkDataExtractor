using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Globals;

namespace XLSXBulkDataExtractor.WPFLogic.Interfaces
{
    public interface IUIControlsService
    {
        void DisplayAlert(string message, MessageType messageType);
    }
}
