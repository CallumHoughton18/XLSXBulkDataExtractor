using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.WPFLogic.Globals;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;

namespace XLSXBulkDataExtractor.Service_Implementations
{
    class UIControlsService : IUIControlsService
    {
        public void DisplayAlert(string message, string caption, MessageType messageType)
        {
            switch (messageType)
            {
                case MessageType.Information:
                    System.Windows.MessageBox.Show(message, caption, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                    break;
                case MessageType.Error:
                    System.Windows.MessageBox.Show(message, caption, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    break;
                case MessageType.None:
                    System.Windows.MessageBox.Show(message, caption, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.None);
                    break;
                default:
                    break;
            }
        }
    }
}
