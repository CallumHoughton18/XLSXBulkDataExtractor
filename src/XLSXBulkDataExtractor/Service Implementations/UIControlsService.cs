using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ToastNotifications;
using ToastNotifications.Lifetime;
using ToastNotifications.Position;
using XLSXBulkDataExtractor.WPFLogic.Globals;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;
using ToastNotifications.Messages;

namespace XLSXBulkDataExtractor.Service_Implementations
{
    class UIControlsService : IUIControlsService
    {
        protected Notifier _notifier;

        public UIControlsService(Notifier notifier)
        {
            _notifier = notifier;
        }

        public void DisplayAlert(string message, MessageType messageType)
        {
            switch (messageType)
            {
                case MessageType.Information:
                    _notifier.ShowInformation(message);
                    break;
                case MessageType.Error:
                    _notifier.ShowError(message);
                    break;
                case MessageType.Success:
                    _notifier.ShowSuccess(message);
                    break;
                default:
                    break;
            }
        }
    }
}
