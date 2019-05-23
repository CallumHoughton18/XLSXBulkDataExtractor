using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using ToastNotifications;
using ToastNotifications.Lifetime;
using ToastNotifications.Position;

namespace XLSXBulkDataExtractor.Singletons
{
    public sealed class NotificationSingleton
    {
        private static readonly object padLock = new object();
        private static Notifier _notifierInstance = null;
        public static Notifier NotifierInstance
        {
            get
            {
                lock (padLock)
                {
                    if (_notifierInstance == null)
                    {
                        _notifierInstance = new Notifier(cfg =>
                        {
                            cfg.PositionProvider = new WindowPositionProvider(
                                parentWindow: Application.Current.MainWindow,
                                corner: Corner.BottomLeft,   
                                offsetX: 10,
                                offsetY: 10);

                            cfg.LifetimeSupervisor = new TimeAndCountBasedLifetimeSupervisor(
                                notificationLifetime: TimeSpan.FromSeconds(3),
                                maximumNotificationCount: MaximumNotificationCount.FromCount(10));

                            cfg.Dispatcher = Application.Current.Dispatcher;
                        });
                    }

                    return _notifierInstance;
                }
            }
        }
    }
}