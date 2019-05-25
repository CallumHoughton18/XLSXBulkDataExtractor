using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using XLSXBulkDataExtractor.Service_Implementations;
using XLSXBulkDataExtractor.Singletons;
using XLSXBulkDataExtractor.WPFLogic.ViewModels;

namespace XLSXBulkDataExtractor
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            this.Dispatcher.UnhandledException += OnDispatcherUnhandledException;
        }

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            string errorMessage = string.Format($"An unhandled exception has occured: {e.Exception.Message}");
            MessageBox.Show(errorMessage, "Oops...", MessageBoxButton.OK, MessageBoxImage.Error);
            Current.Shutdown();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            var vm = new DataRetrievalViewModel(new IOService(), new XLIOService(), new UIControlsService(NotificationSingleton.NotifierInstance));
            MainWindow = new DataRetrievalWindow(vm);
            MainWindow.Show();
        }
    }
}
