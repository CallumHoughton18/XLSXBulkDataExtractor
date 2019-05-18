using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using XLSXBulkDataExtractor.Service_Implementations;

namespace XLSXBulkDataExtractor
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            MainWindow = new DataRetrievalWindow(new IOService(), new XLIOService(), new UIControlsService());
            MainWindow.Show();
        }
    }
}
