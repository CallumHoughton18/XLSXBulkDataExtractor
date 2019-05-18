using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;
using XLSXBulkDataExtractor.WPFLogic.ViewModels;

namespace XLSXBulkDataExtractor
{
    /// <summary>
    /// Interaction logic for DataRetrievalWindow.xaml
    /// </summary>
    public partial class DataRetrievalWindow : Window
    {
        DataRetrievalViewModel vm;
        public DataRetrievalWindow(IIOService ioService, IXLIOService xlioService, IUIControlsService uiControlsService)
        {
            InitializeComponent();
            DataContext = vm = new DataRetrievalViewModel(ioService, xlioService, uiControlsService);
        }
    }
}
