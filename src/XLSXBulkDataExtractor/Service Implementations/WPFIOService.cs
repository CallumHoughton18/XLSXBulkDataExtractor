using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.BL.Interfaces;

namespace XLSXBulkDataExtractor.Service_Implementations
{
    public class WPFIOService : IIOService
    {
        public string ChooseFolderDialog()
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                return dialog.FileName;
            }
            else
            {
                return string.Empty;
            }
        }
    }
}
