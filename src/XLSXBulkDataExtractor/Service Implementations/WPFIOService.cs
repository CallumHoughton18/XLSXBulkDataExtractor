using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXBulkDataExtractor.Common;
using XLSXBulkDataExtractor.WPFLogic.Interfaces;

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

        public ReturnMessage SaveText(string text, string path)
        {
            try
            {
                File.WriteAllText(path, text);
                return new ReturnMessage(true, "");
            }
            catch (Exception e)
            {
                return new ReturnMessage(false, e.Message);
            }
        }
    }
}
