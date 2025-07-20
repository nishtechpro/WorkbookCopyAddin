using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorkbookCopyAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            try
            {
                string destFolder = @"\\server\Users\Public\Documents\xl\";
                string ext = Path.GetExtension(Wb.Name).ToLower();
                string[] addinExts = { ".xlam", ".xla", ".xll" };

                // Skip Add-in files
                if (Array.Exists(addinExts, element => element == ext))
                    return;

                string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (!string.IsNullOrEmpty(Wb.Path))
                {
                    // Saved workbook
                    string baseName = Path.GetFileNameWithoutExtension(Wb.Name);
                    string fileName = $"{baseName}_{timeStamp}{ext}";
                    string destPath = Path.Combine(destFolder, fileName);
                    Wb.SaveCopyAs(destPath);
                }
                else
                {
                    // Unsaved workbook
                    string fileName = $"Unsaved_{timeStamp}.xlsx";
                    string destPath = Path.Combine(destFolder, fileName);
                    Wb.SaveCopyAs(destPath);
                }
            }
            catch
            {
                // Silent fail, do not alert user
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
