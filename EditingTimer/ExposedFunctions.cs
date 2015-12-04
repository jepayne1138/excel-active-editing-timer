using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using System.Data;
using System.Runtime.InteropServices;


namespace EditingTimer
{
    
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IExposedFunctions
    {
        void WorkbookUpdateTimer(Excel.Workbook Wb);
        void WorkbookSaveTimer(Excel.Workbook Wb);
        void WorkbookCloseTimer(Excel.Workbook Wb);
        string WorkbookElapsedTime(Excel.Workbook Wb, string FormatString = "");  // Remove Optional FormatSting when all ListValidation macros have been updated
        void Test(Excel.Workbook Wb);
    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ExposedFunctions: IExposedFunctions
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Test(Excel.Workbook Wb)
        {
            log.Info("Test - Start");
            System.Windows.Forms.MessageBox.Show(Wb.FullName);
            log.Info("Test - Finish");

        }

        public void WorkbookUpdateTimer(Excel.Workbook Wb)
        {
            log.Info("WorkbookUpdateTimer - Start");
            Globals.ThisAddIn.WorkbookChangeHandler(Wb);
            log.Info("WorkbookUpdateTimer - Finish");
        }

        public void WorkbookSaveTimer(Excel.Workbook Wb)
        {
            log.Info("WorkbookSaveTimer - Start");
            Globals.ThisAddIn.WorkbookBeforeSaveHandler(Wb);
            log.Info("WorkbookSaveTimer - Finish");
        }

        public void WorkbookCloseTimer(Excel.Workbook Wb)
        {
            log.Info("WorkbookCloseTimer - Start");
            Globals.ThisAddIn.DeactivateHandler(Wb);
            log.Info("WorkbookCloseTimer - Finish");
        }

        public string WorkbookElapsedTime(Excel.Workbook Wb, string FormatString = "")
        {
            // Refactored so ThisAddIn.GetElapsedTime always returns the same formatted string.
            // Maintained FormatString parameter as an optional argument for backwards compatablity
            log.Info("WorkbookElapsedTime - Start");
            log.Info("WorkbookElapsedTime - Returning");
            return Globals.ThisAddIn.GetElapsedTime(Wb);
        }
    }
}
