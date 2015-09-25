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
        public void Test(Excel.Workbook Wb)
        {
            System.Windows.Forms.MessageBox.Show(Wb.FullName);
        }

        public void WorkbookUpdateTimer(Excel.Workbook Wb)
        {
            Globals.ThisAddIn.WorkbookChangeHandler(Wb);
        }

        public void WorkbookSaveTimer(Excel.Workbook Wb)
        {
            Globals.ThisAddIn.WorkbookSaveHandler(Wb);
        }

        public void WorkbookCloseTimer(Excel.Workbook Wb)
        {
            Globals.ThisAddIn.BeforeCloseHandler(Wb);
        }

        public string WorkbookElapsedTime(Excel.Workbook Wb, string FormatString = "")
        {
            // Refactored so ThisAddIn.GetElapsedTime always returns the same formatted string.
            // Maintained FormatString parameter as an optional argument for backwards compatablity
            return Globals.ThisAddIn.GetElapsedTime(Wb);
        }
    }
}
