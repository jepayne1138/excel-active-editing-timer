using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using AppEvents_Event = Microsoft.Office.Interop.Excel.AppEvents_Event;
using System.IO;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;
using System.Collections.ObjectModel;

namespace EditingTimer
{
    public partial class ThisAddIn
    {
        private Dictionary<string, WorkbookTimers> timers;  // Holds WorkbookTimers for all open workbooks
        private ExposedFunctions exposedFunctions;  // Exposed class for VBA code to interface with
        private double idleTimeoutSeconds;
        private Collection<string> blacklist;


        // Exposes a class as a COM instance for macros to interface with
        protected override object RequestComAddInAutomationService()
        {
            if (exposedFunctions == null)
            {
                exposedFunctions = new ExposedFunctions();
            }
            return exposedFunctions;
        }


        private void LoadConfig()
        {
            string configPath = Path.Combine(Environment.GetEnvironmentVariable("UserProfile"), "Configs", "EditingTimer.yaml");
            StreamReader fileReader ;

            // Set default fallback values
            idleTimeoutSeconds = 120;
            blacklist = new Collection<string>();

            // Load the yaml stream
            YamlStream yaml = new YamlStream();
            try
            {
                using (fileReader = new StreamReader(configPath))
                {
                    yaml.Load(fileReader);
                }
            }
            catch (DirectoryNotFoundException) { return; }
            catch (FileNotFoundException) { return; }
            catch (SemanticErrorException) { return; }

            // Get the root node of the yaml document
            YamlMappingNode rootNode = (YamlMappingNode)yaml.Documents[0].RootNode;

            // Get idle timeout seconds value
            try
            {
                YamlScalarNode idleTimeoutNode = (YamlScalarNode)rootNode.Children[new YamlScalarNode("IdleTimeout")];
                idleTimeoutSeconds = Convert.ToDouble(idleTimeoutNode.Value);
            }
            catch (KeyNotFoundException) { }

            // Get blacklisted filenames
            try
            {
                YamlSequenceNode blacklistNode = (YamlSequenceNode)rootNode.Children[new YamlScalarNode("Blacklist")];
                // Convert to a Collection
                foreach (YamlScalarNode filenameNode in blacklistNode)
                {
                    blacklist.Add(filenameNode.Value.ToLower());
                }
            }
            catch (KeyNotFoundException) { }
            catch (InvalidCastException) { }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get config values
            LoadConfig();
            foreach (string filename in blacklist)
            {
                System.Diagnostics.Debug.Print(filename);
            }

            timers = new Dictionary<String, WorkbookTimers>();

            ((AppEvents_Event)this.Application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(Application_WorkbookOpenOrNew);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpenOrNew);
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
            this.Application.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(Application_SheetChange);

            // Need to iterate though all open workbooks on startup, find any not yet saved, and run handler on those
            foreach (Excel.Workbook Wb in Application.Workbooks)
            {
                Application_WorkbookOpenOrNew(Wb);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {}

        /// <summary>
        /// Initializes the Workbook timers upon opening or creating a new workbook
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpenOrNew(Excel.Workbook Wb)
        {
            // Check if the workbook name is blacklisted (or in XLSTART)
            if (blacklist.Contains(Wb.Name.ToLower()) || Wb.Path == Application.StartupPath) { return; }

            double elapsedMinutes;

            // Try to get the stored elapsed time for this Workbook, or initialize it to zero if it doesn't exist
            if (CustomDocumentPropertiesContain(Wb, "ElapsedTime"))
            {
                elapsedMinutes = Wb.CustomDocumentProperties("ElapsedTime").Value;
            }
            else
            {
                // The CustomDocumentProperty does not exist, so created it and initialize to zero
                elapsedMinutes = 0;
                Wb.CustomDocumentProperties.Add(
                    Name: "ElapsedTime", LinkToContent: false,
                    Type: Office.MsoDocProperties.msoPropertyTypeNumber, Value: elapsedMinutes);
            }

            // Create a new WorkbookTimers instance for the opened Workbook
            if (!timers.ContainsKey(Wb.FullName))
            {
                timers.Add(Wb.FullName, new WorkbookTimers(elapsedMinutes));
            }
            else
            {
                // The timer for this Workbook already existed.  This should never happen, but if so, just overwrite previous timer.
                timers.Remove(Wb.FullName);
                timers.Add(Wb.FullName, new WorkbookTimers(elapsedMinutes));
            }
        }

        private void Application_SheetChange(object Ws, Excel.Range Target)
        {
            WorkbookChangeHandler(((Excel.Worksheet)Ws).Parent);
        }

        public void WorkbookChangeHandler(Excel.Workbook Wb)
        {
            // Checks that the Workbook is being timed
            if (!timers.ContainsKey(Wb.FullName)) { return; }

            WorkbookTimers wbTimer = timers[Wb.FullName];

            if (wbTimer.IdleTimeGreaterThan(idleTimeoutSeconds))
            {
                TimeSpan elapsedTime = wbTimer.ElapsedSinceLastChange();

                string message = "Time since last change to this document: " + elapsedTime.ToString(@"m\:ss") +
                    Environment.NewLine + "Should this time be included in processing time?";
                string caption = "Idle Timeout Reached";
                MessageBoxIcon icon = MessageBoxIcon.Warning;
                DialogResult result;
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button2;
                result = MessageBox.Show(message, caption, buttons, icon, defaultButton);

                if (result == DialogResult.Yes)
                {
                    wbTimer.SaveChangeTime();
                }
                else
                {
                    wbTimer.SaveLastChangeAndReset();
                }
            }
            else
            {
                // If we haven't reached the idle timeout, just save the elapsed time
                wbTimer.SaveChangeTime();
            }
        }
        
        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUi, ref bool Cancel)
        {
            WorkbookSaveHandler(Wb);
        }

        /// <summary>
        /// Before the workbook is saved, finalize the elapsed time, write the ElapsedTime CustomDocumentProperty
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Cancel"></param>
        public void WorkbookSaveHandler(Excel.Workbook Wb)
        {
            // Checks that the Workbook is being timed
            if (!timers.ContainsKey(Wb.FullName)) { return; }

            double elapsedSeconds;

            // Finalize the elasped time
            // First treat as a SheetChange and make sure the timer is up to date
            WorkbookChangeHandler(Wb);
            elapsedSeconds = timers[Wb.FullName].SaveSession();

            // Write the elapsed time to a the CustomDocumentProperty
            Wb.CustomDocumentProperties("ElapsedTime").Value = elapsedSeconds;
            }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            BeforeCloseHandler(Wb);
        }

        public void BeforeCloseHandler(Excel.Workbook Wb)
        {
            if (timers.ContainsKey(Wb.FullName))
            {
                // Clean up the WorkbookTimers instance
                timers[Wb.FullName].Close();
                timers.Remove(Wb.FullName);
            }
        }

        public string GetElapsedTime(Excel.Workbook Wb)
        {
            TimeSpan elapsed = timers[Wb.FullName].GetElapsedTime();

            return string.Format("{0}:{1}",
                (elapsed.Minutes > 0) ? Math.Truncate(elapsed.TotalMinutes).ToString() : "0",
                elapsed.ToString("ss"));
        }

        // ==============================  CustomDocumentProperties Helpers  =======================================

        /// <summary>
        /// Tests if a Workbook contains a CustomDocumentProperty with the given name
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="PropName"></param>
        /// <returns></returns>
        private bool CustomDocumentPropertiesContain(Excel.Workbook Wb, string PropName)
        {
            foreach (Office.DocumentProperty property in Wb.CustomDocumentProperties)
            {
                if (property.Name == PropName)
                {
                    return true;
                }
            }

            return false;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
