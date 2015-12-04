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
using log4net;

namespace EditingTimer
{
    public partial class ThisAddIn
    {
        const string ELAPSEDTIME = "ElapsedTime";
        const string PREVFULLNAME = "PrevFullName";

        private Dictionary<string, WorkbookTimers> timers;  // Holds WorkbookTimers for all open workbooks
        private ExposedFunctions exposedFunctions;  // Exposed class for VBA code to interface with
        private double idleTimeoutSeconds;
        private Collection<string> blacklist;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        // Exposes a class as a COM instance for macros to interface with
        protected override object RequestComAddInAutomationService()
        {
            log.Info("RequestComAddInAutomationService - Start");

            if (exposedFunctions == null)
            {
                exposedFunctions = new ExposedFunctions();
            }
            return exposedFunctions;

            log.Info("RequestComAddInAutomationService - Start");
        }


        private void LoadConfig()
        {
            log.Info("LoadConfig - Start");

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

            log.Info("LoadConfig - Finish");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Logging
            log4net.Config.XmlConfigurator.Configure();

            log.Info(String.Format("ThisAddIn_Startup - Start (sender: {0})", sender.ToString()));

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
            this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
            this.Application.SheetChange += new Excel.AppEvents_SheetChangeEventHandler(Application_SheetChange);

            // Need to iterate though all open workbooks on startup, find any not yet saved, and run handler on those
            foreach (Excel.Workbook Wb in Application.Workbooks)
            {
                Application_WorkbookOpenOrNew(Wb);
            }

            log.Info("ThisAddIn_Startup - Finish");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) {
            log.Info("ThisAddIn_Shutdown\n\n--------------------------------------------------------------------------------------------------------------------------------------\n");
        }

        /// <summary>
        /// Initializes the Workbook timers upon opening or creating a new workbook
        /// </summary>
        /// <param name="Wb"></param>
        private void Application_WorkbookOpenOrNew(Excel.Workbook Wb)
        {
            log.Info(String.Format("Application_WorkbookOpenOrNew: <{0}>", Wb.Name));

            // Check if the workbook name is blacklisted (or in XLSTART)
            if (blacklist.Contains(Wb.Name.ToLower()) || Wb.Path == Application.StartupPath) { return; }

            double elapsedMinutes;

            // Try to get the stored elapsed time for this Workbook, or initialize it to zero if it doesn't exist
            if (CustomDocumentPropertiesContain(Wb, ELAPSEDTIME))
            {
                elapsedMinutes = Wb.CustomDocumentProperties(ELAPSEDTIME).Value;
            }
            else
            {
                // The CustomDocumentProperty does not exist, so created it and initialize to zero
                elapsedMinutes = 0;
                Wb.CustomDocumentProperties.Add(
                    Name: ELAPSEDTIME, LinkToContent: false,
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

            log.Info("Application_WorkbookOpenOrNew - Finish");
        }

        private void Application_SheetChange(object Ws, Excel.Range Target)
        {
            // log.Info("Application_SheetChange - Start");
            WorkbookChangeHandler(((Excel.Worksheet)Ws).Parent);
            // log.Info("Application_SheetChange - Finish");
        }

        public void WorkbookChangeHandler(Excel.Workbook Wb)
        {
            // log.Info("WorkbookChangeHandler - Start");

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

            // log.Info("WorkbookChangeHandler - Finish");
        }

        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUi, ref bool Cancel)
        {
            log.Info(String.Format("Application_WorkbookBeforeSave: <{0}>", Wb.Name));
            WorkbookBeforeSaveHandler(Wb);
            log.Info("Application_WorkbookBeforeSave - Finish");
        }

        /// <summary>
        /// Before the workbook is saved, finalize the elapsed time, write the ElapsedTime CustomDocumentProperty
        /// </summary>
        /// <param name="Wb"></param>
        /// <param name="Cancel"></param>
        public void WorkbookBeforeSaveHandler(Excel.Workbook Wb)
        {
            log.Info("WorkbookBeforeSaveHandler - Start");

            // Checks that the Workbook is being timed
            if (!timers.ContainsKey(Wb.FullName)) { return; }

            // Save the current Workbook FullName as previous name CustomDocumentProperty
            // We will check this value after the save to update the timers Dictionary
            if (CustomDocumentPropertiesContain(Wb, PREVFULLNAME))
            {
                // Delete existing CustomDocumentProperty if it exists
                Wb.CustomDocumentProperties(PREVFULLNAME).Delete();
            }
            // Add a new CustomDocumentProperty saving the current FullName of the workbook
            Wb.CustomDocumentProperties.Add(
                Name: PREVFULLNAME, LinkToContent: false,
                Type: Office.MsoDocProperties.msoPropertyTypeString, Value: Wb.FullName);

            double elapsedSeconds;

            // Finalize the elasped time
            // First treat as a SheetChange and make sure the timer is up to date
            WorkbookChangeHandler(Wb);
            elapsedSeconds = timers[Wb.FullName].SaveSession();

            // Write the elapsed time to a the CustomDocumentProperty
            Wb.CustomDocumentProperties(ELAPSEDTIME).Value = elapsedSeconds;
            log.Info("WorkbookBeforeSaveHandler - Finish");
        }

        private void Application_WorkbookAfterSave(Excel.Workbook Wb, bool success)
        {
            log.Info(String.Format("Application_WorkbookAfterSave: <{0}>", Wb.Name));
            WorkbookAfterSaveHandler(Wb, success);
            log.Info("Application_WorkbookAfterSave - Finish");
        }

        public void WorkbookAfterSaveHandler(Excel.Workbook Wb, bool success)
        {
            log.Info("WorkbookAfterSaveHandler - Start");

            if (!CustomDocumentPropertiesContain(Wb, PREVFULLNAME)) { return; }

            string prevFullName;

            prevFullName = Wb.CustomDocumentProperties(PREVFULLNAME).Value;

            // Check if any timer was counting for the file with the previous name
            // If so, change the Dictionary key to be the new Workbook's FullName
            if (timers.ContainsKey(prevFullName))
            {
                WorkbookTimers tempWorkbookTimers = timers[prevFullName];
                timers.Remove(prevFullName);
                timers.Add(Wb.FullName, tempWorkbookTimers);
            }

            // Delete the PrevFullName CustomDocumentProperty
            Wb.CustomDocumentProperties(PREVFULLNAME).Delete();

            log.Info("WorkbookAfterSaveHandler - Finish");
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            log.Info(String.Format("Application_WorkbookBeforeClose: <{0}>", Wb.Name));
            BeforeCloseHandler(Wb);
            log.Info("Application_WorkbookBeforeClose - Finish");
        }

        public void BeforeCloseHandler(Excel.Workbook Wb)
        {
            log.Info("BeforeCloseHandler - Start");

            if (timers.ContainsKey(Wb.FullName))
            {
                // Clean up the WorkbookTimers instance
                timers[Wb.FullName].Close();
                timers.Remove(Wb.FullName);
            }
            log.Info("BeforeCloseHandler - Finish");
        }

        public string GetElapsedTime(Excel.Workbook Wb)
        {
            log.Info("GetElapsedTime - Start");

            if (timers.ContainsKey(Wb.FullName))
            {
                TimeSpan elapsed = timers[Wb.FullName].GetElapsedTime();

                log.Info("GetElapsedTime - Returning");
                return string.Format("{0}:{1}",
                    (elapsed.TotalMinutes > 0) ? Math.Truncate(elapsed.TotalMinutes).ToString() : "0",
                    elapsed.ToString("ss"));
            }
            else
            {
                return "0:00";
            }
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
