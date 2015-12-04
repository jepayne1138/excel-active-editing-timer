using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EditingTimer
{
    class WorkbookTimers
        /* Total elapsed time is updated on save, the session elapsed time is
         * the total working time since the last save (excluding idle time)
         */
    {
        private TimeSpan elapsedTime;
        private TimeSpan lastChangeTimeSpan;
        private Stopwatch changeStopwatch;
        private TimeSpan sessionElapsedTime;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public WorkbookTimers(double elapsedTimeSeconds)
        {
            log.Info("WorkbookTimers - Start");
            elapsedTime = TimeSpan.FromSeconds(elapsedTimeSeconds);
            sessionElapsedTime = new TimeSpan(0);  // New TimeSpan of length zero for the session
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch = new Stopwatch();
            changeStopwatch.Start();
            log.Info("WorkbookTimers - Finish");
        }


        public void SaveChangeTime()
        {
            // log.Info("SaveChangeTime - Start");
            lastChangeTimeSpan = changeStopwatch.Elapsed;
            // log.Info("SaveChangeTime - Finish");
        }

  
        public void SaveLastChangeAndReset()
        {
            log.Info("SaveLastChangeAndReset - Start");
            sessionElapsedTime += lastChangeTimeSpan;
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch.Restart();
            log.Info("SaveLastChangeAndReset - Finish");
        }


        public bool IdleTimeGreaterThan(double idleTimeSeconds)
        {
            // log.Info("IdleTimeGreaterThan - Start");
            TimeSpan idleTimeSpan = TimeSpan.FromSeconds(idleTimeSeconds);
            // log.Info("IdleTimeGreaterThan - Returning");
            return idleTimeSpan.CompareTo(changeStopwatch.Elapsed - lastChangeTimeSpan) < 0;
        }


        public TimeSpan ElapsedSinceLastChange()
        {
            log.Info("ElapsedSinceLastChange - Start");
            log.Info("ElapsedSinceLastChange - Returning");
            return changeStopwatch.Elapsed - lastChangeTimeSpan;
        }

        public double SaveSession()
        {
            log.Info("SaveSession - Start");
            elapsedTime += sessionElapsedTime + lastChangeTimeSpan;

            // Reset the session
            sessionElapsedTime = new TimeSpan(0);
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch.Restart();

            // Return the total elapsed time in minutes
            log.Info("SaveSession - Returning");
            return elapsedTime.TotalSeconds;
        }

        public void Close()
        {
            log.Info("Close - Start");
            changeStopwatch.Stop();
            log.Info("Close - Finish");
        }

        public TimeSpan GetElapsedTime()
        {
            log.Info("GetElapsedTime - Start");
            log.Info("GetElapsedTime - Returning");
            return elapsedTime;
        }
    }
}
