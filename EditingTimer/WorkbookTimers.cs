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

        public WorkbookTimers(double elapsedTimeSeconds)
        {
            elapsedTime = TimeSpan.FromSeconds(elapsedTimeSeconds);
            sessionElapsedTime = new TimeSpan(0);  // New TimeSpan of length zero for the session
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch = new Stopwatch();
            changeStopwatch.Start();
        }


        public void SaveChangeTime()
        {
            lastChangeTimeSpan = changeStopwatch.Elapsed;
        }

  
        public void SaveLastChangeAndReset()
        {
            sessionElapsedTime += lastChangeTimeSpan;
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch.Restart();
        }


        public bool IdleTimeGreaterThan(double idleTimeSeconds)
        {
            TimeSpan idleTimeSpan = TimeSpan.FromSeconds(idleTimeSeconds);
            return idleTimeSpan.CompareTo(changeStopwatch.Elapsed - lastChangeTimeSpan) < 0;
        }


        public TimeSpan ElapsedSinceLastChange()
        {
            return changeStopwatch.Elapsed - lastChangeTimeSpan;
        }

        public double SaveSession()
        {
            elapsedTime += sessionElapsedTime + lastChangeTimeSpan;

            // Reset the session
            sessionElapsedTime = new TimeSpan(0);
            lastChangeTimeSpan = new TimeSpan(0);
            changeStopwatch.Restart();

            // Return the total elapsed time in minutes
            return elapsedTime.TotalSeconds;
        }

        public void Close()
        {
            changeStopwatch.Stop();
        }

        public TimeSpan GetElapsedTime()
        {
            return elapsedTime;
        }
    }
}
