# Excel Active-editing Timer
Logs time spent actively editing a workbook as CustomDocumentProperty("ElapsedTime") with idle prompt after 120 seconds idle.

I was unsatisfied with the built in timer in MS Office, as I wanted to record only the time that I spend activily working with a Workbook.  In situations where distractions arose and the file was left open, the timer would continue to count this as time the file was open.  I therefore implemented this addin that records all time the file is open, but if 2 minutes (by default) passes without a WorksheetChanged event being triggered, the next edit to the Worksheet with display a prompt asking if they elapsed time with no active changes being made should be included with the total time.

## Configuration
The default time can be changed by creating a yaml configuration file "Configs\EditingTimer.yaml" in the current user directory (e.g. "C:\Users\currentuser\").

Currently Implemented Configurations:
* IdleTimeout - Time in seconds before the timeout pop-up will be triggered.
* Blacklist - A list of files for which time should not be recorded.

If no configuration file is found, time will be tracked for all files and a default timeout of 120 seconds will be used.

## Dependencies
* log4net.2.0.3
* YamlDotNet.3.7.0

## Installation Instruction
Installation instruction coming soon.
