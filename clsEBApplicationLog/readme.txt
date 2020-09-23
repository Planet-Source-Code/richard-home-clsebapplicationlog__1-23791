clsEBApplicationLog

Description:
This self contained Class is designed to simplify creating application logs.


Usage:
See the enclosed demonstration for a working example

'Create an instance of the class
'This must be declared as a global variable (in a module) if you intend
'logging across the whole application
Public oLog As New clsEBApplication Log

'(Optional) set the filename - if you don't a default will be used
oLog.Filename = "C:/mylog.txt"

'Enable logging (overwrite any existing log)
oLog.Logging True, Log_Overwrite

'An error will be raised if the file cannot be written to

'Log Some events
oLog.Log "This is a test message"
oLog.Log "And another"

'Stop Logging
oLog.Logging False

'or
Set oLog = Nothing


Properties:
Filename
Get/Set the file to output logging to


Methods:
Logging (State, AppendOverwrite)
Starts (State=True) or Stops (State=False) logging

If starting to log, AppendOverwrite is used to specify how exisiting log files are handled.


Contact:
If you have any comments, suggestions or improvements please send them to me at:
RichardAllsebrook@earlybirdmarketing.com

You will, of course, receive full credit for your contribution.