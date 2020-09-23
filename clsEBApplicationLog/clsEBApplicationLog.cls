VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEBApplicationLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Description]
'   Application Log manager Class

'[Author]
'   Richard Allsebrook  <RA>    richardallsebrook@earlybirdmarketing.com

'[History]
'   05/06/2001  V1.0.0
'   Initial Release

'[Declarations]
Option Explicit

Private strFilename         As String           'File to output log to
Private flgLogging          As Boolean           'Current logging mode
Private dtLogStart          As Date             'Time current log began

Public Enum eLogAppendOverwrite                 'Log file creating mode
    Log_Append
    Log_Overwrite
End Enum

Private Sub Class_Initialize()

'[Description]
'   Initialise the object

'[Code]
    
    'Default filename
    strFilename = App.Path & "\" & App.Title & ".log"
    
    'Logging is disabled until explicitly enabled
    flgLogging = False
    
End Sub

Public Function Logging(State As Boolean, Optional AppendOverwrite As eLogAppendOverwrite = Log_Overwrite)
Attribute Logging.VB_Description = "Enable/Disable logging"

'[Description]
'   Enable/Disable logging

'[Code]

    If flgLogging And Not State Then
        'User was logging, but has now stopped
        MarkStopLogging
    End If
    
    flgLogging = State
    
    If flgLogging Then
        'User has enabled logging
    
        If AppendOverwrite = Log_Overwrite Then
            'Overwrite existing log file (if it exists)
        
            On Error Resume Next
            Kill strFilename
            On Error GoTo 0
            
        End If
        
        'Log header
        MarkStartLogging
        
    End If
            
End Function

Public Function Log(ByVal Message As String)
Attribute Log.VB_Description = "Append a message to the current log file"

'[Description]
'   Append a message to the current log file

'[Declarations]
Dim fh                      As Integer          'File handle of log file

'[Code]

    On Error GoTo ErrorTrap
    
    If flgLogging Then
        'Logging currently enabled - append message to log file
        fh = FreeFile
        
        Open strFilename For Append As #fh
        
        Message = Format(Now, "HH:MM:SS") & ": " & Message
        Print #fh, Message
        
        Close #fh
        
    End If
    
    Exit Function:
    
ErrorTrap:

    Err.Raise Err.Number, "clsEBApplicationLog_Log", Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Private Sub Class_Terminate()

'[Description]
'   Close any open log files

'[Code]

    If flgLogging Then
        MarkStopLogging
    End If
    
End Sub

Private Function MarkStartLogging()

'[Description]
'   Append a header to the current log file

'[Code]

    'Record start time
    dtLogStart = Now
    
    'Append header
    Log "### Logging Started: " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " on " & Format(Date, "DD MMM YYYY")
    Log "### " & App.Path & "\" & App.EXEName & ".exe"
        
End Function

Private Function MarkStopLogging()

'[Description]
'   Append a footer to the current log file

'[Code]

    Log "### Logging stopped"
    Log "### Total Logging Time: " & DateDiff("s", dtLogStart, Now) & " seconds" & vbCrLf
End Function

Public Property Get Filename() As String
Attribute Filename.VB_Description = "Get/Set the current log filename"

'[Description]
'   Return the current log filename

'[Code]

    Filename = strFilename
    
End Property

Public Property Let Filename(NewValue As String)

'[Description]
'   Set a new log filename

'[Code]

    Filename = NewValue
    
End Property
