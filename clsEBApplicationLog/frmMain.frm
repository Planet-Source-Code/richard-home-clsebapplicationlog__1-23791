VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "clsEBApplicationLog Test"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox chkLogging 
      Caption         =   "Enable Logging"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oLog As clsEBApplicationLog

Private Sub chkLogging_Click()

    oLog.Logging chkLogging.Value = vbChecked, Log_Append
    
End Sub

Private Sub cmdClose_Click()

    oLog.Log "cmdClose_Click"
    Unload Me
    
End Sub

Private Sub Command1_Click()

Static lngClickCount As Long

    lngClickCount = lngClickCount + 1
    oLog.Log "Command1_Click: ClickCount=" & lngClickCount
    
End Sub

Private Sub Form_Load()

    Set oLog = New clsEBApplicationLog
    oLog.Logging True, Log_Overwrite

    
End Sub

Private Sub Text1_Change()

    oLog.Log "Text1_Change: " & Text1.Text
    
End Sub

Private Sub Text1_GotFocus()

    oLog.Log "Text1_GotFocus"
    
End Sub

Private Sub Text1_LostFocus()

    oLog.Log "Text1_LostFocus"
    
End Sub
