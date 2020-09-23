VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5700
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   2940
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CodeTree.ProgressBar pbProgress 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2680
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   450
      Caption         =   "0%"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrLabelPos 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait - Loading                Please Wait - Loading"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4560
      TabIndex        =   0
      Top             =   1320
      Width           =   4350
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Show
    DoEvents
    Load frmMain
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmOptions.lblSound_Click", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", _
        "Sorry there has been an error, Code Tree could not be loaded properly." & vbNewLine & _
        vbNewLine & "Press exit to unload the program" & vbNewLine & "Press Details for Details", "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog Or edfDisableOK)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    Set frmSplash = Nothing
    
End Sub

Private Sub tmrLabelPos_Timer()

    On Error Resume Next
    
    If lblInfo.Left > Int("-" & lblInfo.Width) Then
        lblInfo.Move lblInfo.Left - 50
    Else
        lblInfo.Move Me.Width
    End If
    
End Sub
