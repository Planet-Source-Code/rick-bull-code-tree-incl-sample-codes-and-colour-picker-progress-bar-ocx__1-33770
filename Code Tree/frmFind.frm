VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2020
      TabIndex        =   10
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Fi&nd"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1910
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options:"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   3855
      Begin VB.CheckBox chkStopWhenFound 
         Caption         =   "&Stop When Search Text Is Found"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox Check3 
         Caption         =   "&Whole Word Only"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Codes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sections"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkLanguage 
         Caption         =   "Languages"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Match &Case"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Search In:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare all variables
Private varKeepSearching As Boolean 'Whether to keep searching so as user can press cancel
Private varLastFound As Integer 'Where to start searching from

Private Sub cmdClose_Click()

    On Error Resume Next 'Goto next line on error

    'Unload the find form
    Unload Me
    
    'Clean up memory resources
    Set frmFind = Nothing
    
End Sub

Private Sub cmdFind_Click()
    
    On Error Resume Next 'Goto next line on error

    'If the button captions is Find
    If cmdFind.Caption = "Fi&nd" Then
        'Set the caption to cancel
        cmdFind.Caption = "Ca&ncel"
        'Set the keep searching variable to true
        varKeepSearching = True
        'If the case is to be matched
        If chkCase.Value = 1 Then
            'Find it
            Call Find(vbBinaryCompare)
        'If it isn't
        Else
            'Find it
            Call Find(vbTextCompare)
        End If
    'If the caption is Cancel
    Else
        'Set the keep searching variable to false
        varKeepSearching = False
    End If
    
End Sub

Private Sub Find(Optional ByVal TextCompareMethod As VbCompareMethod = vbTextCompare)

    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer 'For loops
    
    'Loop for all nodes from the last found
    For Counter = varLastFound + 1 To frmMain.tvCodeType.Nodes.Count
        'If the user has pressed cancel exit loop
        If varKeepSearching = False Then
            Exit For
            'Set the last found to the first node
            varLastFound = 0
        End If
        'If the find text is in the current node
        If InStr(1, frmMain.tvCodeType.Nodes(Counter).Text, txtFind.Text, TextCompareMethod) > 0 Then
            'Select/Expand it
            Call frmMain.SelectNode(Counter)
            'Set the last found to the current node number
            varLastFound = Counter
            'If the user wants to stop when found
            If chkStopWhenFound.Value = 1 Then Exit For
        'if the text is not found and the loop is finished
        ElseIf InStr(1, frmMain.tvCodeType.Nodes(Counter).Text, txtFind.Text, TextCompareMethod) = 0 And Counter = frmMain.tvCodeType.Nodes.Count Then
            'Set the last found to the first node
            varLastFound = 0
            'Tell the user
            MsgBox "The search text could not be found.", vbInformation + vbOKOnly, "Not Found"
        End If
        'Let the form redraw and the user press cancel if needed
        DoEvents
    'Onto next node
    Next Counter
    'Once the find process has finished set the caption back
    cmdFind.Caption = "Fi&nd"
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next 'Goto next line on error
    
    'Set the last found to the first node
    varLastFound = 0
    If frmMain.MnuTools_OnTop.Checked = True Then Call OnTop(Me, True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next 'Goto next line on error

    'Get rid of memory resources
    Set frmFind = Nothing
    
End Sub

Private Sub txtFind_Change()

    On Error Resume Next 'Goto next line on error
    
    'Set the last found to the first node
    varLastFound = 0
    
End Sub
