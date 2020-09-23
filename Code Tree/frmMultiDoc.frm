VERSION 5.00
Begin VB.Form frmMultiDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multiple Files Dropped!"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiDoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1455
   End
   Begin VB.ComboBox cmbFiles 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmMultiDoc.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "You dropped multiple files onto the window, please select which file you would like to load from the list below:"
      Height          =   390
      Left            =   720
      TabIndex        =   1
      Top             =   170
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMultiDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOKCancel_Click(Index As Integer)

    If Index = 0 Then
        Select Case frmMain.tbsCodeTabs.SelectedItem.Key
            Case "Code"
                frmMain.rtfCode.LoadFile cmbFiles.Text
            Case "Notes"
                frmMain.rtfNotes.LoadFile cmbFiles.Text
        End Select
    End If
    Unload Me
    
End Sub
