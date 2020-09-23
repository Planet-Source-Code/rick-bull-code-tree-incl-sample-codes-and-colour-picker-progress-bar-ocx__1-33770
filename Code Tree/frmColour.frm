VERSION 5.00
Begin VB.Form frmColour 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   1155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CodeTree.ColourPicker ColourPicker1 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   3175
      AllowCustomize  =   -1  'True
      BackStyle       =   1
   End
End
Attribute VB_Name = "frmColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ColourPicker1_Cancel()

    Unload Me
    
End Sub

Private Sub ColourPicker1_ColourChoosen(ByVal Colour As Long)

    Select Case frmMain.tbsCodeTabs.SelectedItem.Key
        Case "Code"
            frmMain.rtfCode.SelColor = Colour
        Case "Notes"
            frmMain.rtfNotes.SelColor = Colour
    End Select
    frmMain.ColourButton1.Colour = Colour
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Dim MeTop As Long
    Dim MeLeft As Long
    Dim TaskBarHwnd As Long
    Dim TaskRect As RECT
    Dim ColoursRect As RECT
    
    Me.Move frmMain.tbRTF.Buttons("Font Colour").Left + frmMain.tbRTF.Left + frmMain.Left + 70, _
        frmMain.rtfCode.Top + frmMain.Top + 650
    TaskBarHwnd = FindWindow(TaskBar, "")
    Call GetWindowRect(TaskBarHwnd, TaskRect)
    Call GetWindowRect(Me.hwnd, ColoursRect)
    
    If ColoursRect.Right > TaskRect.Right Then
        MeLeft = Screen.Width - Me.Width
    Else
        MeLeft = frmMain.tbRTF.Buttons("Font Colour").Left + frmMain.tbRTF.Left + frmMain.Left + 70
    End If
    
    If ColoursRect.Bottom > TaskRect.Top Then
        MeTop = (TaskRect.Top * Screen.TwipsPerPixelX) - Me.Height
    Else
        MeTop = frmMain.rtfCode.Top + frmMain.Top + 650
    End If
    
    Me.Move MeLeft, MeTop
    
    Select Case frmMain.tbsCodeTabs.SelectedItem.Key
        Case "Code"
            ColourPicker1.SelectedColour = frmMain.rtfCode.SelColor
        Case "Notes"
            ColourPicker1.SelectedColour = frmMain.rtfNotes.SelColor
    End Select
    
    Call ColourPicker1.GetColours(App.EXEName, "Colours")

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call ColourPicker1.SaveColours(App.EXEName, "Colours")
    Call frmMain.ColourButton1.MouseUp
    Set frmColour = Nothing
    
End Sub
