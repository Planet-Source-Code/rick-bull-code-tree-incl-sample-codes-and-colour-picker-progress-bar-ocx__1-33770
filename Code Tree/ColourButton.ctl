VERSION 5.00
Begin VB.UserControl ColourButton 
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   915
   ScaleHeight     =   705
   ScaleWidth      =   915
   ToolboxBitmap   =   "ColourButton.ctx":0000
   Begin VB.Timer tmrCursorPos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.Image imgDisabled 
      Height          =   300
      Left            =   100
      Picture         =   "ColourButton.ctx":0312
      Top             =   10
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Line lnRight 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   620
      X2              =   620
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line lnBottom 
      BorderColor     =   &H00808080&
      Visible         =   0   'False
      X1              =   0
      X2              =   630
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   620
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lnLeft 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   300
   End
   Begin VB.Shape shpColour 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   190
      Left            =   70
      Top             =   60
      Width           =   340
   End
   Begin VB.Image imgArrow 
      Height          =   240
      Left            =   470
      Picture         =   "ColourButton.ctx":06F5
      Top             =   40
      Width           =   120
   End
   Begin VB.Image imgBack 
      Height          =   285
      Left            =   20
      Picture         =   "ColourButton.ctx":0A52
      Top             =   20
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "ColourButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long 'API for finding the hWnd of the window under the cursor
Public Event ButtonDown()
Public Event ButtonUp()

Private Sub imgArrow_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.BorderColor = &H808080 Then
        Call MouseUp
    Else
        Call MouseDown
    End If
    
End Sub

Private Sub imgArrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.Visible = False Then
        lnLeft.Visible = True
        lnTop.Visible = True
        lnBottom.Visible = True
        lnRight.Visible = True
        tmrCursorPos.Enabled = True
    End If
    
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.BorderColor = &H808080 Then
        Call MouseUp
    Else
        Call MouseDown
    End If
    
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.Visible = False Then
        lnLeft.Visible = True
        lnTop.Visible = True
        lnBottom.Visible = True
        lnRight.Visible = True
        tmrCursorPos.Enabled = True
    End If
    
End Sub

Private Sub tmrCursorPos_Timer()

    Dim CursorPos As POINTAPI
    
    Call GetCursorPos(CursorPos)
    If WindowFromPoint(CursorPos.X, CursorPos.Y) <> UserControl.hwnd _
        Then Call MouseOut
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.BorderColor = &H808080 Then
        Call MouseUp
    Else
        Call MouseDown
    End If
    
End Sub

Public Sub MouseDown()

    lnLeft.BorderColor = &H808080
    lnTop.BorderColor = &H808080
    lnBottom.BorderColor = &HFFFFFF
    lnRight.BorderColor = &HFFFFFF
    imgBack.Visible = True
    UserControl.BackColor = &HFFFFFF
    imgArrow.Move 470 + Screen.TwipsPerPixelX, 40 + Screen.TwipsPerPixelY
    shpColour.Move 70 + Screen.TwipsPerPixelX, 60 + Screen.TwipsPerPixelY
    RaiseEvent ButtonDown
    
End Sub

Public Sub MouseUp()

    lnLeft.BorderColor = &HFFFFFF
    lnTop.BorderColor = &HFFFFFF
    lnBottom.BorderColor = &H808080
    lnRight.BorderColor = &H808080
    imgBack.Visible = False
    UserControl.BackColor = &H8000000F
    imgArrow.Move 470, 40
    shpColour.Move 70, 60
    If lnLeft.Visible = False Then
        lnLeft.Visible = True
        lnTop.Visible = True
        lnBottom.Visible = True
        lnRight.Visible = True
        tmrCursorPos.Enabled = True
    End If
    RaiseEvent ButtonUp
    
End Sub


Public Property Get Colour() As OLE_COLOR

    Colour = shpColour.FillColor
    
End Property

Public Property Let Colour(ByVal NewColour As OLE_COLOR)

    shpColour.FillColor = NewColour
    PropertyChanged ("Colour")
    
End Property

Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

    UserControl.Enabled = NewEnabled
    imgDisabled.Visible = Not NewEnabled
    PropertyChanged ("Enabled")

End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If lnLeft.Visible = False Then
        lnLeft.Visible = True
        lnTop.Visible = True
        lnBottom.Visible = True
        lnRight.Visible = True
        tmrCursorPos.Enabled = True
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Colour = PropBag.ReadProperty("Colour", &HFFFFFF)
    Enabled = PropBag.ReadProperty("Enabled", True)

End Sub

Private Sub UserControl_Resize()

    UserControl.Height = 315
    UserControl.Width = 630
    
End Sub

Private Sub MouseOut()

    If lnLeft.Visible = True And UserControl.BackColor <> &HFFFFFF Then
        lnLeft.Visible = False
        lnTop.Visible = False
        lnBottom.Visible = False
        lnRight.Visible = False
        tmrCursorPos.Enabled = False
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Colour", shpColour.FillColor, &HFFFFFF)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    
End Sub
