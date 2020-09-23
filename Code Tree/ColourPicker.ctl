VERSION 5.00
Begin VB.UserControl ColourPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1170
   DefaultCancel   =   -1  'True
   PropertyPages   =   "ColourPicker.ctx":0000
   ScaleHeight     =   1815
   ScaleWidth      =   1170
   ToolboxBitmap   =   "ColourPicker.ctx":0026
   Begin VB.Frame fraSelection 
      BorderStyle     =   0  'None
      Height          =   265
      Left            =   50
      TabIndex        =   22
      Top             =   50
      Width           =   265
      Begin VB.Shape shpSelection 
         Height          =   270
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   270
      End
      Begin VB.Shape shpSelection 
         BorderColor     =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   15
         Top             =   15
         Width           =   240
      End
      Begin VB.Shape shpSelection 
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   210
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdOther 
      Caption         =   "&Other..."
      Height          =   255
      Left            =   60
      TabIndex        =   20
      Top             =   1500
      Width           =   775
   End
   Begin VB.PictureBox picCustomColour 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H8000000F&
      Height          =   245
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   21
      Top             =   1500
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00A4A0A0&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   19
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   19
      Top             =   1140
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00F0FBFF&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   18
      Left            =   600
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   18
      Top             =   1140
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00F0CAA6&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   17
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   17
      Top             =   1140
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00C0DCC0&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   16
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   16
      Top             =   1140
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00800080&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   15
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   15
      Top             =   870
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00FF00FF&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   14
      Left            =   600
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   14
      Top             =   870
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00800000&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   13
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   13
      Top             =   870
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00FF0000&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   12
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   870
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00808000&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   11
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   600
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00FFFF00&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   10
      Left            =   600
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   10
      Top             =   600
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00008000&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   9
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   600
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H0000FF00&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   8
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   600
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00008080&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   7
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   7
      Top             =   330
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   6
      Left            =   600
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   330
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00000080&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   5
      Left            =   330
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   5
      Top             =   330
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H000000FF&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   4
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   4
      Top             =   330
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00808080&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   3
      Left            =   870
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   60
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00C0C0C0&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   2
      Left            =   595
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   2
      Top             =   60
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00000000&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   1
      Left            =   335
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   60
      Width           =   245
   End
   Begin VB.PictureBox picDefaultColour 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H8000000F&
      Height          =   245
      Index           =   0
      Left            =   60
      ScaleHeight     =   180
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   60
      Width           =   245
   End
   Begin VB.Line lnSeperator 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   55
      X2              =   1085
      Y1              =   1430
      Y2              =   1430
   End
   Begin VB.Line lnSeperator 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   70
      X2              =   1085
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "ColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private varSelectionPos As Integer 'Which colour the mouse is over
Private varAllowCustomize As Boolean
Public Event ColourChoosen(ByVal Colour As Long)
Public Event MouseMove(ByVal Colour As Long, X As Single, Y As Single, Button As Integer, Shift As Integer)
Public Event MouseDown(ByVal Colour As Long, X As Single, Y As Single, Button As Integer, Shift As Integer)
Public Event MouseUp(ByVal Colour As Long, X As Single, Y As Single, Button As Integer, Shift As Integer)
Public Event CustomClick()
Public Event Cancel()
Public Enum BackStyleConsts
    Transparent = 0
    Opaque = 1
End Enum
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" _
    Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long 'API needed for showing the colour dialog
Private Type CHOOSECOLOR 'Type for holding show-colour info
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Const CC_FULLOPEN = &H2 'Constant for showing the full open dialog
Private Const CC_NOHELP = 2048 'Constant for not showing help button

Public Function ShowColor(FormHwnd As Long) As Long
    
    On Error Resume Next
    Dim ColourDialog As CHOOSECOLOR
    Dim CustomColors() As Byte
    Dim Counter As Integer
    
    'Required to use custom colors
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte
    'Loop for all custom colours
    For Counter = LBound(CustomColors) To UBound(CustomColors)
        'Set all custom colours to white
        CustomColors(Counter) = 255
    'Onto next colour
    Next Counter
    
    ColourDialog.lStructSize = Len(ColourDialog)
    ColourDialog.hwndOwner = FormHwnd
    ColourDialog.hInstance = App.hInstance
    ColourDialog.lpCustColors = StrConv(CustomColors, vbUnicode)
    ColourDialog.flags = CC_FULLOPEN Or CC_NOHELP
    If CHOOSECOLOR(ColourDialog) <> 0 Then
        ShowColor = ColourDialog.rgbResult
        CustomColors = StrConv(ColourDialog.lpCustColors, vbFromUnicode)
    'Cancel
    Else
        ShowColor = -1
    End If
    
End Function

Private Sub cmdOther_Click()

    On Error Resume Next
    
    picCustomColour.SetFocus
    'Show the colour dialog
    Call ShowColourDialog
    
End Sub

Private Sub fraSelection_Click()
    
    On Error Resume Next
    
    Call ColourClicked
    
End Sub

Private Sub fraSelection_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    Dim TheColour As Long
    
    If Button = vbRightButton And AllowCustomize = True And _
        varSelectionPos <= picDefaultColour.UBound Then
        TheColour = ShowColor(UserControl.hwnd)
        If TheColour > -1 Then
            picDefaultColour(varSelectionPos).BackColor = TheColour
            shpSelection(2).FillColor = TheColour
        End If
    End If
        
End Sub

Private Sub picCustomColour_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    'If the selection needs moving move the frame and set the shape's colour
    If varSelectionPos <> picDefaultColour.UBound + 1 Then
        shpSelection(2).FillColor = picCustomColour.BackColor
        fraSelection.Move picCustomColour.Left - Screen.TwipsPerPixelX, picCustomColour.Top - Screen.TwipsPerPixelY
        varSelectionPos = picDefaultColour.UBound + 1
    End If

End Sub

Private Sub picDefaultColour_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    'If the selection needs moving move the frame and set the shape's colour
    If varSelectionPos <> Index Then
        shpSelection(2).FillColor = picDefaultColour(Index).BackColor
        fraSelection.Move picDefaultColour(Index).Left - Screen.TwipsPerPixelX, picDefaultColour(Index).Top - Screen.TwipsPerPixelY
        varSelectionPos = Index
    End If

End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    Select Case KeyAscii
        Case vbKeyEscape
            RaiseEvent Cancel
        Case vbKeyReturn
            Call ColourClicked
    End Select

End Sub

Private Sub UserControl_ExitFocus()
    
    On Error Resume Next
    
    RaiseEvent Cancel
    
End Sub

Private Sub UserControl_Initialize()

    On Error Resume Next
    
    varSelectionPos = picDefaultColour.LBound

End Sub

Private Sub ShowColourDialog()

    On Error Resume Next
    Dim TheColour As Long
    
    TheColour = ShowColor(UserControl.hwnd)
    If TheColour > -1 Then
        picCustomColour.BackColor = TheColour
        shpSelection(2).FillColor = TheColour
        fraSelection.Move picCustomColour.Left - Screen.TwipsPerPixelX, picCustomColour.Top - Screen.TwipsPerPixelY
        varSelectionPos = picDefaultColour.UBound + 1
        'Raise the event with the colour of the selected picturebox's background/shpSelection(2)'s fill colour
        RaiseEvent ColourChoosen(shpSelection(2).FillColor)
    End If
    
End Sub

Public Property Get SelectedColour() As OLE_COLOR

    On Error Resume Next
    
    'Send back the selection box's fill colour
    SelectedColour = shpSelection(2).FillColor
    
End Property

Public Property Let SelectedColour(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    Dim Counter As Integer 'For loops
    
    'Loop for all picture boxes
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        'If the new colour is the same as that of the current picbox's backcolour
        If NewColour = picDefaultColour(Counter).BackColor Then
            'Make out that the mouse has been moved over it to set the selection
            Call picDefaultColour_MouseMove(Counter, 0, 0, 0, 0)
            'Exit loop
            Exit For
        
        'Else if the counter is at the last picbox and the colour is not = to it's backcolour
        ElseIf Counter = picDefaultColour.UBound And NewColour <> picDefaultColour(Counter).BackColor Then
            'Make out that the mouse has been moved over the custom colour picbox to set the selection
            Call picCustomColour_MouseMove(0, 0, 0, 0)
            'Set it's colour
            picCustomColour.BackColor = NewColour
        End If
    'On to next picbox
    Next Counter
    'Set the selection box's fill colour
    shpSelection(2).FillColor = NewColour
    PropertyChanged ("SelectedColour")
    
End Property

Private Sub UserControl_InitProperties()

    On Error Resume Next
    
    AllowCustomize = False
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    On Error Resume Next
    
    AllowCustomize = PropBag.ReadProperty("AllowCustomize", False)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", Transparent)
    CustomColour = PropBag.ReadProperty("CustomColour", RGB(255, 255, 255))
    DefaultColour01 = PropBag.ReadProperty("DefaultColour01", &HFFFFFF)
    DefaultColour02 = PropBag.ReadProperty("DefaultColour02", &H0&)
    DefaultColour03 = PropBag.ReadProperty("DefaultColour03", &HC0C0C0)
    DefaultColour04 = PropBag.ReadProperty("DefaultColour04", &H808080)
    DefaultColour05 = PropBag.ReadProperty("DefaultColour05", &HFF&)
    DefaultColour06 = PropBag.ReadProperty("DefaultColour06", &H80&)
    DefaultColour07 = PropBag.ReadProperty("DefaultColour07", &HFFFF&)
    DefaultColour08 = PropBag.ReadProperty("DefaultColour08", &H8080&)
    DefaultColour09 = PropBag.ReadProperty("DefaultColour09", &HFF00&)
    DefaultColour10 = PropBag.ReadProperty("DefaultColour10", &H8000&)
    DefaultColour11 = PropBag.ReadProperty("DefaultColour11", &HFFFF00)
    DefaultColour12 = PropBag.ReadProperty("DefaultColour12", &H808000)
    DefaultColour13 = PropBag.ReadProperty("DefaultColour13", &HFF0000)
    DefaultColour14 = PropBag.ReadProperty("DefaultColour14", &H800000)
    DefaultColour15 = PropBag.ReadProperty("DefaultColour15", &HFF00FF)
    DefaultColour16 = PropBag.ReadProperty("DefaultColour16", &H800080)
    DefaultColour17 = PropBag.ReadProperty("DefaultColour17", &HC0DCC0)
    DefaultColour18 = PropBag.ReadProperty("DefaultColour18", &HF0CAA6)
    DefaultColour19 = PropBag.ReadProperty("DefaultColour19", &HF0FBFF)
    DefaultColour20 = PropBag.ReadProperty("DefaultColour20", &HA4A0A0)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", vbDefault)
    SelectedColour = PropBag.ReadProperty("SelectedColour", RGB(255, 255, 255))
    
End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    UserControl.Height = picCustomColour.Top + picCustomColour.Height + (4 * Screen.TwipsPerPixelY)
    UserControl.Width = picDefaultColour(3).Left + picDefaultColour(3).Width + (4 * Screen.TwipsPerPixelX)
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next
    
    Call PropBag.WriteProperty("AllowCustomize", varAllowCustomize, False)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, Transparent)
    Call PropBag.WriteProperty("CustomColour", picCustomColour.BackColor, RGB(255, 255, 255))
    Call PropBag.WriteProperty("DefaultColour01", picDefaultColour(0).BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("DefaultColour02", picDefaultColour(1).BackColor, &H0&)
    Call PropBag.WriteProperty("DefaultColour03", picDefaultColour(2).BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("DefaultColour04", picDefaultColour(3).BackColor, &H808080)
    Call PropBag.WriteProperty("DefaultColour05", picDefaultColour(4).BackColor, &HFF&)
    Call PropBag.WriteProperty("DefaultColour06", picDefaultColour(5).BackColor, &H80&)
    Call PropBag.WriteProperty("DefaultColour07", picDefaultColour(6).BackColor, &HFFFF&)
    Call PropBag.WriteProperty("DefaultColour08", picDefaultColour(7).BackColor, &H8080&)
    Call PropBag.WriteProperty("DefaultColour09", picDefaultColour(8).BackColor, &HFF00&)
    Call PropBag.WriteProperty("DefaultColour10", picDefaultColour(9).BackColor, &H8000&)
    Call PropBag.WriteProperty("DefaultColour11", picDefaultColour(10).BackColor, &HFFFF00)
    Call PropBag.WriteProperty("DefaultColour12", picDefaultColour(11).BackColor, &H808000)
    Call PropBag.WriteProperty("DefaultColour13", picDefaultColour(12).BackColor, &HFF0000)
    Call PropBag.WriteProperty("DefaultColour14", picDefaultColour(13).BackColor, &H800000)
    Call PropBag.WriteProperty("DefaultColour15", picDefaultColour(14).BackColor, &HFF00FF)
    Call PropBag.WriteProperty("DefaultColour16", picDefaultColour(15).BackColor, &H800080)
    Call PropBag.WriteProperty("DefaultColour17", picDefaultColour(16).BackColor, &HC0DCC0)
    Call PropBag.WriteProperty("DefaultColour18", picDefaultColour(17).BackColor, &HF0CAA6)
    Call PropBag.WriteProperty("DefaultColour19", picDefaultColour(18).BackColor, &HF0FBFF)
    Call PropBag.WriteProperty("DefaultColour20", picDefaultColour(19).BackColor, &HA4A0A0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", picDefaultColour(0).MousePointer, vbDefault)
    Call PropBag.WriteProperty("SelectedColour", shpSelection(2).FillColor, RGB(255, 255, 255))
    
End Sub

Public Property Get CustomColour() As OLE_COLOR

    On Error Resume Next
    
    CustomColour = picCustomColour.BackColor
    
End Property

Public Property Let CustomColour(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picCustomColour.BackColor = NewColour
    If varSelectionPos = picDefaultColour.UBound + 1 Then _
    SelectedColour = NewColour
    PropertyChanged ("CustomColour")
    
End Property

Public Property Get DefaultColour01() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour01 = picDefaultColour(0).BackColor
    
End Property

Public Property Let DefaultColour01(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(0).BackColor = NewColour
    If varSelectionPos = 0 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour01")
    
End Property

Public Property Get DefaultColour02() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour02 = picDefaultColour(1).BackColor
    
End Property

Public Property Let DefaultColour02(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(1).BackColor = NewColour
    If varSelectionPos = 1 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour02")
    
End Property

Public Property Get DefaultColour03() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour03 = picDefaultColour(2).BackColor
    
End Property

Public Property Let DefaultColour03(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(2).BackColor = NewColour
    If varSelectionPos = 2 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour03")
    
End Property

Public Property Get DefaultColour04() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour04 = picDefaultColour(3).BackColor
    
End Property

Public Property Let DefaultColour04(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(3).BackColor = NewColour
    If varSelectionPos = 3 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour04")
    
End Property

Public Property Get DefaultColour05() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour05 = picDefaultColour(4).BackColor
    
End Property

Public Property Let DefaultColour05(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(4).BackColor = NewColour
    If varSelectionPos = 4 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour05")
    
End Property

Public Property Get DefaultColour06() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour06 = picDefaultColour(5).BackColor
    
End Property

Public Property Let DefaultColour06(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(5).BackColor = NewColour
    If varSelectionPos = 5 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour06")
    
End Property

Public Property Get DefaultColour07() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour07 = picDefaultColour(6).BackColor
    
End Property

Public Property Let DefaultColour07(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(6).BackColor = NewColour
    If varSelectionPos = 6 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour07")
    
End Property

Public Property Get DefaultColour08() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour08 = picDefaultColour(7).BackColor
    
End Property

Public Property Let DefaultColour08(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(7).BackColor = NewColour
    If varSelectionPos = 7 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour08")
    
End Property

Public Property Get DefaultColour09() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour09 = picDefaultColour(8).BackColor
    
End Property

Public Property Let DefaultColour09(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(8).BackColor = NewColour
    If varSelectionPos = 8 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour09")
    
End Property

Public Property Get DefaultColour10() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour10 = picDefaultColour(9).BackColor
    
End Property

Public Property Let DefaultColour10(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(9).BackColor = NewColour
    If varSelectionPos = 9 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour10")
    
End Property

Public Property Get DefaultColour11() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour11 = picDefaultColour(10).BackColor
    
End Property

Public Property Let DefaultColour11(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(10).BackColor = NewColour
    If varSelectionPos = 10 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour11")
    
End Property

Public Property Get DefaultColour12() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour12 = picDefaultColour(11).BackColor
    
End Property

Public Property Let DefaultColour12(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(11).BackColor = NewColour
    If varSelectionPos = 11 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour12")
    
End Property

Public Property Get DefaultColour13() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour13 = picDefaultColour(12).BackColor
    
End Property

Public Property Let DefaultColour13(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(12).BackColor = NewColour
    If varSelectionPos = 12 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour13")
    
End Property

Public Property Get DefaultColour14() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour14 = picDefaultColour(13).BackColor
    
End Property

Public Property Let DefaultColour14(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(13).BackColor = NewColour
    If varSelectionPos = 13 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour14")
    
End Property

Public Property Get DefaultColour15() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour15 = picDefaultColour(14).BackColor
    
End Property

Public Property Let DefaultColour15(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(14).BackColor = NewColour
    If varSelectionPos = 14 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour15")
    
End Property

Public Property Get DefaultColour16() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour16 = picDefaultColour(15).BackColor
    
End Property

Public Property Let DefaultColour16(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(15).BackColor = NewColour
    If varSelectionPos = 15 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour16")
    
End Property

Public Property Get DefaultColour17() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour17 = picDefaultColour(16).BackColor
    
End Property

Public Property Let DefaultColour17(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(16).BackColor = NewColour
    If varSelectionPos = 16 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour17")
    
End Property

Public Property Get DefaultColour18() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour18 = picDefaultColour(17).BackColor
    
End Property

Public Property Let DefaultColour18(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(17).BackColor = NewColour
    If varSelectionPos = 17 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour18")
    
End Property

Public Property Get DefaultColour19() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour19 = picDefaultColour(18).BackColor
    
End Property

Public Property Let DefaultColour19(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(18).BackColor = NewColour
    If varSelectionPos = 18 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour19")
    
End Property

Public Property Get DefaultColour20() As OLE_COLOR

    On Error Resume Next
    
    DefaultColour20 = picDefaultColour(19).BackColor
    
End Property

Public Property Let DefaultColour20(ByVal NewColour As OLE_COLOR)

    On Error Resume Next
    
    picDefaultColour(19).BackColor = NewColour
    If varSelectionPos = 19 Then shpSelection(2).FillColor = NewColour
    PropertyChanged ("DefaultColour20")
    
End Property

Public Sub Refresh()

    On Error Resume Next
    Dim Counter As Integer
    
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        picDefaultColour(Counter).Refresh
    Next Counter
    picCustomColour.Refresh
    lnSeperator(0).Refresh
    lnSeperator(1).Refresh
    cmdOther.Refresh
    For Counter = shpSelection.LBound To shpSelection.UBound
        shpSelection(Counter).Refresh
    Next Counter
    fraSelection.Refresh
    Me.Refresh
    
End Sub

Private Sub ColourClicked()

    On Error Resume Next
    
    'Raise the event with the colour of the selected picturebox's background/shpSelection(2)'s fill colour
    RaiseEvent ColourChoosen(shpSelection(2).FillColor)
    
End Sub

Public Property Get BackStyle() As BackStyleConsts

    On Error Resume Next
    
    BackStyle = UserControl.BackStyle
    
End Property

Public Property Let BackStyle(ByVal newBackStyle As BackStyleConsts)

    On Error Resume Next
    
    UserControl.BackStyle = newBackStyle
    
End Property

Public Property Get MousePointer() As MousePointerConstants

    On Error Resume Next
    
    MousePointer = picDefaultColour(0).MousePointer
    
End Property

Public Property Let MousePointer(ByVal NewPointer As MousePointerConstants)

    On Error Resume Next
    Dim Counter As Integer
    
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        picDefaultColour(Counter).MousePointer() = NewPointer
    Next Counter
    picCustomColour.MousePointer() = NewPointer
    fraSelection.MousePointer() = NewPointer
    PropertyChanged "MousePointer"
    
End Property

Public Property Get MouseIcon() As Picture
    
    On Error Resume Next
    
    Set MouseIcon = UserControl.MouseIcon
    
End Property

Public Property Set MouseIcon(ByVal NewMouseIcon As Picture)
    
    On Error Resume Next
    Dim Counter As Integer
    
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        Set picDefaultColour(Counter).MouseIcon = NewMouseIcon
    Next Counter
    Set picCustomColour.MouseIcon = NewMouseIcon
    Set fraSelection.MouseIcon = NewMouseIcon
    PropertyChanged "MouseIcon"
    
End Property


Public Property Get AllowCustomize() As Boolean

    On Error Resume Next
    
    AllowCustomize = varAllowCustomize
    
End Property

Public Property Let AllowCustomize(ByVal NewAllowCustomize As Boolean)

    On Error Resume Next
    
    varAllowCustomize = NewAllowCustomize
    PropertyChanged ("AllowCustomize")
    
End Property


Public Sub SaveColours(ByVal AppName As String, ByVal Section As String)

    On Error Resume Next
    Dim Counter As Integer
    
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        SaveSetting AppName, Section, "Colour" & Counter, picDefaultColour(Counter).BackColor
    Next Counter
    SaveSetting AppName, Section, "CustomColour", picCustomColour.BackColor
    
End Sub

Public Sub GetColours(ByVal AppName As String, ByVal Section As String)

    On Error Resume Next
    Dim Counter As Integer
    
    For Counter = picDefaultColour.LBound To picDefaultColour.UBound
        picDefaultColour(Counter).BackColor = GetSetting(AppName, Section, "Colour" & Counter, picDefaultColour(Counter).BackColor)
    Next Counter
    picCustomColour.BackColor = GetSetting(AppName, Section, "CustomColour", picCustomColour.BackColor)
    
End Sub

Public Sub RestoreColours()

    On Error Resume Next
    
    picDefaultColour(0).BackColor = &HFFFFFF
    picDefaultColour(1).BackColor = &H0&
    picDefaultColour(2).BackColor = &HC0C0C0
    picDefaultColour(3).BackColor = &H808080
    picDefaultColour(4).BackColor = &HFF&
    picDefaultColour(5).BackColor = &H80&
    picDefaultColour(6).BackColor = &HFFFF&
    picDefaultColour(7).BackColor = &H8080&
    picDefaultColour(8).BackColor = &HFF00&
    picDefaultColour(9).BackColor = &H8000&
    picDefaultColour(10).BackColor = &HFFFF00
    picDefaultColour(11).BackColor = &H808000
    picDefaultColour(12).BackColor = &HFF0000
    picDefaultColour(13).BackColor = &H800000
    picDefaultColour(14).BackColor = &HFF00FF
    picDefaultColour(15).BackColor = &H800080
    picDefaultColour(16).BackColor = &HC0DCC0
    picDefaultColour(17).BackColor = &HF0CAA6
    picDefaultColour(18).BackColor = &HF0FBFF
    picDefaultColour(19).BackColor = &HA4A0A0
    picCustomColour.BackColor = &HFFFFFF
    
End Sub


