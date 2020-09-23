VERSION 5.00
Begin VB.UserControl ProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   DrawMode        =   10  'Mask Pen
   PropertyPages   =   "ProgressBar.ctx":0000
   ScaleHeight     =   1575
   ScaleWidth      =   7980
   ToolboxBitmap   =   "ProgressBar.ctx":0023
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Public Enum AppearanceConsts
    [Flat]
    [3D]
End Enum
Public Enum BorderStyleConsts
    [None]
    [Fixed Single]
End Enum
Public Enum CaptionStyleConsts
    [ShowPercentage]
    [ShowValue]
    [UserDefined]
End Enum
Public Enum OrientationConsts
    [Horizontal]
    [Vertical]
End Enum
Private varCaption As String
Private varCaptionAlignment As AlignmentConstants
Private varCaptionStyle As CaptionStyleConsts
Private varFillColor As OLE_COLOR
'Private varFillColor2 As OLE_COLOR
Private varMax As Integer
Private varMin As Integer
Private varPercentage As Integer
Private varOrientation As OrientationConsts
Private varShowCaption As Boolean
Private varValue As Integer
Public Event Changed(Value As Integer)
Attribute Changed.VB_Description = "Occurs when the value changes."
Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Public Event KeyDown(ByVal KeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants)
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyUp(ByVal KeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants)
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseDown(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(ByVal Button As MouseButtonConstants, ByVal Shift As ShiftConstants, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Public Event Paint()
Public Event Resize()

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    Call SetValue
    Call PropertyChanged("BackColor")
End Property

Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute FillColor.VB_UserMemId = -510
    FillColor = varFillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    varFillColor = New_FillColor
    Call SetValue
    Call PropertyChanged("FillColor")
End Property

'Public Property Get FillColor2() As OLE_COLOR
'    FillColor2 = varFillColor2
'End Property
'
'Public Property Let FillColor2(ByVal New_FillColor2 As OLE_COLOR)
'    varFillColor2 = New_FillColor2
'    Call SetValue
'    Call PropertyChanged("FillColor2")
'End Property

Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Call SetValue
    Call PropertyChanged("Font")
End Property

Public Property Get Max() As Integer
Attribute Max.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Max = varMax
End Property

Public Property Let Max(ByVal New_Max As Integer)
    varMax = New_Max
    Call SetValue
    Call PropertyChanged("Max")
End Property

Public Property Get Min() As Integer
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Min = varMin
End Property

Public Property Let Min(ByVal New_Min As Integer)
    varMin = New_Min
    Call SetValue
    Call PropertyChanged("Min")
End Property

Private Property Get Percentage() As Integer
    Percentage = varPercentage
End Property

Private Property Let Percentage(ByVal New_Percentage As Integer)
    varPercentage = New_Percentage
    Call PropertyChanged("Percentage")
End Property

Public Property Get Value() As Integer
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Value = varValue
End Property

Public Property Let Value(ByVal New_Value As Integer)
    If Enabled = False Then Exit Property
    If New_Value > Max Or New_Value < Min Then
        'Err.Raise 380, "Value", "Invalid property value"
        Call Error(380)
        Exit Property
    End If
    varValue = New_Value
    Call SetValue
    Call PropertyChanged("Value")
End Property

Private Sub SetValue()
    On Error Resume Next
    Dim Counter As Integer 'For loops
    
    'Set the percentage
    Percentage = ((Value - Min) / (Max - Min)) * 100
    If CaptionStyle = ShowPercentage Then
        varCaption = Percentage & "%"
    ElseIf CaptionStyle = ShowValue Then
        varCaption = Value
    End If
    
    With UserControl
        'Clear the usercontrol
        .Cls
        If ShowCaption = True Then
            'Set the position for the text
            .CurrentY = (.Height \ 2) - (.TextHeight(Caption) \ 2) - Screen.TwipsPerPixelY
            Select Case CaptionAlignment
                Case vbCenter
                    .CurrentX = (.Width \ 2) - (.TextWidth(Caption) \ 2) - Screen.TwipsPerPixelX
                    
                Case vbLeftJustify
                    .CurrentX = 6 * Screen.TwipsPerPixelX
                    
                Case vbRightJustify
                    .CurrentX = .Width - (.TextWidth(Caption) + 6 * Screen.TwipsPerPixelX)
                    
            End Select
            'Set the text
            UserControl.Print Caption
        End If
        Select Case Orientation
            Case Horizontal
                'Draw a Filled Box
                UserControl.Line (-Screen.TwipsPerPixelX, 0)- _
                    (((.Width / 100) * Percentage) - (2 * Screen.TwipsPerPixelX), .Height), _
                    FillColor, BF
            Case Vertical
                'Draw a Filled Box
                UserControl.Line (0, .Height)- _
                    (.Width, (.Height / 100) * (100 - Percentage)), _
                    FillColor, BF
        End Select
        'Refresh it
        If .AutoRedraw = True Then .Refresh
    End With
    RaiseEvent Changed(Value)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next

    BackColor = RGB(255, 255, 255)
    CaptionAlignment = vbCenter
    CaptionStyle = ShowPercentage
    FillColor = RGB(0, 0, 128)
    'FillColor2 As OLE_COLOR
    Max = 100
    Min = 0
    Percentage = 0
    Value = 0
    Orientation = Horizontal
    ShowCaption = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    RaiseEvent Paint
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next

    With PropBag
        Appearance = .ReadProperty("Appearance", [3D])
        BackColor = .ReadProperty("BackColor", RGB(255, 255, 255)) 'Ambient.BackColor)
        BorderStyle = .ReadProperty("BorderStyle", None)
        CaptionAlignment = .ReadProperty("CaptionAlignment", vbCenter)
        CaptionStyle = .ReadProperty("CaptionStyle", ShowPercentage)
        Caption = .ReadProperty("Caption", "")
        Enabled = .ReadProperty("Enabled", True)
        FillColor = .ReadProperty("FillColor", RGB(0, 0, 128)) 'vbHighlight)
        'FillColor2 = .ReadProperty("FillColor2", vbHighlight)
        Set Font = .ReadProperty("Font", Ambient.Font)
        Max = .ReadProperty("Max", 100)
        Min = .ReadProperty("Min", 0)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        MousePointer = .ReadProperty("MousePointer", vbDefault)
        Orientation = .ReadProperty("Orientation", Horizontal)
        Percentage = .ReadProperty("Percentage", 0)
        Set Picture = .ReadProperty("Picture", Nothing)
        ShowCaption = .ReadProperty("ShowCaption", True)
        Value = .ReadProperty("Value", 0)
    End With
End Sub

Private Sub UserControl_Resize()
    Call SetValue
    RaiseEvent Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    
    With PropBag
        Call .WriteProperty("Appearance", Appearance, [3D])
        Call .WriteProperty("BackColor", BackColor, RGB(255, 255, 255)) 'Ambient.BackColor)
        Call .WriteProperty("BorderStyle", BorderStyle, None)
        Call .WriteProperty("CaptionAlignment", CaptionAlignment, vbCenter)
        Call .WriteProperty("CaptionStyle", CaptionStyle, ShowPercentage)
        Call .WriteProperty("Caption", Caption, "")
        Call .WriteProperty("Enabled", Enabled, True)
        Call .WriteProperty("FillColor", FillColor, RGB(0, 0, 128)) 'vbHighlight)
        'Call .WriteProperty("FillColor2", FillColor2, vbHighlight)
        Call .WriteProperty("Font", Font, Ambient.Font)
        Call .WriteProperty("Max", Max, 100)
        Call .WriteProperty("Min", Min, 0)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", MousePointer, vbDefault)
        Call .WriteProperty("Orientation", Orientation, Horizontal)
        Call .WriteProperty("Percentage", Percentage, 0)
        Call .WriteProperty("Picture", Picture, Nothing)
        Call .WriteProperty("ShowCaption", ShowCaption, True)
        Call .WriteProperty("Value", Value, 0)
    End With
End Sub

Public Property Get Appearance() As AppearanceConsts
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConsts)
    UserControl.Appearance = New_Appearance
    Call SetValue
    Call PropertyChanged("Appearance")
End Property

Public Property Get BorderStyle() As BorderStyleConsts
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BorderStyle.VB_UserMemId = -504
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConsts)
    UserControl.BorderStyle = New_BorderStyle
    Call SetValue
    Call PropertyChanged("BorderStyle ")
End Property

Public Property Get Orientation() As OrientationConsts
Attribute Orientation.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Orientation = varOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As OrientationConsts)
    varOrientation = New_Orientation
    Call SetValue
    Call PropertyChanged("Orientation")
End Property

Public Property Get ShowCaption() As Boolean
Attribute ShowCaption.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowCaption = varShowCaption
End Property

Public Property Let ShowCaption(ByVal New_ShowCaption As Boolean)
    varShowCaption = New_ShowCaption
    Call SetValue
    Call PropertyChanged("ShowCaption")
End Property

Public Sub About()
Attribute About.VB_Description = "Shows the About dialog."
Attribute About.VB_UserMemId = -552
    On Error Resume Next
    
    Load frmAbout
    frmAbout.Show vbModal
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
    hWnd = UserControl.hWnd
End Property

Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property


Public Function GetPercentage() As Integer
    On Error Resume Next

    GetPercentage = Percentage
End Function

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    Call PropertyChanged("Enabled")
End Property

Public Property Get CaptionStyle() As CaptionStyleConsts
Attribute CaptionStyle.VB_Description = "Returns/sets the style of the caption (i.e. whether to automatically show the percentage, value or a custom caption)."
Attribute CaptionStyle.VB_ProcData.VB_Invoke_Property = ";Text"
    CaptionStyle = varCaptionStyle
End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As CaptionStyleConsts)
    varCaptionStyle = New_CaptionStyle
    Call SetValue
    Call PropertyChanged("CaptionStyle")
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Caption.VB_UserMemId = -518
    Caption = varCaption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    varCaption = New_Caption
    Call SetValue
    Call PropertyChanged("Caption")
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
Attribute CaptionAlignment.VB_Description = "Returns/sets the alignment of a CheckBox, or OptionButton, or a control's text."
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Text"
    CaptionAlignment = varCaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As AlignmentConstants)
    varCaptionAlignment = New_CaptionAlignment
    Call SetValue
    Call PropertyChanged("CaptionAlignment")
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer = New_MousePointer
    Call PropertyChanged("MousePointer")
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    Call PropertyChanged("MouseIcon")
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    Call SetValue
    Call PropertyChanged("Picture")
End Property
