VERSION 5.00
Begin VB.UserControl Link 
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2070
   ScaleHeight     =   210
   ScaleWidth      =   2070
   ToolboxBitmap   =   "ctrlLink.ctx":0000
   Begin VB.Timer tmrMousePos 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblLinkText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.rickmusic.co.uk/"
      Height          =   195
      Left            =   0
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Link"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'Declare all variables

'Private variables for the apperance of the label
Private varCaption As String
Private varHotBackColour As Long
Private varHotForeColour As Long
Private varHotTracking As Boolean
Private varLinkAddress As String
Private varNormalBackColour As Long
Private varNormalFont As Font
Private varNormalForeColour As Long

'Default font specs variables
Private DefaultNormalFont As Font
    
'API Declarations
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long 'For opening files in their default app
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long 'Find where the cursor is
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
    lpRect As RECT) As Long 'Find the top, left, rottom & right of a window/control

'Type declarations
Private Type RECT 'The rectangle of a window/control
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI 'The X & Y co-ords of the cursor
        X As Long
        Y As Long
End Type

'Events for the control
Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)

'-=-=-=-=-=-=-=-=-=-=- Start of Properties -=-=-=-=-=-=-=-=-=-=-'
Public Property Get Caption() As String

    'Return the caption from the variable
    Caption = varCaption

End Property

Public Property Let Caption(ByVal NewCaption As String)

    'Set the caption into the variable
    varCaption = NewCaption
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get Enabled() As Boolean

    'Return the hot backcolour from the variable
    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal NewEnabled As Boolean)

    'Set the hot backcolour into the variable
    UserControl.Enabled = NewEnabled
    'Apply the changes
    'Call SetApperance
    
End Property

Public Property Get HotBackColour() As OLE_COLOR

    'Return the hot backcolour from the variable
    HotBackColour = varHotBackColour
    
End Property

Public Property Let HotBackColour(ByVal NewColour As OLE_COLOR)

    'Set the hot backcolour into the variable
    varHotBackColour = NewColour
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get HotForeColour() As OLE_COLOR

    'Return the hot fore colour from the variable
    HotForeColour = varHotForeColour
    
End Property

Public Property Let HotForeColour(ByVal NewColour As OLE_COLOR)

    'Set the hot forecolour into the variable
    varHotForeColour = NewColour
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get HotTracking() As Boolean

    'Return the hot tracking (true or false) from the variable
    HotTracking = varHotTracking
       
End Property

Public Property Let HotTracking(ByVal NewHotTracking As Boolean)

    'Set the hot trackng (true or false) into the variable
    varHotTracking = NewHotTracking
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get LinkAddress() As String

    'Return the link address from the variable
    LinkAddress = varLinkAddress
    
End Property

Public Property Let LinkAddress(ByVal NewAddress As String)

    'Set the link address into the variable
    varLinkAddress = NewAddress
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get NormalBackColour() As OLE_COLOR

    'Return the normal font backcolour from the variable
    NormalBackColour = varNormalBackColour
    
End Property

Public Property Let NormalBackColour(ByVal NewColour As OLE_COLOR)

    'Set the normal backcolour into the variable
    varNormalBackColour = NewColour
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get Font() As Font

    'Return the normal font from the variable
    Set Font = varNormalFont
    
End Property

Public Property Set Font(ByVal NewNormalFont As Font)

    'Set the hot font into the variable
    Set varNormalFont = NewNormalFont
    'Return that the property has been changed
    PropertyChanged "NormalFont"
    'Apply the changes
    Call SetApperance
    
End Property

Public Property Get NormalForeColour() As OLE_COLOR

    'Return the normal font forecolour from the variable
    NormalForeColour = varNormalForeColour
    
End Property

Public Property Let NormalForeColour(ByVal NewColour As OLE_COLOR)

    'Set the normal fore colour into the variable
    varNormalForeColour = NewColour
    'Apply the changes
    Call SetApperance
    
End Property
'
'-=-=-=-=-=-=-=-=-=-=-=- End of Properties -=-=-=-=-=-=-=-=-=-=-=-'

Private Sub SetApperance() 'Change the look of the label when needed
    
    On Error Resume Next 'If there is an error goto next line
    
    lblLinkText.Caption = varCaption
    UserControl.BackColor = varNormalBackColour
    lblLinkText.ForeColor = varNormalForeColour
    UserControl.Width = lblLinkText.Width
    UserControl.Height = lblLinkText.Height

End Sub

Private Sub lblLinkText_Change()

    On Error Resume Next 'If there is an error goto next line
    
    'Set the look of the form (update the caption)
    Call SetApperance
    
End Sub

Private Sub lblLinkText_Click()

    On Error Resume Next 'If there is an error goto next line
    
    'Open the Address
    Call ShellExecute(UserControl.hwnd, vbNullString, varLinkAddress, "", "", vbNormalFocus)
    'Send back a click event
    RaiseEvent Click
    
End Sub

Private Sub lblLinkText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'If there is an error goto next line
    
    'Enable then timer if needed
    If tmrMousePos.Enabled = False And varHotTracking = True Then tmrMousePos.Enabled = True
    'Set the forecolour if needed
    If lblLinkText.ForeColor <> varHotForeColour And varHotTracking = True Then lblLinkText.ForeColor = varHotForeColour
    'Set the backcolour if needed
    If UserControl.BackColor <> varHotBackColour And varHotTracking = True Then UserControl.BackColor = varHotBackColour
    'set the underline
    If lblLinkText.FontUnderline = False Then lblLinkText.FontUnderline = True
    'Set the font if needed
    'If lblLinkText.Font <> varHotFont Then lblLinkText.Font = varHotFont
    
End Sub

Private Sub tmrMousePos_Timer()

    On Error Resume Next 'If there is an error goto next line
    
    Dim ControlPos As RECT 'For the control's top, left, bottom & right
    Dim CursorPos As POINTAPI 'For the cursor's X & Y
    
    Call GetCursorPos(CursorPos) 'Find the cursor's X & Y
    Call GetWindowRect(UserControl.hwnd, ControlPos) 'Find the control's Positions
    
    'If the cursor is off of the label
    If CursorPos.X < ControlPos.Left Or CursorPos.X > ControlPos.Right Or _
        CursorPos.Y < ControlPos.Top Or CursorPos.Y > ControlPos.Bottom Then
            'Set the forecolour if needed
            If lblLinkText.ForeColor <> varNormalForeColour Then lblLinkText.ForeColor = varNormalForeColour
            'Set the backcolour if needed
            If UserControl.BackColor <> varNormalBackColour Then UserControl.BackColor = varNormalBackColour
            'set the underline
            If lblLinkText.FontUnderline = True Then lblLinkText.FontUnderline = False
            'Disable then timer if needed
            If tmrMousePos.Enabled = True Then tmrMousePos.Enabled = False
    End If
    
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    
    On Error Resume Next 'If there is an error goto next line
    
    'Send Click event back if the default/cancel is pressed (and needed)
    RaiseEvent Click
    
End Sub

Private Sub UserControl_InitProperties()
    
    On Error Resume Next 'If there is an error goto next line
    
    Caption = "http://www.rickmusic.co.uk"
    HotBackColour = &H8000000F
    Set Font = DefaultNormalFont
    HotForeColour = vbBlue
    HotTracking = True
    LinkAddress = "http://www.rickmusic.co.uk"
    NormalBackColour = &H8000000F
    NormalForeColour = &H80000012
    BackStyle = vbSolid

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    On Error Resume Next 'If there is an error goto next line
    
    'Get all of the properties
    Caption = PropBag.ReadProperty("Caption", "http://www.rickmusic.co.uk")
    Enabled = PropBag.ReadProperty("Enabled", True)
    HotBackColour = PropBag.ReadProperty("HotBackColour", &H8000000F)
    HotForeColour = PropBag.ReadProperty("HotForeColour", vbBlue)
    HotTracking = PropBag.ReadProperty("HotTracking", True)
    LinkAddress = PropBag.ReadProperty("LinkAddress", "http://www.rickmusic.co.uk")
    NormalBackColour = PropBag.ReadProperty("NormalBackColour", &H8000000F)
    Set Font = PropBag.ReadProperty("NormalFont", DefaultNormalFont)
    NormalForeColour = PropBag.ReadProperty("NormalForeColour", &H80000012)

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next 'If there is an error goto next line
    
    'Set the look of the control
    Call SetApperance
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next 'If there is an error goto next line
    
    'Write all of the properties
    Call PropBag.WriteProperty("BackStyle", 1)
    Call PropBag.WriteProperty("Caption", varCaption, "http://www.rickmusic.co.uk")
    Call PropBag.WriteProperty("Enabled", True)
    Call PropBag.WriteProperty("HotBackColour", varHotBackColour, &H8000000F)
    Call PropBag.WriteProperty("HotForeColour", varHotForeColour, vbBlue)
    Call PropBag.WriteProperty("HotTracking", varHotTracking, True)
    Call PropBag.WriteProperty("LinkAddress", varLinkAddress, "http://www.rickmusic.co.uk")
    Call PropBag.WriteProperty("NormalBackColour", varNormalBackColour, &H8000000F)
    Call PropBag.WriteProperty("NormalFont", varNormalFont, DefaultNormalFont)
    Call PropBag.WriteProperty("NormalForeColour", varNormalForeColour, &H80000012)
    
End Sub

