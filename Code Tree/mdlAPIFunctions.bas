Attribute VB_Name = "mdlAPIFunctions"
Option Explicit
'API Declarations
Private Declare Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
  As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
  ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
  As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
  As String, ByVal lpKeyName As String, ByVal lpString As String, _
  ByVal lpFileName As String) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long 'For opening files in their default app
Private Declare Function GetMenu Lib "user32" _
(ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" _
(ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As _
Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked _
As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" _
    (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long 'API for making a form on top or not

'API Constants
'On Top:
Private Const HWND_TOPMOST As Long = -1 'Constant for making a form stay on top
Private Const HWND_NOTOPMOST As Long = -2 'Constant for making a form not on top
Private Const SWP_NOMOVE As Long = &H2 'Flags for Always On Top
Private Const SWP_NOSIZE As Long = &H1
'Sound
'Private Const SND_LOOP = &H8         '  loop the sound until next sndPlaySound
'Private Const SND_ALIAS = &H10000     '  name is a WIN.INI [sounds] entry
'Private Const SND_ALIAS_ID = &H110000    '  name is a WIN.INI [sounds] entry identifier
'Private Const SND_ALIAS_START = 0  '  must be > 4096 to keep strings in same section of resource file
'Private Const SND_APPLICATION = &H80         '  look for application specific association
Private Const SND_ASYNC = &H1         '  play asynchronously
Private Const SND_FILENAME = &H20000     '  name is a file name
'Private Const SND_MEMORY = &H4         '  lpszSoundName points to a memory file
'Private Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
'Private Const SND_NOSTOP = &H10        '  don't stop any currently playing sound
'Private Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy
'Private Const SND_PURGE = &H40               '  purge non-static events for task
'Private Const SND_RESERVED = &HFF000000  '  In particular these flags are reserved
'Private Const SND_RESOURCE = &H40004     '  name is a resource name or atom
'Private Const SND_SYNC = &H0         '  play synchronously (default)
'Private Const SND_TYPE_MASK = &H170007
'Private Const SND_VALID = &H1F        '  valid flags          / ;Internal /
'Private Const SND_VALIDFLAGS = &H17201F    '  Set of valid flag bits.  Anything outside

'Font enumeration types
Public Const LF_FACESIZE = 32
Public Const LF_FULLFACESIZE = 64

Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte

        lfFaceName(LF_FACESIZE) As Byte
End Type

Type NEWTEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte

        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
        ntmFlags As Long
        ntmSizeEM As Long
        ntmCellHeight As Long
        ntmAveWidth As Long
End Type

' ntmFlags field flags
Public Const NTM_REGULAR = &H40&
Public Const NTM_BOLD = &H20&
Public Const NTM_ITALIC = &H1&

'  tmPitchAndFamily flags
Public Const TMPF_FIXED_PITCH = &H1

Public Const TMPF_VECTOR = &H2
Public Const TMPF_DEVICE = &H8
Public Const TMPF_TRUETYPE = &H4

Public Const ELF_VERSION = 0
Public Const ELF_CULTURE_LATIN = 0

'  EnumFonts Masks
Public Const RASTER_FONTTYPE = &H1
Public Const DEVICE_FONTTYPE = &H2
Public Const TRUETYPE_FONTTYPE = &H4

Declare Function EnumFontFamilies Lib "gdi32" Alias _
     "EnumFontFamiliesA" _
     (ByVal hDC As Long, ByVal lpszFamily As String, _
     ByVal lpEnumFontFamProc As Long, lParam As Any) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
     ByVal hDC As Long) As Long
Private Const EM_UNDO = &HC7
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const TaskBar As String = "Shell_traywnd"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Function Undo(RTFName As RichTextBox)

    On Error Resume Next
    
    Call SendMessage(RTFName.hwnd, EM_UNDO, 0, 0)

End Function

Private Function EnumFontFamProc(lpNLF As LOGFONT, lpNTM As NEWTEXTMETRIC, _
     ByVal FontType As Long, lParam As ListBox) As Long

    On Error Resume Next
    Dim FaceName As String
    Dim FullName As String
    
    FaceName = StrConv(lpNLF.lfFaceName, vbUnicode)
    lParam.AddItem Left$(FaceName, InStr(FaceName, vbNullChar) - 1)
    EnumFontFamProc = 1

End Function

Public Sub FillComboWithFonts(CB As ComboBox)

    On Error Resume Next
    Dim hDC As Long
    
    CB.Clear
    hDC = GetDC(CB.hwnd)
    EnumFontFamilies hDC, vbNullString, AddressOf EnumFontFamProc, CB
    ReleaseDC CB.hwnd, hDC
    
End Sub

Public Sub MenuBold(FormName As Form, MenuItemNo As Long, SubMenuItemNo As Long)
    
    On Error Resume Next
    Dim Menu As Long
    Dim SubMenu As Long
    
    Menu = GetMenu(FormName.hwnd)
    SubMenu = GetSubMenu(Menu, MenuItemNo)
    Call SetMenuDefaultItem(SubMenu, SubMenuItemNo, 1&)

End Sub

Public Sub MenuIcons(ByVal FormName As Form, MenuItemNo As Integer, _
    SubMenuItemNo As Integer, ByVal UnCheckedPicture, _
    Optional ByVal CheckedPicture)
    
    On Error Resume Next
    Dim Menu As Long
    Dim SubMenu As Long
    Dim MenuItemID As Long
    Dim ReturnValue As Long

    If CheckedPicture = vbNullString Then CheckedPicture = UnCheckedPicture

    Menu = GetMenu(FormName.hwnd) 'Get the Hwnd of form
    SubMenu = GetSubMenu(Menu, MenuItemNo) 'Find Menu Index MenuItemNo
    MenuItemID = GetMenuItemID(SubMenu, SubMenuItemNo) 'Find submenu Index SubMenuItemNo
    ReturnValue = SetMenuItemBitmaps(Menu, MenuItemID, 0, _
    UnCheckedPicture, CheckedPicture)

End Sub

Public Sub OnTop(ByVal FormName As Form, ByVal OnTop As Boolean)

    On Error Resume Next 'Goto next line on an error

    'If the form is wanted to be on top
    If OnTop = True Then
        'Set it on top
        SetWindowPos FormName.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'If the form isn't
    Else
        'Stop it always being on top
        SetWindowPos FormName.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
End Sub

Public Sub ErrorFunction(ByVal ErrorNumber As Long, ByVal ErrorDescription As String, _
    Optional ByVal IgnoreCancelError As Boolean = True, Optional ByVal WriteLog As Boolean = True, _
    Optional ByVal LogPath As String, Optional ByVal DisplayMsgBox As Boolean = True, _
    Optional ByVal EMailAddress As String, Optional ByVal OwnerHwnd As Long, _
    Optional ByVal Procedure As String, Optional ByVal Message As String = "")
    
    On Error Resume Next
    Dim ErrorLog As String
    Dim TheError As String
    Dim Msg As String
    Dim ShowButtons As Integer
    Dim Answer As Integer
    
    'User chose Cancel and cancel error is to be ignored or there is no error
    If (ErrorNumber = 32755 And IgnoreCancelError = True) Or ErrorNumber = 0 Then
        Exit Sub
    Else
        'If LogPath = "" Then LogPath = App.Path & "\ErrorLog.txt"
        'Create the error description in a variable
        If Procedure <> "" Then Procedure = Procedure & " generated the error: "
        TheError = Procedure & vbNewLine & ErrorNumber & ": " & ErrorDescription & vbNewLine
        Msg = Message
        If Msg = "" Then Msg = "Sorry, an error has occured:" & vbNewLine & vbNewLine & TheError
        If WriteLog = True Then
            ErrorLog = LogPath
            'Write the details to a file
            Open ErrorLog For Append As #1
                Print #1, "[" & Date & " - " & Time & "]" & vbNewLine & TheError
            Close #1
            Msg = Msg & vbNewLine & "The details have been recorded to the file '" & ErrorLog & "'."
        End If
        
        'Set which buttons should appear on the msgbox
        If EMailAddress = "" Then
            ShowButtons = vbOKOnly + vbExclamation
        Else
            ShowButtons = vbYesNo + vbDefaultButton2 + vbExclamation
            Msg = Msg & vbNewLine & "Would you like to E-mail the author with this error?"
        End If
        'If the use wants a message box to be displayed
        If DisplayMsgBox = True Then Answer = MsgBox(Msg, ShowButtons, "Error")
        If Answer = vbYes Then Call ShellExecute(OwnerHwnd, vbNullString, _
            "mailto:" & EMailAddress & "?body=" & App.EXEName & " (v" & App.Major & "." & App.Minor & "." & App.Revision & ")" _
            & ": " & TheError, vbNullString, "", vbNormalFocus)
    End If
    
End Sub

Public Sub PlaySound(ByVal Filename As String)
    
    On Error Resume Next 'Goto next line on an error
    
    'If the filename is not "" then play the sound
    If frmMain.varUseSound = True And Filename <> "" Then Call sndPlaySound(Filename, SND_FILENAME Or SND_ASYNC)
    
End Sub

Public Function GetINI(ByVal Filename As String, ByVal Section As String, _
    ByVal Key As String, Optional ByVal Default As String = "") As String
  
  On Error Resume Next
  Dim TempString As String * 256
  Dim ReturnValue As Long
  
  ReturnValue = GetPrivateProfileString(Section, Key, Default, _
    TempString, Len(TempString), Filename)
  GetINI = Left$(TempString, ReturnValue)
  
End Function

Public Sub SaveINI(ByVal Filename As String, ByVal Section As String, _
    ByVal Key As String, Optional ByVal Value As String = "")
  
  On Error Resume Next
  Call WritePrivateProfileString(Section, Key, Value, Filename)
    
End Sub
