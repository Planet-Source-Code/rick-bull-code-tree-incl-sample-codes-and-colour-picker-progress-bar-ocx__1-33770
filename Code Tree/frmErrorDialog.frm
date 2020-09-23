VERSION 5.00
Begin VB.Form frmErrorDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WINDOWTITLE"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
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
   ScaleHeight     =   2625
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLinkPos 
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   3000
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrLinkHot 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   3000
      Top             =   960
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details..."
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   980
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   550
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line lnSeperatror 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   120
      X2              =   6240
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line lnSeperatror 
      BorderColor     =   &H80000011&
      Index           =   0
      X1              =   120
      X2              =   6240
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label lblErrorDetails 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ERRORDETAILS"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   6060
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LINK"
      Height          =   195
      Left            =   3060
      MouseIcon       =   "frmErrorDialog.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmErrorDialog.frx":030A
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmErrorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#If Win32 Then
    Private Const OperatingSystem As String = "Microsoft Windows (32-Bit)"
#ElseIf Win16 Then
    Private Const OperatingSystem As String = "Microsoft Windows (16-Bit)"
#ElseIf Mac Then
    Private Const OperatingSystem As String = "Apple Mac OS"
#Else
    Private Const OperatingSystem As String = "Unknown"
#End If
Public Enum ErrorDialogFlags 'Flags for the ShowDialog routine - all in bit values (i.e. 0,1,2,4,8, etc)
    edfNone = 0 'No flags
    edfDisableOK = 1 'Disable the OK button
    edfDisableExit = 2 'Disable the Exit button
    edfDisableDetails = 4 'Disable the Details button
    edfIncludeReport = 8 'Send a reprot in HTML format when the user E-Mails you
    edfShowDetails = 16 'Show the details as though user had pressed the detail button
    edfSoundBeep = 32 'Beep when loading the form
    edfWriteErrorLog = 64 'Write an error log @ app.path\Error Log.txt
End Enum
Public Enum OutputModeConsts 'How to save files
    Add = 0 'Append to the existing data
    OverWrite = 1 'Overwrite all the existing data
End Enum
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Body As String
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
'Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
'Private Const BITSPIXEL = 12         '  Number of bits per pixel
'  Device Parameters for GetDeviceCaps()
Const DRIVERVERSION = 0      '  Device driver version
Const TECHNOLOGY = 2         '  Device classification
Const HORZSIZE = 4           '  Horizontal size in millimeters
Const VERTSIZE = 6           '  Vertical size in millimeters
Const HORZRES = 8            '  Horizontal width in pixels
Const VERTRES = 10           '  Vertical width in pixels
Const BITSPIXEL = 12         '  Number of bits per pixel
Const PLANES = 14            '  Number of planes
Const NUMBRUSHES = 16        '  Number of brushes the device has
Const NUMPENS = 18           '  Number of pens the device has
Const NUMMARKERS = 20        '  Number of markers the device has
Const NUMFONTS = 22          '  Number of fonts the device has
Const NUMCOLORS = 24         '  Number of colors the device supports
Const PDEVICESIZE = 26       '  Size required for device descriptor
Const CURVECAPS = 28         '  Curve capabilities
Const LINECAPS = 30          '  Line capabilities
Const POLYGONALCAPS = 32     '  Polygonal capabilities
Const TEXTCAPS = 34          '  Text capabilities
Const CLIPCAPS = 36          '  Clipping capabilities
Const RASTERCAPS = 38        '  Bitblt capabilities
Const ASPECTX = 40           '  Length of the X leg
Const ASPECTY = 42           '  Length of the Y leg
Const ASPECTXY = 44          '  Length of the hypotenuse

Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Const SIZEPALETTE = 104      '  Number of entries in physical palette
Const NUMRESERVED = 106      '  Number of reserved entries in palette
Const COLORRES = 108         '  Actual color resolution

'  Printing related DeviceCaps. These replace the appropriate Escapes
Const PHYSICALWIDTH = 110 '  Physical Width in device units
Const PHYSICALHEIGHT = 111 '  Physical Height in device units
Const PHYSICALOFFSETX = 112 '  Physical Printable Area x margin
Const PHYSICALOFFSETY = 113 '  Physical Printable Area y margin
Const SCALINGFACTORX = 114 '  Scaling factor x
Const SCALINGFACTORY = 115 '  Scaling factor y


Private Sub cmdDetails_Click()

    On Error Resume Next
    
    Me.Height = lblErrorDetails.Top + lblErrorDetails.Height + 500
    cmdDetails.Enabled = False
    
End Sub

Private Sub cmdExit_Click()

    On Error Resume Next
    
    End
    
End Sub

Private Sub cmdOK_Click()

    On Error Resume Next
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    'Clean up the used resources - REMEMBER to change _
     this if you change the name of the form!
    Set frmErrorDialog = Nothing
    
End Sub

Private Sub lblLink_Click()

    On Error Resume Next
    
    If lblLink.Tag <> "" And InStr(1, lblLink.Tag, "?body=") = 0 And InStr(1, lblLink.Tag, "mailto:") > 0 And Body <> "" Then
        Call ShellExecute(Me.hwnd, "", lblLink.Tag & Body, "", "", vbNormalFocus)
    ElseIf lblLink.Tag <> "" And (InStr(1, lblLink.Tag, "?body=") > 0 Or Body <> "") Then
        Call ShellExecute(Me.hwnd, "", lblLink.Tag, "", "", vbNormalFocus)
    End If
    
End Sub

Private Sub lblLink_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    If lblLink.ForeColor <> vbBlue And Button = 0 Then
        lblLink.ForeColor = vbBlue
        lblLink.FontUnderline = True
        fraLinkPos.Move lblLink.Left, lblLink.Top, lblLink.Width, lblLink.Height
        tmrLinkHot.Enabled = True
    End If
    
End Sub

Private Sub tmrLinkHot_Timer()

    On Error Resume Next
    Dim LinkPos As RECT
    Dim CursorPos As POINTAPI
    
    Call GetCursorPos(CursorPos)
    Call GetWindowRect(fraLinkPos.hwnd, LinkPos)
    
    If CursorPos.x < LinkPos.Left Or CursorPos.x > LinkPos.Right Or _
        CursorPos.y < LinkPos.Top Or CursorPos.y > LinkPos.Bottom Then
        lblLink.ForeColor = vbButtonText
        lblLink.FontUnderline = False
        fraLinkPos.Move lblLink.Left, lblLink.Top, lblLink.Width, lblLink.Height
        tmrLinkHot.Enabled = False
    End If
    
End Sub

Public Sub ShowDialog(ErrNumber As Long, Procedure As String, _
    Optional LinkAddress As String = "", Optional LinkCaption As String = "", _
    Optional LinkToolTipText As String = "", Optional Message As String = "", _
    Optional WindowTitle As String = "Alert", _
    Optional ByRef Flags As ErrorDialogFlags = edfNone)

    On Error Resume Next
    Dim ErrorDetails As String
    
    'Set the label captions
    Me.Caption = WindowTitle
    If Trim(Message) = "" Then
        lblMessage.Caption = "Sorry an error has occured." & vbNewLine & vbNewLine
        'If OK is not disabled
        If (Flags And edfDisableOK) = 0 Then _
            lblMessage.Caption = lblMessage.Caption & "To try to keep working press OK" & vbNewLine
        'If Exit is not disabled or OK is disabled (as we need to at least exit the form)
        If (Flags And edfDisableExit) = 0 Or (Flags And edfDisableOK) Then _
            lblMessage.Caption = lblMessage.Caption & "To exit the program press Exit" & vbNewLine
        'If Details is not disabled and they are not shown
        If (Flags And edfDisableDetails) = 0 And (Flags And edfShowDetails) = 0 Then _
            lblMessage.Caption = lblMessage.Caption & "For Details press Details"
    Else
        lblMessage.Caption = Message
    End If
    
    lblLink.Tag = LinkAddress
    lblLink.Caption = LinkCaption
    lblLink.ToolTipText = LinkToolTipText
    
    If lblLink.Caption = "" Then lblLink.Visible = False
    ErrorDetails = "Error " & ErrNumber & ": " & Error$(ErrNumber) & vbNewLine & _
        "Last DLL Error: " & GetLastError & vbNewLine & _
        "Error Source: " & App.EXEName & vbNewLine & _
        "Procedure: " & Procedure
    If Flags And edfIncludeReport Then
        Body = "?body=" & ErrorDetails
        Call SaveText(App.Path & "\Report.htm", GetReport(ErrNumber, Procedure))
    End If
    lblErrorDetails.Caption = ErrorDetails
    
    'Set the sizes
    If cmdDetails.Top + cmdDetails.Height >= lblMessage.Top + lblMessage.Height Then
        lblLink.Move (Me.Width / 2) - (lblLink.Width / 2), _
            cmdDetails.Top + cmdDetails.Height + 150
    Else
        lblLink.Move (Me.Width / 2) - (lblLink.Width / 2), _
            lblMessage.Top + lblMessage.Height + 150
    End If
    fraLinkPos.Move lblLink.Left, lblLink.Top, lblLink.Width, lblLink.Height
    lnSeperatror(0).Y1 = lblLink.Top + lblLink.Height + 150
    lnSeperatror(0).Y2 = lnSeperatror(0).Y1
    lnSeperatror(1).Y1 = lnSeperatror(0).Y1 + Screen.TwipsPerPixelY
    lnSeperatror(1).Y2 = lnSeperatror(1).Y1
    lnSeperatror(0).X1 = 100
    lnSeperatror(0).X2 = Me.Width - 215
    lnSeperatror(1).X1 = 100
    lnSeperatror(1).X2 = Me.Width - 200
    
    lblErrorDetails.Top = lnSeperatror(1).Y1 + 150
    
    If lblLink.Caption = "" And cmdDetails.Top + cmdDetails.Height >= lblMessage.Top + lblMessage.Height Then
        Me.Height = cmdDetails.Top + cmdDetails.Height + 500
    ElseIf lblLink.Caption = "" And cmdDetails.Top + cmdDetails.Height < lblMessage.Top + lblMessage.Height Then
        Me.Height = lblMessage.Top + lblMessage.Height + 500
    Else
        Me.Height = lnSeperatror(0).Y1 + 300
    End If
    
    'Set the flags
    If Flags And edfShowDetails Then Call cmdDetails_Click
    If Flags And edfDisableOK Then cmdOK.Enabled = False
    If (Flags And edfDisableExit) And (Flags And edfDisableOK) = 0 Then cmdExit.Enabled = False
    If Flags And edfDisableDetails Then cmdDetails.Enabled = False
    If Flags And edfSoundBeep Then Beep
    If Flags And edfWriteErrorLog Then Call SaveText(App.Path & "\Error Log.txt", _
        "An Error occured on " & Date & " at " & Time & ":" & vbNewLine & _
        vbNewLine & ErrorDetails & vbNewLine & vbNewLine & _
        "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-" _
        & vbNewLine, Add)
    
    'If Trim(LogPath) <> "" Then Call SaveText(LogPath, _
        "An Error occured on " & Date & " at " & Time & ":" & vbNewLine & _
        vbNewLine & ErrorDetails & vbNewLine & vbNewLine & _
        "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-" _
        & vbNewLine, Add)
    'Reset the flags otherwise they will stay as last time if called again
    Flags = edfNone
    
    'Show the form
    Me.Show vbModal

End Sub

Private Sub SaveText(Filename As String, TextToSave As String, _
    Optional OutputMode As OutputModeConsts = OverWrite)
    
    On Error GoTo ErrorHandler
    Dim FileNumber As Byte
    
    FileNumber = FreeFile
    'If the FileName is not ""
    If Filename <> "" Then
        If OutputMode = OverWrite Then
            'Open the for output
            Open Filename For Output As #FileNumber
                'Write the text to the file
                Print #FileNumber, TextToSave
            'Close the file
            Close #FileNumber
            
        ElseIf OutputMode = Add Then
            'Open the for output
            Open Filename For Append As #FileNumber
                'Write the text to the file
                Print #FileNumber, TextToSave
            'Close the file
            Close #FileNumber
        End If
    End If
    'Exit the sub so as not to cause an error
    Exit Sub

ErrorHandler:
    'Close the file in case it is open
    Close #FileNumber
End Sub

Private Function GetReport(ErrNumber As Long, Procedure As String) As String

    On Error GoTo ErrorHandler
    Dim SysInfo As SYSTEM_INFO 'Variable for holding the system info
    Dim SystemDirectory As String, TempDirectory As String, WindowsDirectory As String 'Variables for the directories
    Dim TempString As String * 256 'Temp string for holding returned dir in array format (C++ style)
    Dim ReturnValue As Long 'Returned value for Get*Directory - needed to convert back to normal string
    Dim WindowsInfo As OSVERSIONINFO 'Variable for holding Windows info
    Dim MemoryInfo As MEMORYSTATUS
    Dim OSRealName As String, ServicePack As String
    
    'Get the path of:
    'System Directory
    ReturnValue = GetSystemDirectory(TempString, Len(TempString))
    SystemDirectory = Left$(TempString, ReturnValue)
    'Temp Directory
    ReturnValue = GetTempPath(Len(TempString), TempString)
    TempDirectory = Left$(TempString, ReturnValue)
    'Windows Directory
    ReturnValue = GetWindowsDirectory(TempString, Len(TempString))
    WindowsDirectory = Left$(TempString, ReturnValue)
       
    'Get the memory info
    Call GlobalMemoryStatus(MemoryInfo)
    
    'Get the system info
    Call GetSystemInfo(SysInfo)
    
    'Get Windows Info
    WindowsInfo.dwOSVersionInfoSize = Len(WindowsInfo)
    'WindowsInfo.szCSDVersion = Space(128)
    ReturnValue = GetVersionEx(WindowsInfo)
    
    If Asc(Left$(WindowsInfo.szCSDVersion, 1)) = 0 Then
        ServicePack = "None"
    Else
        ReturnValue = InStr(WindowsInfo.szCSDVersion, Chr(0))
        ServicePack = Left$(WindowsInfo.szCSDVersion, ReturnValue - 1)
    End If
    
    Select Case WindowsInfo.dwPlatformId
            
        Case VER_PLATFORM_WIN32_NT
            OSRealName = IIf(WindowsInfo.dwMajorVersion < 5, _
                "Windows NT", "Windows 2000 or Newer")
            
        Case VER_PLATFORM_WIN32_WINDOWS
            Select Case WindowsInfo.dwMinorVersion
                Case Is < 90
                    OSRealName = IIf(WindowsInfo.dwBuildNumber > 0, _
                        "Windows 95", "Windows 98")
                Case Is >= 90
                    OSRealName = "Windows Me or Newer"
            End Select
                
        Case VER_PLATFORM_WIN32s
            OSRealName = "Windows 3.1"
            
        Case Else
            OSRealName = "Unknown"
                
    End Select
        
    
    'Return the report in HTML format:
    'Header
    GetReport = "<HTML>" & vbNewLine & "<HEAD>" & vbNewLine & "<TITLE>" & App.EXEName & _
        " Error Report</TITLE>" & vbNewLine & "<STYLE TYPE=" & Chr(34) & "text/css" & Chr(34) & ">" & vbNewLine & _
        "    P.title { font-family: Tahoma, Arial, MS Sans Serif; font-size: 16px; text-align: center; text-decoration: underline; font-weight: bold; };" & vbNewLine & _
        "    P.heading { font-family: Tahoma, Arial, MS Sans Serif; font-size: 14px; text-decoration: underline; };" & vbNewLine & _
        "    P.normal { font-family: Tahoma, Arial, MS Sans Serif; font-size: 12px; };" & vbNewLine & _
        "    P.note { font-family: Tahoma, Arial, MS Sans Serif; font-size: 10px; };" & vbNewLine & _
        "    blockquote { font-family: Tahoma, Arial, MS Sans Serif; font-size: 12px; font-wieght: italic; };" & vbNewLine & _
        "</STYLE>" & vbNewLine & "</HEAD>" & vbNewLine & "<BODY>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "note" & Chr(34) & " ALIGN=" & Chr(34) & "RIGHT" & Chr(34) & _
        ">Generated on " & Date & " @ " & Time & "</P>" & vbNewLine
    
    'Body - Error Info
    GetReport = GetReport & "<P CLASS=" & Chr(34) & "title" & Chr(34) & ">" & App.EXEName & _
        " v" & App.Major & "." & App.Minor & "." & App.Revision & "</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "heading" & Chr(34) & "><B>Error Info</B>:</P>" & vbNewLine & "<P CLASS=" & Chr(34) & "normal" & Chr(34) & "><B>Error</B> " & _
        ErrNumber & ": " & Error$(ErrNumber) & "<BR>" & vbNewLine & _
        "<B>Last DLL Error</B>: " & GetLastError & "<BR>" & vbNewLine & "<B>Error Source</B>: " & App.EXEName & _
        "<BR>" & vbNewLine & "<B>Procedure</B>: " & Procedure & "</P>" & vbNewLine & "<P></P>" & vbNewLine & "<HR>" & vbNewLine & "<P></P>" & vbNewLine
    
    'User's System Info
    GetReport = GetReport & "<P CLASS=" & Chr(34) & "title" & Chr(34) & ">User Info</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "heading" & Chr(34) & "><B>System</B>:</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & ">" & _
        "<B>Active Processor Mask</B>: " & SysInfo.dwActiveProcessorMask & "<BR>" & vbNewLine & _
        "<B>Allocation Granularity</B>: " & SysInfo.dwAllocationGranularity & "<BR>" & vbNewLine & _
        "<B>Number Of Processors</B>: " & SysInfo.dwNumberOrfProcessors & "<BR>" & vbNewLine & _
        "<B>OEM ID</B>: " & SysInfo.dwOemID & "<BR>" & vbNewLine & _
        "<B>Page Size</B>: " & SysInfo.dwPageSize & "<BR>" & vbNewLine & _
        "<B>Processor Type</B>: " & SysInfo.dwProcessorType & "<BR>" & vbNewLine & _
        "<B>Reserved</B>: " & SysInfo.dwReserved & "<BR>" & vbNewLine & _
        "<B>Maximum Application Address</B>: " & SysInfo.lpMaximumApplicationAddress & "<BR>" & vbNewLine & _
        "<B>Minimum Application Address</B>: " & SysInfo.lpMinimumApplicationAddress & "</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & "><B>Screen Resolution</B>: " & GetDeviceCaps(Me.hdc, HORZRES) & _
        " * " & GetDeviceCaps(Me.hdc, VERTRES) & " * " & GetDeviceCaps(Me.hdc, BITSPIXEL) & "</P>" & vbNewLine
        
    'Memory
    GetReport = GetReport & "<P CLASS=" & Chr(34) & "heading" & Chr(34) & "><B>Memory</B>:</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & "><B>Available Page File</B>: " & MemoryInfo.dwAvailPageFile & "<BR>" & vbNewLine & _
        "<B>Available Physical</B>: " & MemoryInfo.dwAvailPhys & "<BR>" & vbNewLine & _
        "<B>Available Virtual</B>: " & MemoryInfo.dwAvailVirtual & "<BR>" & vbNewLine & _
        "<B>Length</B>: " & MemoryInfo.dwLength & "<BR>" & vbNewLine & _
        "<B>Load</B>: " & MemoryInfo.dwMemoryLoad & "<BR>" & vbNewLine & _
        "<B>Total Page File</B>: " & MemoryInfo.dwTotalPageFile & "<BR>" & vbNewLine & _
        "<B>Total Physical</B>: " & MemoryInfo.dwTotalPhys & "<BR>" & vbNewLine & _
        "<B>Total Virtual</B>: " & MemoryInfo.dwTotalVirtual & "</P>" & vbNewLine
        
    'User's OS Info
    GetReport = GetReport & "<P CLASS=" & Chr(34) & "heading" & Chr(34) & "><B>Operating System</B>:</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & ">" & "<B>Manufacturer/OS Name</B>: " & OperatingSystem & "</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & ">" & "<B>Platform ID</B>: " & WindowsInfo.dwPlatformId & " (" & OSRealName & ")<BR>" & vbNewLine & _
        "<B>Major Version</B>: " & WindowsInfo.dwMajorVersion & "<BR>" & vbNewLine & _
        "<B>Minor Version</B>: " & WindowsInfo.dwMinorVersion & "<BR>" & vbNewLine & _
        "<B>Build Number</B>: " & (WindowsInfo.dwBuildNumber And &HFFFF&) & "<BR>" & vbNewLine & _
        "<B>Service Pack Version</B>: " & ServicePack & "</P>" & vbNewLine & _
        "<P CLASS=" & Chr(34) & "normal" & Chr(34) & "><B>Windows Directory</B>: " & WindowsDirectory & "<BR>" & vbNewLine & _
        "<B>System Directory</B>: " & SystemDirectory & "<BR>" & vbNewLine & _
        "<B>Temp Directory</B>: " & TempDirectory & "</P>" & vbNewLine & _
        "<HR>" & vbNewLine
    
    'Footer
    GetReport = GetReport & "</BODY>" & vbNewLine & "</HTML>"
    
    'Exit so as not to cause an unjustified error
    Exit Function
    
ErrorHandler:
    'Return an error string
    GetReport = "!!! An Error Occured while generating the report in the GetReport Function !!!"
End Function
