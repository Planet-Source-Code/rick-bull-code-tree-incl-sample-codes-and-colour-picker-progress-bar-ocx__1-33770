VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOKCancelApply 
      Caption         =   "&Apply"
      Height          =   375
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdOKCancelApply 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdOKCancelApply 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame fraOptions 
      Height          =   3375
      Index           =   6
      Left            =   4920
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
      Begin VB.OptionButton optSettings 
         Caption         =   "INI File"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   51
         Top             =   1850
         Width           =   975
      End
      Begin VB.OptionButton optSettings 
         Caption         =   "Registry"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   50
         Top             =   1850
         Width           =   975
      End
      Begin VB.CheckBox chkAskQuitQuestion 
         Caption         =   "Ask for Confirmation on Exit"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   3000
         Width           =   2310
      End
      Begin VB.ComboBox cmbCopyFormat 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   2520
         Width           =   4095
      End
      Begin VB.TextBox txtDefaultLargePic 
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         Top             =   1080
         Width           =   2000
      End
      Begin VB.CheckBox chkSaveSettings 
         Caption         =   "Save Settings on Exit to:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1560
         Width           =   2175
      End
      Begin VB.CheckBox chkUseConfig 
         Caption         =   "Use configuration file"
         Height          =   255
         Left            =   2520
         TabIndex        =   25
         Top             =   480
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtLineSeperator 
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtDefaultPic 
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   2000
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "When Copying use:"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   47
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Default Large Picture Name:"
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   33
         Top             =   840
         Width           =   2025
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Line Seperator in Config.ini:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Default Small Picture Name:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1980
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   1215
      Index           =   2
      Left            =   4920
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkLoadDefault 
         Caption         =   "Load File when there is no code"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2600
      End
      Begin VB.TextBox txtDefaultFile 
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   4095
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Default File to Load:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   3735
      Index           =   7
      Left            =   4920
      TabIndex        =   35
      Top             =   3480
      Visible         =   0   'False
      Width           =   4575
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   3240
         Width           =   4095
      End
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   2640
         Width           =   4095
      End
      Begin VB.TextBox txtSound 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   1440
         Width           =   4095
      End
      Begin VB.CheckBox chkUseSound 
         Caption         =   "Use Sound"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Tree Click:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Tree Shrink:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Add Bookmark:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Copy Notes:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   3
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   39
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Tree Expand:"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   36
         Top             =   1200
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ilTabPics 
      Left            =   4560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483648
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":000C
            Key             =   "Tree"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0560
            Key             =   "Code Window"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":08B4
            Key             =   "Notes"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0C08
            Key             =   "Bookmarks"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":0F5C
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":12B0
            Key             =   "General"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1604
            Key             =   "Sounds"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOptions 
      Height          =   1935
      Index           =   3
      Left            =   120
      TabIndex        =   30
      Top             =   4920
      Visible         =   0   'False
      Width           =   4575
      Begin RichTextLib.RichTextBox rtfSeperator 
         Height          =   1215
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2143
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"frmOptions.frx":1B58
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Seperator When Copying Notes:"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   31
         Top             =   240
         Width           =   2340
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   1815
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   2400
      Width           =   4575
      Begin VB.ComboBox cmbStyle 
         Height          =   315
         ItemData        =   "frmOptions.frx":1BCB
         Left            =   240
         List            =   "frmOptions.frx":1BCD
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1320
         Width           =   3975
      End
      Begin VB.CheckBox chkHotTracking 
         Caption         =   "Use Hot Tracking"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox chkOneCode 
         Caption         =   "Allow more than one code to be viewed at the same time"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Style:"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   420
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   1575
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Top             =   6840
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CheckBox chkAutoSwitch 
         Caption         =   "Auto-Switch to Code When Bookmark is Double-Clicked"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   4250
      End
      Begin VB.CheckBox chkSaveBookmarks 
         Caption         =   "Save Bookmarks on Exit/Load Bookmarks at Start"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtBookmarksFile 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Saved Bookmarks File:"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1605
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   1815
      Index           =   5
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
      Begin MSComctlLib.Slider sldAniSpeed 
         Height          =   510
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   900
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         SelStart        =   100
         TickStyle       =   1
         TickFrequency   =   5
         Value           =   100
      End
      Begin VB.CheckBox chkMenuIcons 
         Caption         =   "Show Icons in menus"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   500
         Width           =   1815
      End
      Begin VB.CheckBox chkShowDragLine 
         Caption         =   "Show Drag-Line when Resizing windows"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3160
      End
      Begin VB.Label lblOptName 
         AutoSize        =   -1  'True
         Caption         =   "Tree Hide/Show Animation Speed:"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   2445
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4215
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   7435
      HotTracking     =   -1  'True
      ImageList       =   "ilTabPics"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   7
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tree"
            Key             =   "Tree"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Code Window"
            Key             =   "CodeWindow"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Notes"
            Key             =   "Notes"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bookmarks"
            Key             =   "Bookmarks"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "View"
            Key             =   "View"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
            ImageIndex      =   6
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sounds"
            Key             =   "Sounds"
            ImageVarType    =   2
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LastOn As Integer

Private Sub chkSaveSettings_Click()

    If chkSaveSettings.Value = 0 Then
        optSettings(0).Enabled = False
        optSettings(1).Enabled = False
    Else
        optSettings(0).Enabled = True
        optSettings(1).Enabled = True
    End If
        
End Sub

Private Sub chkUseSound_Click()

    Dim Counter As Integer
    Dim TrueFalse As Boolean
    
    If chkUseSound.Value = 0 Then
        TrueFalse = False
    Else
        TrueFalse = True
    End If
    For Counter = lblSound.LBound To lblSound.UBound
        lblSound(Counter).Enabled = TrueFalse
        txtSound(Counter).Enabled = TrueFalse
    Next Counter
    
End Sub

Private Sub cmdOKCancelApply_Click(Index As Integer)

    On Error Resume Next 'Goto next line on error

    'Find which button has been pressed
    Select Case Index
        'OK
        Case 0
            'Apply the options
            Call ApplyOptions
            'Close the form
            Unload Me
        
        'Cancel
        Case 1
            'Close the form
            Unload Me
            
        'Apply
        Case 2
            'Apply the options
            Call ApplyOptions
            
    End Select
    
End Sub

Private Sub GetOptions()

    On Error Resume Next 'Goto next line on error

    'Ask Quit Question
    If frmMain.varAskQuitQuestion = False Then
        chkAskQuitQuestion.Value = 0
    Else
        chkAskQuitQuestion.Value = 1
    End If
    
    'Allow more than one code at a time
    If frmMain.tvCodeType.SingleSel = False Then
        chkOneCode.Value = 1
    Else
        chkOneCode.Value = 0
    End If
    
    'Hot tracking
    If frmMain.tvCodeType.HotTracking = False Then
        chkHotTracking.Value = 0
    Else
        chkHotTracking.Value = 1
    End If
    
    'Combo
    cmbStyle.AddItem "0 -  Text Only", 0
    cmbStyle.AddItem "1 -  Pictures & Text", 1
    cmbStyle.AddItem "2 -  Plus/Minus Signs & Text", 2
    cmbStyle.AddItem "3 -  Plus/Minus Signs, Pictures & Text", 3
    cmbStyle.AddItem "4 -  Treelines & Text", 4
    cmbStyle.AddItem "5 -  Treelines, Pictures & Text ", 5
    cmbStyle.AddItem "6 -  Treelines, Plus/Minus Signs & Text", 6
    cmbStyle.AddItem "7 -  Treelines, Plus/Minus Signs, Pictures & Text", 7
    cmbStyle.Text = cmbStyle.List(frmMain.tvCodeType.Style)
    cmbCopyFormat.AddItem "Plain Text", 0
    cmbCopyFormat.AddItem "RTF Text", 1
    'Copy format
    If frmMain.varCopyFormat = vbCFRTF Then
        'RTF Constant
        cmbCopyFormat.ListIndex = 1
    Else
        'RTF Constant
        cmbCopyFormat.ListIndex = 0
    End If
    
    'Use config file
    If frmMain.varUseConfig = False Then
        chkUseConfig.Value = 0
    Else
        chkUseConfig.Value = 1
    End If
    
    'Default file
    txtDefaultFile.Text = frmMain.varDefaultFile
    
    'Load default file?
    If frmMain.varLoadDefaultFile = False Then
        chkLoadDefault.Value = 0
    Else
        chkLoadDefault.Value = 1
    End If
    
    'Auto swicth to code?
    If frmMain.varAutoSwitch = False Then
        chkAutoSwitch.Value = 0
    Else
        chkAutoSwitch.Value = 1
    End If
    
    'Line Seperator
    txtLineSeperator.Text = frmMain.varSeperatingChar
    
    'Default Picture Name
    txtDefaultPic.Text = frmMain.varDefaultPicture
    txtDefaultLargePic.Text = frmMain.varDefaultLargePicture
    
    'Saved Boomarks file
    txtBookmarksFile.Text = frmMain.varBookmarksFileName
    
    'Save Bookmarks on exit?
    If frmMain.varSaveBookmarks = False Then
        chkSaveBookmarks.Value = 0
    Else
        chkSaveBookmarks.Value = 1
    End If
    
    'Show drag line
    If frmMain.varShowDragLine = False Then
        chkShowDragLine.Value = 0
    Else
        chkShowDragLine.Value = 1
    End If
    
    'Show menu icons
    If frmMain.varAddMenuIcons = False Then
        chkMenuIcons.Value = 0
    Else
        chkMenuIcons.Value = 1
    End If
    
    'Animation Speed
    sldAniSpeed.Value = ((sldAniSpeed.Max - frmMain.tmrHideTree.Interval) + sldAniSpeed.Min) + 1
    
    'Save settings
    If frmMain.varSaveSettings = False Then
        chkSaveSettings.Value = 0
        optSettings(0).Enabled = False
        optSettings(1).Enabled = False
    Else
        chkSaveSettings.Value = 1
        optSettings(0).Enabled = True
        optSettings(1).Enabled = True
    End If
    
    'Save in reg
    optSettings(0).Value = frmMain.varSaveInReg
    optSettings(1).Value = Not frmMain.varSaveInReg
    
    'Seperator
    rtfSeperator.SelRTF = frmMain.varSeperator
    
    'Use sounds
    If frmMain.varUseSound = False Then
        chkUseSound.Value = 0
    Else
        chkUseSound.Value = 1
    End If
    
    'Sounds
    txtSound(0).Text = frmMain.varTreeClickSound
    txtSound(1).Text = frmMain.varTreeExpandSound
    txtSound(2).Text = frmMain.varTreeShrinkSound
    txtSound(3).Text = frmMain.varNotesSound
    txtSound(4).Text = frmMain.varBookmarksSound

End Sub

Private Sub ApplyOptions()

    On Error Resume Next 'Goto next line on error
    
    'Ask Quit Question
    If chkAskQuitQuestion.Value = 0 Then
        frmMain.varAskQuitQuestion = False
    Else
        frmMain.varAskQuitQuestion = True
    End If
    
    'Sounds
    frmMain.varTreeClickSound = txtSound(0).Text
    frmMain.varTreeExpandSound = txtSound(1).Text
    frmMain.varTreeShrinkSound = txtSound(2).Text
    frmMain.varNotesSound = txtSound(3).Text
    frmMain.varBookmarksSound = txtSound(4).Text
        
    'Allow more than one code at a time
    If chkOneCode.Value = 0 Then
        frmMain.tvCodeType.SingleSel = True
    Else
        frmMain.tvCodeType.SingleSel = False
    End If
    
    'Hot tracking
    If chkHotTracking.Value = 0 Then
        frmMain.tvCodeType.HotTracking = False
    Else
        frmMain.tvCodeType.HotTracking = True
    End If
    
    'Tree style
    frmMain.tvCodeType.Style = cmbStyle.ListIndex
    
    'Copy format
    If cmbCopyFormat.ListIndex = 1 Then
        'RTF Constant
        frmMain.varCopyFormat = vbCFRTF
    Else
        'Text Constant
        frmMain.varCopyFormat = 1
    End If
    
    'Default file

    frmMain.varDefaultFile = txtDefaultFile.Text
    
    'Load default file?
    If chkLoadDefault.Value = 0 Then
        frmMain.varLoadDefaultFile = False
    Else
        frmMain.varLoadDefaultFile = True
    End If
    
    'Auto swicth to code?
    If chkAutoSwitch.Value = 0 Then
        frmMain.varAutoSwitch = False
    Else
        frmMain.varAutoSwitch = True
    End If
    
    'Line Seperator
    frmMain.varSeperatingChar = txtLineSeperator.Text
    
    'Default Picture Name
    frmMain.varDefaultPicture = txtDefaultPic.Text
    frmMain.varDefaultLargePicture = txtDefaultLargePic.Text
    
    'Saved Boomarks file
    frmMain.varBookmarksFileName = txtBookmarksFile.Text
    
    'Save Bookmarks on exit?
    If chkSaveBookmarks.Value = 0 Then
        frmMain.varSaveBookmarks = False
    Else
        frmMain.varSaveBookmarks = True
    End If
    
    'Show drag line
    If chkShowDragLine.Value = 0 Then
        frmMain.varShowDragLine = False
    Else
        frmMain.varShowDragLine = True
    End If
    
    'Show menu icons
    If chkMenuIcons.Value = 0 Then
        frmMain.varAddMenuIcons = False
        Call frmMain.RemoveMenuIcons
    Else
        frmMain.varAddMenuIcons = True
        Call frmMain.AddMenuIcons
    End If
    
    'Animation Speed
    frmMain.tmrHideTree.Interval = ((sldAniSpeed.Max - sldAniSpeed.Value) + sldAniSpeed.Min) + 1
    frmMain.tmrShowTree.Interval = frmMain.tmrHideTree.Interval
    
    'Use config file
    If chkUseConfig.Value = 0 Then
        frmMain.varUseConfig = False
    Else
        frmMain.varUseConfig = True
    End If
    
    'Save settings
    If chkSaveSettings.Value = 0 Then
        frmMain.varSaveSettings = False
    Else
        frmMain.varSaveSettings = True
    End If
    
    'Save in reg
    frmMain.varSaveInReg = optSettings(0).Value
    
    'Seperator
    frmMain.varSeperator = rtfSeperator.SelRTF
    
    'Use sounds
    If chkUseSound.Value = 0 Then
        frmMain.varUseSound = False
    Else
        frmMain.varUseSound = True
    End If
    
End Sub

Private Sub Form_Load()

    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer
    
    'Get the options
    Call GetOptions
    
    For Counter = fraOptions.LBound To fraOptions.UBound
        fraOptions(Counter).Move 240, 480
    Next Counter
    For Counter = lblSound.LBound To lblSound.UBound
        lblSound(Counter).MouseIcon = frmMain.imgHandCursor.Picture
    Next Counter
    LastOn = 1
    Me.Width = tbsOptions.Width + (tbsOptions.Left * 3)
    Me.Height = cmdOKCancelApply(0).Top + cmdOKCancelApply(0).Height + 500
    
    If frmMain.MnuTools_OnTop.Checked = True Then Call OnTop(Me, True)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmOptions = Nothing
    
End Sub

Private Sub lblSound_Click(Index As Integer)

    On Error GoTo ErrorHandler
    
    With frmMain.CD1
        .CancelError = True
        .DialogTitle = "Find Sound File"
        .Filter = "Wave Files (*.wav)|*.wav"
        .ShowOpen
        txtSound(Index).Text = .Filename
    End With
    Exit Sub
    
ErrorHandler:
    If Err.Number <> 32775 Then _
        Call frmErrorDialog.ShowDialog(Err.Number, "frmOptions.lblSound_Click", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
    
End Sub

Private Sub tbsOptions_Click()

    fraOptions(LastOn).Visible = False
    fraOptions(tbsOptions.SelectedItem.Index).Visible = True
    LastOn = tbsOptions.SelectedItem.Index
    
End Sub
