VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Code Tree"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   10485
   Begin VB.PictureBox picDragLine 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   2320
      MouseIcon       =   "frmMain.frx":08CA
      MousePointer    =   99  'Custom
      ScaleHeight     =   4935
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   360
      Width           =   80
   End
   Begin MSComctlLib.ImageList ilRTFToolbar 
      Left            =   9720
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A1C
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B30
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C44
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D58
            Key             =   "Left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AC
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1400
            Key             =   "Right"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1754
            Key             =   "FontColour"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8
            Key             =   "BulList"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DFC
            Key             =   "NumList"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2150
            Key             =   "Symbol"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24A4
            Key             =   "ChangeCase"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbRTF 
      Height          =   330
      Left            =   2520
      TabIndex        =   11
      Top             =   720
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ilRTFToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Font"
            Style           =   4
            Object.Width           =   2500
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Font Size"
            Style           =   4
            Object.Width           =   550
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Description     =   "Bold"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Description     =   "Italic"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Description     =   "Underline"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Description     =   "Left"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Description     =   "Center"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Description     =   "Right"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font Colour"
            Description     =   "Font Colour"
            Style           =   4
            Object.Width           =   660
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BulList"
            Description     =   "Bullet List"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NumList"
            Description     =   "Number List"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ChangeCase"
            Description     =   "Change Case"
            Object.ToolTipText     =   "Change Case to Lower Case"
            ImageIndex      =   11
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "LowerCase"
                  Text            =   "Lower Case"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "UpperCase"
                  Text            =   "Upper Case"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SentenceCase"
                  Text            =   "Sentence Case"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ToggleCase"
                  Text            =   "Toggle Case"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TitleCase"
                  Text            =   "Title Case"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VaryCaseLower"
                  Text            =   "Vary Case - Lower First"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VaryCaseUpper"
                  Text            =   "Vary Case - Upper First"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Symbol"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin CodeTree.ColourButton ColourButton1 
         Height          =   315
         Left            =   5500
         TabIndex        =   14
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   556
         Colour          =   0
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   315
         Left            =   2330
         TabIndex        =   13
         Top             =   0
         Width           =   735
      End
      Begin VB.ComboBox cmbFontName 
         Height          =   315
         ItemData        =   "frmMain.frx":27F8
         Left            =   10
         List            =   "frmMain.frx":27FA
         TabIndex        =   12
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSComctlLib.ImageList ilMenus 
      Left            =   8520
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   12
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27FC
            Key             =   "About"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A3C
            Key             =   "Bookmark"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C7C
            Key             =   "ClearBookmarks"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EB0
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30E4
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3318
            Key             =   "CloseAll"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":354C
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3780
            Key             =   "Cross"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39B4
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BE8
            Key             =   "Rebuild"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E1C
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4050
            Key             =   "Expand"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4284
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44B8
            Key             =   "Goto"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":46EC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4920
            Key             =   "IconLarge"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B24
            Key             =   "IconList"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D28
            Key             =   "New"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F5C
            Key             =   "Notes"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5190
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53C4
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55F8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":582C
            Key             =   "PrintPreview"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A60
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C94
            Key             =   "ReadMe"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EC8
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":60FC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6330
            Key             =   "IconReport"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6534
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6768
            Key             =   "SaveAs"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":699C
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BD0
            Key             =   "SaveAll"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E04
            Key             =   "SelectAll"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7038
            Key             =   "Shrink"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":726C
            Key             =   "IconSmall"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7470
            Key             =   "SpellCheck"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76A4
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78D8
            Key             =   "Code"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B0C
            Key             =   "Error"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D40
            Key             =   "NewWin"
         EndProperty
      EndProperty
   End
   Begin VB.FileListBox fileList 
      Height          =   1065
      Left            =   8520
      Pattern         =   "*.rtf;*.txt"
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.DirListBox dirFoldersList 
      Height          =   1440
      Left            =   8520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ilTreeView 
      Left            =   8520
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   9120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8628
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8988
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8CE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":900C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":936C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":96CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A2D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A828
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B790
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BAE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilTabList 
      Left            =   9120
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BE38
            Key             =   "Code"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C18C
            Key             =   "Pad"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C4E0
            Key             =   "BookMark"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C834
            Key             =   "Descending"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB88
            Key             =   "Ascending"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8520
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrVisualSettings 
      Interval        =   25
      Left            =   9000
      Top             =   4320
   End
   Begin VB.Timer tmrHideTree 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   4320
   End
   Begin VB.Timer tmrShowTree 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   4320
   End
   Begin MSComctlLib.Toolbar tbStandard 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "New"
            Object.ToolTipText     =   "New (Ctrl + N)"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewCode"
                  Text            =   "Code"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewNotes"
                  Text            =   "Notes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open"
            Object.ToolTipText     =   "Open (Ctrl + O)"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenCode"
                  Text            =   "Code"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenNotes"
                  Text            =   "Notes"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Description     =   "Save"
            Object.ToolTipText     =   "Save (Ctrl + S)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAs"
            Description     =   "Save As"
            Object.ToolTipText     =   "Save As"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveAll"
            Description     =   "Save All"
            Object.ToolTipText     =   "Save All (Ctrl + L)"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PrintPreview"
            Description     =   "Print Preview"
            Object.ToolTipText     =   "Print Preview"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print (Ctrl + P)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SpellCheck"
            Description     =   "Spell Check"
            Object.ToolTipText     =   "Spell Check"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Description     =   "Find"
            Object.ToolTipText     =   "Find (Ctrl + F)"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut (Ctrl + X)"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy (Ctrl + C)"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste (Ctrl + V)"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Notes"
            Description     =   "Notes"
            Object.ToolTipText     =   "Copy Selected Text to your Notes"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bookmarks"
            Description     =   "Add To Bookmarks"
            Object.ToolTipText     =   "Add Current Code To Bookmarks (Ctrl + B)"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo (Ctrl + Z)"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Redo"
            Description     =   "Redo"
            Object.ToolTipText     =   "Redo (Ctrl + Y)"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Description     =   "Help"
            Object.ToolTipText     =   "Help (F1)"
            ImageIndex      =   17
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "back"
         Height          =   315
         Left            =   6840
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar sbSectionInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5370
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   2249
            MinWidth        =   2
            Text            =   "Language: None"
            TextSave        =   "Language: None"
            Key             =   "Language"
            Object.ToolTipText     =   "Shows which language you are in"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1931
            MinWidth        =   2
            Text            =   "Section: None"
            TextSave        =   "Section: None"
            Key             =   "Section"
            Object.ToolTipText     =   "Shows which section you are in"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Object.Width           =   1667
            MinWidth        =   2
            Text            =   "Code: None"
            TextSave        =   "Code: None"
            Key             =   "Code"
            Object.ToolTipText     =   "Shows which code you are viewing"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   7717
            MinWidth        =   2293
            Text            =   "Code Count: 0"
            TextSave        =   "Code Count: 0"
            Key             =   "CodeCount"
            Object.ToolTipText     =   "Shows the amount of code present"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
            Key             =   "Insert"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
            Key             =   "NumLock"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
            Key             =   "CapsLock"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "SCRL"
            Key             =   "ScrollLock"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvBookmarks 
      Height          =   4095
      Left            =   4320
      TabIndex        =   9
      Top             =   1080
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7223
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ilListLarge"
      SmallIcons      =   "ilTreeView"
      ColHdrIcons     =   "ilTabList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Language"
         Text            =   "Language"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Section"
         Text            =   "Section"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Code"
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfNotes 
      Height          =   4095
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":CEDC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfCode 
      Height          =   4095
      Left            =   2475
      TabIndex        =   7
      Top             =   1080
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"frmMain.frx":CF43
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsCodeTabs 
      Height          =   4935
      Left            =   2400
      TabIndex        =   6
      Top             =   360
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   8705
      HotTracking     =   -1  'True
      ImageList       =   "ilTabList"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Documentation"
            Key             =   "Code"
            Object.ToolTipText     =   "Shows the current code"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "My Notes"
            Key             =   "Notes"
            Object.ToolTipText     =   "Shows the notes that you have taken"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Bookmarks"
            Key             =   "Bookmarks"
            Object.ToolTipText     =   "Shows the codes that you have bookmarked"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvCodeType 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8705
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   618
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ilTreeView"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilListLarge 
      Left            =   9720
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin VB.Image imgHandCursor 
      Height          =   480
      Left            =   9120
      Picture         =   "frmMain.frx":CFAA
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBlankMenuIcon 
      Height          =   195
      Left            =   8640
      Top             =   4920
      Width           =   195
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile_Open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFile_SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFile_SaveAll 
         Caption         =   "Save A&ll"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuFile_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_PrintPreview 
         Caption         =   "Print Preview..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_NewWin 
         Caption         =   "New &Window"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit_Undo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEdit_Redo 
         Caption         =   "&Redo"
      End
      Begin VB.Menu mnuEdit_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEdit_Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEdit_Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEdit_Clear 
         Caption         =   "Cl&ear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_SelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Find 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit_Goto 
         Caption         =   "&Goto..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuView_Refresh 
         Caption         =   "&Refresh"
      End
   End
   Begin VB.Menu mnuBookmarks 
      Caption         =   "&Bookmarks"
      Begin VB.Menu mnuBookmarks_Add 
         Caption         =   "&Add Current Code"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBookmarks_Clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuBookmarks_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookmarks_View 
         Caption         =   "&Icon"
         Index           =   0
      End
      Begin VB.Menu mnuBookmarks_View 
         Caption         =   "&Small Icon"
         Index           =   1
      End
      Begin VB.Menu mnuBookmarks_View 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuBookmarks_View 
         Caption         =   "&Details"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuTree 
      Caption         =   "&Tree"
      Begin VB.Menu mnuTree_ReBuild 
         Caption         =   "Re Build..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTree_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTree_ExpandAll 
         Caption         =   "&Expand All"
      End
      Begin VB.Menu mnuTree_ShrinkAll 
         Caption         =   "&Shrink All"
      End
      Begin VB.Menu mnuTree_FindCodes 
         Caption         =   "Show All &Codes"
      End
      Begin VB.Menu mnuTree_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTree_Hide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuTree_Show 
         Caption         =   "&Show"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu MnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTools_CopyToNotes 
         Caption         =   "Copy To &Notes"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTools_SpellCheck 
         Caption         =   "&Spell Check"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MnuTools_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTools_Options 
         Caption         =   "&Options..."
      End
      Begin VB.Menu MnuTools_ErrorLog 
         Caption         =   "&View Error Log"
      End
      Begin VB.Menu MnuTools_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuTools_OnTop 
         Caption         =   "Always On Top"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_Help 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp_ReadMe 
         Caption         =   "&Read Me"
      End
      Begin VB.Menu mnuHelp_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuPop_Format 
         Caption         =   "&Format"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPop_Standard 
         Caption         =   "&Standard"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPop_Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_CustomizeStandard 
         Caption         =   "&Customize..."
      End
      Begin VB.Menu mnuPop_CustomizeFormat 
         Caption         =   "&Customize..."
      End
      Begin VB.Menu mnuPop_Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPop_Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPop_Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPop_SelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu mnuPop_Clear 
         Caption         =   "&Clear"
      End
      Begin VB.Menu mnuPop_Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPop_Notes 
         Caption         =   "Add To &Notes"
      End
      Begin VB.Menu mnuPop_Bookmark 
         Caption         =   "&Bookmark Current Code"
      End
      Begin VB.Menu mnuPop_Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Refresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuPop_Seperator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPop_Hide 
         Caption         =   "&Hide Tree"
      End
      Begin VB.Menu mnuPop_Show 
         Caption         =   "&Show Tree"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Declare all variables
'Variable Declarations: Public
Public varDefaultFile As String 'The default file's FileName
Public varSeperatingChar As String 'The seperator in the ini file
Public varDefaultPicture As String 'The name of the Default bitmap for each node
Public varDefaultLargePicture As String 'The name of the Default Large bitmap for each node
Public varLoadDefaultFile As Boolean 'Whether to load the default file
Public varBookmarksFileName As String 'The bookmarks file
Public varSaveBookmarks As Boolean 'If the BMs are to be saved
Public varShowDragLine As Boolean 'Whether to show the drag line when resizing
Public varAddMenuIcons As Boolean 'Whether to add icons to the menus
Public varAutoSwitch As Boolean 'Whether to switch to code when bookmark is clicked
Public varUseConfig As Boolean 'whether to use the config file
Public varSaveSettings As Boolean 'Whether to use config file
Public varSaveInReg As Boolean  'Where to save
Public varSeperator As String 'The seperator when copying notes
Public varUseSound As Boolean 'Whether to use sound
Public varTreeClickSound As String
Public varTreeExpandSound As String
Public varTreeShrinkSound As String
Public varNotesSound As String
Public varBookmarksSound As String
Public varCopyFormat As String
Public varAskQuitQuestion As Boolean
'Private
Private varLanguageName() As String 'The name of the language
Private varMouseX As Single 'Where the mouse is on the drag line
Private varTreeLeft As Single 'The tree's left pos
Private varNotesFileName As String 'The Note's Filename
Private varCodeFileName As String 'The Code's Filename
Private varTreeHidden As Boolean 'Whether the tree is hidden
Private varCodeCount As Integer  'How many codes there are
Private varHistory(-10 To 0) As Integer 'The history in node numbers
Private varLastHistory As Integer
Private varChangeTo As CaseConsts
Private varFirstCharCase As UpperLowerCaseConsts
'API Declarations
Private Declare Function GetCursorPos Lib "user32" _
    (lpPoint As POINTAPI) As Long 'The API for finding where the cursor is
'Type declarations
Private Type POINTAPI 'The Type for holding the X & Y values of the cursor pos
        X As Long
        Y As Long
End Type
'Enumberations
Private Enum PopUpType 'The popupmenu type
    FormatToolbar = 0
    StandardToolbar = 1
    TreeView = 2
    Code = 3
    Notes = 4
    Bookmark = 5
    Form = 6
End Enum
Private Enum TabKeys 'The tabs key strings
    CodeWindow = 0
    NotesWindow = 1
    BookmarksWindow = 2
End Enum
Private Enum HideShowConsts 'Hide or show the mouse consts
    MouseHide = 0
    MouseShow = 1
End Enum

Private Sub cmbFontName_Click()

    On Error Resume Next
    
    Select Case tbsCodeTabs.SelectedItem.Key
        Case "Code"
            rtfCode.SelFontName = cmbFontName.Text
        Case "Notes"
            rtfNotes.SelFontName = cmbFontName.Text
    End Select
    
End Sub

Private Sub cmbFontSize_Click()

    On Error Resume Next
    
    If IsNumeric(cmbFontSize.Text) Then
        Select Case tbsCodeTabs.SelectedItem.Key
            Case "Code"
                rtfCode.SelFontSize = cmbFontSize.Text
            Case "Notes"
                rtfNotes.SelFontSize = cmbFontSize.Text
        End Select
    End If
    
End Sub

Private Sub ColourButton1_ButtonDown()

    On Error Resume Next
    
    Load frmColour
    frmColour.Show 0, Me
    
End Sub

Private Sub ColourButton1_ButtonUp()

    On Error Resume Next
    If tbsCodeTabs.SelectedItem.Key = "Code" Then
        rtfCode.SetFocus
    ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
        rtfNotes.SetFocus
    End If
    
End Sub

Private Sub Command1_Click()

    On Error Resume Next
    
    Call History(-1)
    
End Sub

Private Sub Form_Load()

    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer
    
    Load frmSplash
    frmSplash.Show 0, Me
    'Initialize the variables
    varMouseX = -1
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    'Set the tree left value
    varTreeLeft = varTreeLeft = picDragLine.Left
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    lvBookmarks.ColumnHeaders(1).Icon = 0
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    lvBookmarks.ColumnHeaders(2).Icon = 0
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    lvBookmarks.ColumnHeaders(3).Icon = 0
    'Get the registry settings
    Call GetRegSettings
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    'Make the code tree
    Call MakeTree
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    'Add the menu icons
    If varAddMenuIcons = True Then Call AddMenuIcons
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    Unload frmSplash
    Call FillComboWithFonts(cmbFontName)
    For Counter = 2 To 12
        cmbFontSize.AddItem Counter
    Next Counter
    For Counter = 14 To 28 Step 2
        cmbFontSize.AddItem Counter
    Next Counter
    For Counter = 36 To 156 Step 12
        cmbFontSize.AddItem Counter
    Next Counter
    
    frmSplash.pbProgress.Value = frmSplash.pbProgress.Value + 1
    varChangeTo = LowerCase
    varFirstCharCase = vbLowerCase
    Me.Show
    Call Resize
    
End Sub

Private Sub MakeTree()
    
    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer
    Dim Counter2 As Integer
    Dim TempString As String
    Dim FileNumber As Byte
    Dim LineStart As Integer 'Where the line starts in the ini file
    Dim LineEnd As Integer 'Where the ; is in the ini file
    Dim SeperatorPos As Integer 'Where the = is in the ini file
        
    'Load the default file if wanted
    If varLoadDefaultFile = True Then rtfCode.LoadFile varDefaultFile
    'Clear the tree of previous entries
    tvCodeType.Nodes.Clear
    'Clear the previous images otherwise they won't work if the tree is rebuilt
    ilTreeView.ListImages.Clear
    ilTreeView.ImageHeight = 16
    ilTreeView.ImageWidth = 16
    ilListLarge.ListImages.Clear
    ilListLarge.ImageHeight = 32
    ilListLarge.ImageWidth = 32
    'Initialize the code count
    varCodeCount = 0
    'Clear the bookmarks
    lvBookmarks.ListItems.Clear
    
    'If we are to use the config file
    If varUseConfig = True Then
        'Find a free file number
        FileNumber = FreeFile
        'Open the file for input
        Open App.Path & "\Config.ini" For Input As #FileNumber
            'Return the file's contents
            TempString = Input(LOF(FileNumber), FileNumber)
        'Close the file
        Close #FileNumber
        
        'Initialize the start of the line variable
        LineStart = 1
        'Initialize the counter variable
        Counter = 0
        'Do while there are more lines in the config file
        Do While LineStart > 0
            'Find the end of the line (i.e. the varSeperatingChar)
            LineEnd = InStr(LineStart, TempString, varSeperatingChar, vbTextCompare)
            'If there is a line of text
            If LineEnd > 0 Then
                'Resize the language name variable to the required size
                ReDim Preserve varLanguageName(Counter)
                'Get the language name
                varLanguageName(Counter) = Trim(Mid(TempString, LineStart, LineEnd - LineStart))
                'make the line start equal to the line end + 1 (i.e. the next line)
                LineStart = LineEnd + Len(vbNewLine) + Len(varSeperatingChar)
                'Increment the counter by one
                Counter = Counter + 1
            
            'If there are no more lines of text
            Else
                'Make linestart = 0
                LineStart = 0
            End If
        Loop
    
    'If we aren't to use the config file & we just want to use all folders
    Else
        dirFoldersList.Path = App.Path
        For Counter = 0 To dirFoldersList.ListCount - 1
            'Resize the language name variable to the required size
            ReDim Preserve varLanguageName(Counter)
            'Get the language name
            varLanguageName(Counter) = Right(dirFoldersList.List(Counter), Len(dirFoldersList.List(Counter)) - InStrRev(dirFoldersList.List(Counter), "\", , vbTextCompare))  ' Len(dirFoldersList.List(Counter)) - 2)
        Next Counter
    End If
    
    'Loop for all languages
    For Counter = LBound(varLanguageName) To UBound(varLanguageName)
        'Change the dir to that of the current language name (in the array)
        dirFoldersList.Path = App.Path & "\" & varLanguageName(Counter)
        'Add the image to the image list
        ilTreeView.ListImages.Add , varLanguageName(Counter), LoadPicture(App.Path & "\" & varLanguageName(Counter) & "\" & varDefaultPicture)
        ilListLarge.ListImages.Add , varLanguageName(Counter), LoadPicture(App.Path & "\" & varLanguageName(Counter) & "\" & varDefaultLargePicture)
        'Add a node
        tvCodeType.Nodes.Add , , varLanguageName(Counter), varLanguageName(Counter), varLanguageName(Counter)
        'Loop for all dirs in list
        For Counter2 = 0 To dirFoldersList.ListCount - 1
            'Find the \ from the path
            SeperatorPos = InStrRev(dirFoldersList.List(Counter2), "\", , vbTextCompare)
            'If it is found then make the text to pass = just the file name
            If SeperatorPos > 0 Then TempString = Right(dirFoldersList.List(Counter2), Len(dirFoldersList.List(Counter2)) - SeperatorPos)
            'Set a node for the code
            Call GetCodes(dirFoldersList.List(Counter2), varLanguageName(Counter), TempString)
        'Onto next dir
        Next Counter2
    'Onto next language
    Next Counter
    'Set the amount of codes
    sbSectionInfo.Panels("CodeCount").Text = "Code Count: " & varCodeCount + 1
    Call GetBookmarks
    
End Sub

Private Sub Resize()

    On Error Resume Next 'Goto next line on error
    Dim TreeWidth As Single
    
    If varTreeHidden = False Then 'mnuTree_Hide.Enabled = True Then
        TreeWidth = picDragLine.Left
    Else
        TreeWidth = tvCodeType.Width
    End If
    
    'Move the objects on the form to the correct propertions:
    'If the toolbar is visible
    If tbStandard.Visible = True Then
        'Tree View
        tvCodeType.Move varTreeLeft, tbStandard.Top + tbStandard.Height + 50, _
            TreeWidth, (sbSectionInfo.Top - sbSectionInfo.Height) - 170
    'If it isn't
    Else
        'Tree View
        tvCodeType.Move varTreeLeft, 50, picDragLine.Left, sbSectionInfo.Top - 60
    End If
    'Drag Line
    picDragLine.Move tvCodeType.Left + tvCodeType.Width, tvCodeType.Top, _
        picDragLine.Width, tvCodeType.Height
    'Tab Strip
    tbsCodeTabs.Move picDragLine.Left + picDragLine.Width, picDragLine.Top, _
        Me.Width - (picDragLine.Width + picDragLine.Left) - 170, picDragLine.Height
    'Toolbar
    tbRTF.Move tbsCodeTabs.Left + 75, tbsCodeTabs.Top + 390, _
        tbsCodeTabs.Width - 170, tbsCodeTabs.Height - 480
    'Code RTF Box
    If tbRTF.Visible = True Then
        rtfCode.Move tbsCodeTabs.Left + 75, tbRTF.Top + tbRTF.Height, _
            tbsCodeTabs.Width - 170, tbsCodeTabs.Height - (480 + tbRTF.Height)
    Else
        rtfCode.Move tbsCodeTabs.Left + 75, tbsCodeTabs.Top + 390, _
            tbsCodeTabs.Width - 170, tbsCodeTabs.Height - 480
    End If
    'Notes RTF Box
    rtfNotes.Move rtfCode.Left, rtfCode.Top, rtfCode.Width, rtfCode.Height
    'List view (Bookmarks)
    lvBookmarks.Move rtfCode.Left, rtfCode.Top, rtfCode.Width, rtfCode.Height
    'Set the back colour of the drag line
    If picDragLine.BackColor <> Me.BackColor Then picDragLine.BackColor = Me.BackColor

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    
    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Form)
    
End Sub

Private Sub Form_Paint()

    On Error Resume Next 'Goto next line on error
    
    If varMouseX = -1 Then Call Resize
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next 'Goto next line on error
    
    Call Resize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next 'Goto next line on error
    Dim Answer As VbMsgBoxResult
    
    If varAskQuitQuestion = True Then _
        Answer = MsgBox("Are you sure that you want to quit?", _
            vbQuestion + vbYesNo + vbDefaultButton2, "Quit")
    If Answer = vbYes Or varAskQuitQuestion = False Then
        'Save the registry settings
        Call SaveRegSettings
        SaveINISettings
        'Save the BMs
        Call SaveBookmarks
            
        'Get rid of used memoery resources
        Set frmMain = Nothing
        
        'End the program (in case other forms are still open)
        End
    Else
        Cancel = 1
    End If

End Sub

Private Sub lvBookmarks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    On Error Resume Next
    
    'Set the right icon, sorted and sort order
    If ColumnHeader.Icon = 0 Then
        ColumnHeader.Icon = "Descending"
        lvBookmarks.Sorted = True
        lvBookmarks.SortOrder = lvwAscending
    ElseIf ColumnHeader.Icon = "Descending" Then
        ColumnHeader.Icon = "Ascending"
        lvBookmarks.Sorted = True
        lvBookmarks.SortOrder = lvwDescending
    ElseIf ColumnHeader.Icon = "Ascending" Then
        ColumnHeader.Icon = 0
        lvBookmarks.Sorted = False
    End If
    
    'Set which column is to be sorted
    lvBookmarks.SortKey = ColumnHeader.Index - 1
    
    'Remove the icons from the other columns
    If ColumnHeader.Index = 1 Then
        lvBookmarks.ColumnHeaders(2).Icon = 0
        lvBookmarks.ColumnHeaders(3).Icon = 0
        
    ElseIf ColumnHeader.Index = 2 Then
        lvBookmarks.ColumnHeaders(1).Icon = 0
        lvBookmarks.ColumnHeaders(3).Icon = 0
        
    ElseIf ColumnHeader.Index = 3 Then
        lvBookmarks.ColumnHeaders(1).Icon = 0
        lvBookmarks.ColumnHeaders(2).Icon = 0
        
    End If
    
End Sub

Private Sub lvBookmarks_DblClick()
    
    On Error GoTo ErrorHandler
    Dim NodeNo As Integer 'variable to find the node's index number
    
    If lvBookmarks.ListItems.Count > 0 Then
        'Find the index from the selected item in the list view (minus the BM from the beginning)
        NodeNo = Right(lvBookmarks.SelectedItem.Key, Len(lvBookmarks.SelectedItem.Key) - 2)
        'Select it
        Call SelectNode(NodeNo)
        'Select the code window if wanted
        If varAutoSwitch = True Then tbsCodeTabs.Tabs("Code").Selected = True
    End If
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.lvBookmarks_DblClick", _
    "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
    edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub lvBookmarks_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    'Remove the selected one if delet is pressed
    If (KeyCode = vbKeyDelete Or KeyCode = vbKeyBack) And lvBookmarks.ListItems.Count > 0 Then _
        lvBookmarks.ListItems.Remove lvBookmarks.SelectedItem.Key

End Sub

Private Sub lvBookmarks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Bookmark)
    
End Sub

Private Sub mnuBookmarks_Add_Click()

    On Error GoTo ErrorHandler
    Dim IconKey As String
    
    'Play the sound
    Call PlaySound(Me.varBookmarksSound)
    
    If InStr(1, tvCodeType.SelectedItem.Key, "_Code", vbTextCompare) > 0 Then
        IconKey = tvCodeType.SelectedItem.Parent.Key
    Else
        IconKey = tvCodeType.SelectedItem.Key
    End If
    lvBookmarks.ListItems.Add , "BM" & tvCodeType.SelectedItem.Index, _
        Mid(sbSectionInfo.Panels("Language").Text, 11), IconKey, tvCodeType.SelectedItem.Image
    lvBookmarks.ListItems(lvBookmarks.ListItems.Count).ListSubItems.Add , , _
        Mid(sbSectionInfo.Panels("Section").Text, 10)
    lvBookmarks.ListItems(lvBookmarks.ListItems.Count).ListSubItems.Add , , _
        Mid(sbSectionInfo.Panels("Code").Text, 7)
    Exit Sub
    
ErrorHandler:
    If Err.Number = 35602 Then
        MsgBox "You cannot add the current code to your bookmarks as it already exists.", vbExclamation + vbOKOnly, "Bookmark already Exisits"
        lvBookmarks.ListItems("BM" & tvCodeType.SelectedItem.Index).Selected = True
    Else
        Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.mnuBookmarks_Add_Click", _
            "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
            edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
    End If
    
End Sub

Private Sub mnuBookmarks_Clear_Click()

    On Error Resume Next 'Goto next line on error
    Dim Answer As VbMsgBoxResult 'Msgbox answer
    
    'Make sure user wants to
    Answer = MsgBox("Are you sure that you want to clear all Bookmarks?", _
        vbYesNo + vbExclamation + vbDefaultButton2, "Clear All Bookmarks?")
    
    'Clear the bookmarks
    If Answer = vbYes Then lvBookmarks.ListItems.Clear
    
End Sub

Private Sub mnuBookmarks_Click()

    On Error Resume Next 'Goto next line on error

    If lvBookmarks.ListItems.Count > 0 Then
        mnuBookmarks_Clear.Enabled = True
    Else
        mnuBookmarks_Clear.Enabled = False
    End If
    
End Sub

Private Sub mnuBookmarks_View_Click(Index As Integer)

    On Error Resume Next
    Dim Counter As Integer 'For loops
    
    'Loop for all views
    For Counter = mnuBookmarks_View.LBound To mnuBookmarks_View.uBound
        'Uncheck it
        mnuBookmarks_View(Counter).Checked = False
    'Onto next view
    Next Counter
    'Check the seleced one
    mnuBookmarks_View(Index).Checked = True
    'Change the view
    lvBookmarks.View = Index
    
End Sub

Private Sub mnuEdit_Click()
    
    On Error Resume Next 'Goto next line on error

    'Set the cut, copy paste and clear menu items according to the tool bar buttons
    mnuEdit_Undo.Enabled = tbStandard.Buttons("Undo").Enabled
    mnuEdit_Redo.Enabled = tbStandard.Buttons("Redo").Enabled
    mnuEdit_Cut.Enabled = tbStandard.Buttons("Cut").Enabled
    mnuEdit_Copy.Enabled = tbStandard.Buttons("Copy").Enabled
    mnuEdit_Paste.Enabled = tbStandard.Buttons("Paste").Enabled
    mnuEdit_Clear.Enabled = tbStandard.Buttons("Cut").Enabled
    
End Sub

Private Sub mnuEdit_Copy_Click()

    On Error Resume Next 'Goto next line on error
    
    'Clear the clipboard
    Clipboard.Clear
    'Find which RTF has the focus
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            If varCopyFormat = 1 Then
                'Set the text to the Clipboard
                Clipboard.SetText rtfCode.SelText, varCopyFormat
            Else
                'Set the RTF text to the Clipboard
                Clipboard.SetText rtfCode.SelRTF, varCopyFormat
            End If
        
        'Notes
        Case "Notes"
            If varCopyFormat = 1 Then
                'Set the text to the Clipboard
                Clipboard.SetText rtfNotes.SelText, varCopyFormat
            Else
                'Set the RTF text to the Clipboard
                Clipboard.SetText rtfNotes.SelRTF, varCopyFormat
            End If
            
    End Select
    
End Sub

Private Sub mnuEdit_Cut_Click()

    On Error Resume Next 'Goto next line on error
    
    'Clear the clipboard
    Clipboard.Clear
    'Find which RTF has the focus
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            If varCopyFormat = 1 Then
                'Set the text to the Clipboard
                Clipboard.SetText rtfCode.SelText, varCopyFormat
            Else
                'Set the RTF text to the Clipboard
                Clipboard.SetText rtfCode.SelRTF, varCopyFormat
            End If
            'Remove the text from the rtf
            rtfCode.SelText = ""
        
        'Notes
        Case "Notes"
            If varCopyFormat = 1 Then
                'Set the text to the Clipboard
                Clipboard.SetText rtfNotes.SelText, varCopyFormat
            Else
                'Set the text to the Clipboard
                Clipboard.SetText rtfNotes.SelRTF, varCopyFormat
            End If
            'Remove the text from the rtf
            rtfNotes.SelText = ""
    End Select
    
End Sub

Private Sub mnuEdit_Clear_Click()

    On Error Resume Next 'Goto next line on error
    
    'Find which tab has focus
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'Delete the selected text
            rtfCode.SelText = ""
            
        'Notes
        Case "Notes"
            'Delete the selected text
            rtfNotes.SelText = ""
        
        'Bookmarks
        Case "Bookmarks"
            'Clear them all
            lvBookmarks.ListItems.Clear
    
    End Select
    
End Sub

Private Sub mnuEdit_Find_Click()

    On Error Resume Next 'Goto next line on error
    
    'Load and show the find form
    Load frmFind
    frmFind.Show 0, Me
    
End Sub

Private Sub mnuEdit_Goto_Click()
    
    On Error GoTo ErrorHandler
    Dim LineNo As Long 'What number line the wants to go to
    Dim CurrentLineNo As Long 'What line number the loop is on
    Dim Counter As Long 'Used for the loop
    
    CurrentLineNo = 1
    Counter = 0
    
    LineNo = InputBox("Please enter the line number that you would like to Goto:", _
        "Goto Line:", , Me.Left + 100, Me.Top + 1080)

    If LineNo <> Null Then
        Select Case tbsCodeTabs.SelectedItem.Key
            Case "Code"
                Do While CurrentLineNo <> LineNo
                    Counter = Counter + 1
                    CurrentLineNo = rtfCode.GetLineFromChar(Counter)
                Loop
                'Select the position
                rtfCode.SelStart = Counter + 1
                'Set thefocus
                rtfCode.SetFocus
            Case "Notes"
                Do While CurrentLineNo <> LineNo
                    Counter = Counter + 1
                    CurrentLineNo = rtfNotes.GetLineFromChar(Counter)
                Loop
                'Select the position
                rtfNotes.SelStart = Counter + 1
                'Set thefocus
                rtfNotes.SetFocus
        End Select
    End If
    Exit Sub

ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.mnuEdit_Goto_Click", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub mnuEdit_Paste_Click()

    On Error Resume Next 'Goto next line on error
    
    'Find which RTF has the focus
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'If the text is rtf
            If Clipboard.GetText(ccCFRTF) <> "" Then
                'Set the RTF text from the Clipboard
                rtfCode.SelRTF = Clipboard.GetText(ccCFRTF)
            'If it is normal text
            ElseIf Clipboard.GetText(ccCFText) <> "" Then
                'Set the text from the Clipboard
                rtfCode.SelText = Clipboard.GetText(ccCFText)
            End If
        
        'Notes
        Case "Notes"
            'If the text is rtf
            If Clipboard.GetText(ccCFRTF) <> "" Then
                'Set the RTF text from the Clipboard
                rtfNotes.SelRTF = Clipboard.GetText(ccCFRTF)
            'If it is normal text
            ElseIf Clipboard.GetText(ccCFText) <> "" Then
                'Set the text from the Clipboard
                rtfNotes.SelText = Clipboard.GetText(ccCFText)
            End If
            
    End Select
    
End Sub

Private Sub mnuEdit_Redo_Click()
    
    On Error Resume Next 'Goto next line on error


End Sub

Private Sub mnuEdit_SelectAll_Click()
    
    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer 'For Loops
    
    'Find which tab has focus
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'Set the start selection
            rtfCode.SelStart = 0
            'Select all text
            rtfCode.SelLength = Len(rtfCode.Text)
            rtfCode.SetFocus
            
        'Notes
        Case "Notes"
            'Set the start selection
            rtfNotes.SelStart = 0
            'Select all text
            rtfNotes.SelLength = Len(rtfNotes.Text)
            rtfNotes.SetFocus
        
        'Boomarks
        Case "Bookmarks"
            For Counter = 1 To lvBookmarks.ListItems.Count
                lvBookmarks.ListItems(Counter).Selected = True
                'Let form repaint
                DoEvents
            Next Counter
            lvBookmarks.SetFocus
    End Select
    
End Sub

Private Sub mnuEdit_Undo_Click()
    
    On Error Resume Next 'Goto next line on error

    Select Case tbsCodeTabs.SelectedItem.Key
        Case "Code"
            Call Undo(rtfCode)
        Case "Notes"
            Call Undo(rtfNotes)
    End Select
    
End Sub

Private Sub mnuFile_Click()

    On Error Resume Next
    
    mnuFile_Print.Enabled = tbStandard.Buttons("Print").Enabled
    
End Sub

Private Sub mnuFile_Exit_Click()

    On Error Resume Next 'Goto next line on error
    
    'Unload the form
    Unload Me
    
End Sub

Private Sub mnuFile_New_Click()
    
    On Error Resume Next 'Goto next line on error
    
    'Make the code's file name blank
    varCodeFileName = ""
    'Make the text blank
    rtfCode.Text = ""
    Call SelectNone
    
End Sub

Private Sub mnuFile_NewWin_Click()

    On Error Resume Next
    
    Call Shell(App.Path & "\" & App.EXEName, vbNormalFocus)
    
End Sub

Private Sub mnuFile_Open_Click()
    
    On Error GoTo ErrorHandler
    
    With CD1
        .CancelError = True
        .Filter = "Rich Text Format (*.rtf)|*.rtf|Text Files (*.txt)|*.txt|All Files|*.*"
        .ShowOpen
        rtfCode.LoadFile .Filename
        varCodeFileName = .Filename
    End With
    Call SelectNone
    Exit Sub

ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.mnuFile_Open_Click", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub mnuFile_Print_Click()

    On Error GoTo ErrorHandler
    'Dim Counter As Integer
    
    With CD1
        .CancelError = True
        'Set flags to NO PRINT TO FILE + RETURN THE PRINTER DC + NO PRINT FROM/TO SELECTION + NO COLATE
        .Flags = cdlPDHidePrintToFile Or cdlPDReturnDC Or cdlPDNoSelection Or cdlPDCollate
        'Show the printer dialog
        .ShowPrinter
        'Loop for the number of copies wanted
        'For Counter = 1 To .Copies
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                'Print the rtf code with the selected printer
                rtfCode.SelPrint .hdc
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                'Print the rtf code with the selected printer
                rtfNotes.SelPrint .hdc
            End If
        'Onto next copy
        'Next Counter
    End With
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.mnuFile_Print_Click", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub mnuFile_PrintPreview_Click()
    
    On Error Resume Next 'Goto next line on error


End Sub

Private Sub mnuFile_Save_Click()

    On Error Resume Next 'Goto next line on an error
    
    'Select which tab we are on
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'Save the code
            Call SaveFiles(CodeWindow, varCodeFileName)
        'Notes
        Case "Notes"
            'Save the notes
            Call SaveFiles(NotesWindow, varNotesFileName)
        'Bookmarks
        Case "Bookmarks"
            'Save the bookmarks
            Call SaveFiles(BookmarksWindow)
    End Select
    
End Sub

Private Sub mnuFile_SaveAll_Click()
    
    On Error Resume Next 'Goto next line on error

    'Save all windows
    Call SaveFiles(CodeWindow, varCodeFileName)
    Call SaveFiles(NotesWindow, varNotesFileName)
    Call SaveFiles(BookmarksWindow)
    
End Sub

Private Sub mnuFile_SaveAs_Click()
    
    On Error Resume Next 'Goto next line on error

    'Select which tab we are on
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'Save the code
            Call SaveFiles(CodeWindow)
        'Notes
        Case "Notes"
            'Save the notes
            Call SaveFiles(NotesWindow)
    
    End Select
    
End Sub

Private Sub mnuHelp_About_Click()
    
    On Error Resume Next 'Goto next line on error

    'Load and show the about screen
    Load frmAbout
    frmAbout.Show 1, Me
    
End Sub

Private Sub mnuHelp_Help_Click()

    On Error Resume Next 'Goto next line on error
    
    With CD1
        'Set the help file
        .HelpFile = App.Path & "\Help.hlp"
        'Show it
        .ShowHelp
    End With
    
End Sub

Private Sub mnuHelp_ReadMe_Click()
    
    On Error Resume Next 'Goto next line on error

    'Open the Address
    Call ShellExecute(Me.hwnd, "", App.Path & "\ReadMe.txt", "", "", vbNormalFocus)
    
End Sub

Private Sub mnuPop_Bookmark_Click()

    On Error Resume Next 'Goto next line on error
    
    'Add to the bookmarks
    Call mnuBookmarks_Add_Click
    
End Sub

Private Sub mnuPop_Clear_Click()

    On Error Resume Next 'Goto next line on error

    Call mnuEdit_Clear_Click
    
End Sub

Private Sub mnuPop_Copy_Click()

    On Error Resume Next 'Goto next line on error

    Call mnuEdit_Copy_Click
    
End Sub

Private Sub mnuPop_CustomizeFormat_Click()

    On Error Resume Next 'Goto next line on error

    'Show the customize dialog
    tbRTF.Customize

End Sub

Private Sub mnuPop_CustomizeStandard_Click()

    On Error Resume Next 'Goto next line on error

    'Show the customize dialog
    tbStandard.Customize
    
End Sub

Private Sub mnuPop_Cut_Click()

    On Error Resume Next 'Goto next line on error

    Call mnuEdit_Cut_Click
    
End Sub

Private Sub mnuPop_Delete_Click()

    On Error Resume Next
    Dim Counter As Integer
    
    'For Counter = 1 To lvBookmarks.ListItems.Count
        'Remove the selected one
        'If lvBookmarks.ListItems(Counter).Selected = True Then
            lvBookmarks.ListItems.Remove (lvBookmarks.SelectedItem) ' (Counter)
    'Next Counter

End Sub

Private Sub mnuPop_Format_Click()

    On Error Resume Next 'Goto next line on error
    
    'Invert the check
    mnuPop_Format.Checked = Not mnuPop_Format.Checked
    'Make it invisble/visible as needed
    If tbsCodeTabs.SelectedItem.Key <> "Bookmarks" Then _
    tbRTF.Visible = mnuPop_Format.Checked
    'Make sure propertions are correct
    Call Resize
    
End Sub

Private Sub mnuPop_Hide_Click()

    On Error Resume Next
    
    'Hide the tree
    Call mnuTree_Hide_Click
    
End Sub

Private Sub mnuPop_Notes_Click()

    On Error Resume Next
    
    'Copy the selected text
    Call mnuTools_CopyToNotes_Click
    
End Sub

Private Sub mnuPop_Paste_Click()

    On Error Resume Next 'Goto next line on error

    Call mnuEdit_Paste_Click
    
End Sub

Private Sub mnuPop_Refresh_Click()

    On Error Resume Next
    
    'Refresh the form
    Call mnuView_Refresh_Click
    
End Sub

Private Sub mnuPop_SelectAll_Click()

    On Error Resume Next 'Goto next line on error

    'Show the tree
    Call mnuEdit_SelectAll_Click
    

End Sub

Private Sub mnuPop_Show_Click()

    On Error Resume Next 'Goto next line on error

    'Show the tree
    Call mnuTree_Show_Click
    
End Sub

Private Sub mnuPop_Standard_Click()

    On Error Resume Next 'Goto next line on error
    
    'Invert the check
    mnuPop_Standard.Checked = Not mnuPop_Standard.Checked
    'Make it invisble/visible as needed
    tbStandard.Visible = mnuPop_Standard.Checked
    'Make sure propertions are correct
    Call Resize
    
End Sub

Private Sub mnuTools_Click()
    
    On Error Resume Next 'Goto next line on error

    'Set the copy to notes menu items according to the tool bar button
    mnuTools_CopyToNotes.Enabled = tbStandard.Buttons("Notes").Enabled
    
End Sub

Private Sub mnuTools_CopyToNotes_Click()

    On Error Resume Next 'Goto next line on error
    
    'Play the sound
    Call PlaySound(Me.varNotesSound)
    
    'Add the selected text to the notes
    'rtfNotes.SelStart = Len(rtfNotes.Text)
    rtfNotes.SelRTF = rtfCode.SelRTF
    rtfNotes.SelRTF = varSeperator
    
End Sub

Private Sub MnuTools_ErrorLog_Click()
    
    On Error Resume Next 'Goto next line on error

    'Open the Address
    Call ShellExecute(Me.hwnd, "", App.Path & "\Error Log.txt", "", "", vbNormalFocus)

End Sub

Private Sub MnuTools_OnTop_Click()

    On Error Resume Next
    
    'Invert the check
    MnuTools_OnTop.Checked = Not MnuTools_OnTop.Checked
    'Set the ontop value
    Call OnTop(Me, MnuTools_OnTop.Checked)

End Sub

Private Sub MnuTools_Options_Click()
    
    On Error Resume Next 'Goto next line on error
    
    'Load & show the options
    Load frmOptions
    frmOptions.Show 1, Me
    
End Sub

Private Sub mnuTools_SpellCheck_Click()

    On Error Resume Next 'Goto next line on error

End Sub

Private Sub mnuTree_ExpandAll_Click()

    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer 'For loops
    
    'Loop for all nodes
    For Counter = 1 To tvCodeType.Nodes.Count
        'Make it visible
        tvCodeType.Nodes(Counter).Expanded = True
        'Let form repaint
        DoEvents
    'On to next node
    Next Counter

End Sub

Private Sub mnuTree_FindCodes_Click()

    On Error Resume Next
    Dim Counter As Integer 'For loops
    
    'Loop for all nodes
    For Counter = 1 To tvCodeType.Nodes.Count
        'If _Code is in the node's key select it
        If InStr(1, tvCodeType.Nodes(Counter).Key, "_Code") > 0 Then _
            tvCodeType.Nodes(Counter).EnsureVisible
        'Let form repaint
        DoEvents
    'On to next node
    Next Counter
    
End Sub

Private Sub mnuTree_Hide_Click()

    On Error Resume Next 'Goto next line on error

    'Disable the hide buttons
    mnuTree_Hide.Enabled = False
    mnuPop_Hide.Enabled = False
    'Disable the drag line
    picDragLine.Enabled = False
    picDragLine.Visible = False
    'Set the tree hidden variable
    varTreeHidden = True
    'Hide the tree via timer
    tmrHideTree.Enabled = True
    
End Sub

Private Sub MnuTree_ReBuild_Click()

    On Error Resume Next 'Goto next line on error
    Dim Answer As VbMsgBoxResult
    
    Answer = MsgBox("Are you sure that you want to rebuild the tree, you will loose your current position & bookmarks?", _
        vbYesNo + vbQuestion + vbDefaultButton2, "Rebuild Tree?")
    
    If Answer = vbYes Then Call MakeTree
    
End Sub

Private Sub mnuTree_Show_Click()

    On Error Resume Next 'Goto next line on error

    'Disable the buttons
    mnuTree_Show.Enabled = False
    mnuPop_Show.Enabled = False
    'Show the tree via timer
    tmrShowTree.Enabled = True
    
End Sub

Private Sub mnuTree_ShrinkAll_Click()

    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer 'For loops
    
    'Loop for all nodes (backwards - for looks!)
    For Counter = tvCodeType.Nodes.Count To 1 Step -1
        'Make it visible
        tvCodeType.Nodes(Counter).Expanded = False
        'Let form repaint
        DoEvents
    'On to next node
    Next Counter

End Sub

Private Sub mnuView_Refresh_Click()

    On Error Resume Next
    
    'Give it all a quick refresh and make sure sizes are right
    tvCodeType.Refresh
    rtfCode.Refresh
    rtfNotes.Refresh
    lvBookmarks.Refresh
    tbsCodeTabs.Refresh
    tbStandard.Refresh
    sbSectionInfo.Refresh
    picDragLine.Refresh
    Call Resize
    Me.Refresh
    
End Sub

Private Sub picDragLine_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error
    
    If Button = vbLeftButton Then
        'Set the back colour
        If varShowDragLine = True Then picDragLine.BackColor = vbApplicationWorkspace
        'Set the X value into a variable so we can center the _
        drag line and stop the form repainting
        If varShowDragLine = True Then varMouseX = X
    
    'If the right button is presses show the popupmenu
    ElseIf Button = vbRightButton Then
        Call ShowPopUpMenu(TreeView)

    End If
    
End Sub

Private Sub picDragLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next 'Goto next line on error
    Dim CursorPos As POINTAPI 'Variable for curosr position
    Dim LinePos As Single
    
    If Button = vbLeftButton Then
        'Find the cursor position
        Call GetCursorPos(CursorPos)
        'Input the line's position to a variable (the maths are supposed to make _
        the line centered with the cursor, but they could be wrong!)
        LinePos = picDragLine.Left + (X / 2)
        'If the line position is greater than 1500 and less than the _
        form's width / 2 then
        If LinePos > 1500 And LinePos < Me.Width / 2 Then
            'Move the drag line
            picDragLine.Left = LinePos
            'If the linepos is less than 1500
        ElseIf LinePos < 1500 Then
            'Move the drag line to the lowest point. This is done so _
            that is the cursor is less than 1500, the line will _
            be at the lowest value
            picDragLine.Left = 1500
            'If the linepos is more than the form width / 2
        ElseIf LinePos > Me.Width / 2 Then
            'Move the drag line to the greatest point. This is done so _
            that is the cursor is more than form width / 2, the line will _
            be at the biggest value
            picDragLine.Left = Me.Width / 2
        End If
    End If

End Sub

Private Sub picDragLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error
    
    If Button = vbLeftButton Then
        'Make the mousex = -1 so as not to keep redrawing _
        the form due to the paint event
        varMouseX = -1
        'Resize the form
        Call Resize
    End If
    
End Sub

Private Sub rtfCode_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    Call SetToolBarButtons(rtfCode)
    
End Sub

Private Sub rtfCode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error
    
    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Code)

End Sub

Private Sub rtfCode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    Call SetToolBarButtons(rtfCode)
    
End Sub

Private Sub rtfCode_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next
    Dim Counter As Integer
    
    If Data.Files.Count > 1 Then
        Load frmMultiDoc
        For Counter = 1 To Data.Files.Count
            frmMultiDoc.cmbFiles.AddItem Data.Files(Counter)
        Next Counter
        frmMultiDoc.cmbFiles.Text = Data.Files(1)
        frmMultiDoc.Show 1, Me
    Else
        Call rtfCode.LoadFile(Data.Files(1))
    End If
    
End Sub

Private Sub rtfNotes_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    Call SetToolBarButtons(rtfNotes)
    
End Sub

Private Sub rtfNotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error
    
    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Notes)
    
End Sub

Private Sub rtfNotes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error Resume Next 'Goto next line on error

    Call SetToolBarButtons(rtfNotes)
    
End Sub

Private Sub rtfNotes_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    Dim Counter As Integer
    
    If Data.Files.Count > 1 Then
        Load frmMultiDoc
        For Counter = 1 To Data.Files.Count
            frmMultiDoc.cmbFiles.AddItem Data.Files(Counter)
        Next Counter
        frmMultiDoc.cmbFiles.Text = Data.Files(1)
        frmMultiDoc.Show 1, Me
    Else
        Call rtfNotes.LoadFile(Data.Files(1))
    End If
    
End Sub

Private Sub sbSectionInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Form)
    
End Sub

Private Sub tbRTF_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
    Select Case Button.Key
        Case "Bold"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelBold = Button.Value
                Case "Notes"
                    rtfNotes.SelBold = Button.Value
            End Select
        Case "Italic"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelItalic = Button.Value
                Case "Notes"
                    rtfNotes.SelItalic = Button.Value
            End Select
        Case "Underline"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelUnderline = Button.Value
                Case "Notes"
                    rtfNotes.SelUnderline = Button.Value
            End Select
        Case "Left"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelAlignment = rtfLeft
                Case "Notes"
                    
                    rtfNotes.SelAlignment = rtfLeft
            End Select
        Case "Center"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelAlignment = rtfCenter
                Case "Notes"
                    rtfNotes.SelAlignment = rtfCenter
            End Select
        Case "Right"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    rtfCode.SelAlignment = rtfRight
                Case "Notes"
                    rtfNotes.SelAlignment = rtfRight
            End Select
        Case "Font Colour"
            If tbRTF.Buttons("Font Colour").Value = tbrPressed Then
                Load frmColour
                frmColour.Show 0, Me
                tbRTF.Buttons("Font Colour").Value = tbrPressed
            End If

        Case "BulList"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                    'rtfCode.BulletIndent = 10
                Case "Notes"
            End Select
        
        Case "NumList"
            Select Case tbsCodeTabs.SelectedItem.Key
                Case "Code"
                Case "Notes"
            End Select
            
        Case "ChangeCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, varChangeTo, varFirstCharCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, varChangeTo, varFirstCharCase)
            End If

        Case "Symbol"
            Load frmCMap
            frmCMap.Show 0, Me
    End Select
    
End Sub

Private Sub tbRTF_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

    On Error Resume Next
    
    Select Case ButtonMenu.Key
        Case "LowerCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, LowerCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, LowerCase)
            End If
            varChangeTo = LowerCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Lower Case"
            
        Case "UpperCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, UpperCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, UpperCase)
            End If
            varChangeTo = UpperCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Upper Case"
            
        Case "SentenceCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, SentenceCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, SentenceCase)
            End If
            varChangeTo = SentenceCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Sentence Case"
            
        Case "ToggleCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, ToggleCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, ToggleCase)
            End If
            varChangeTo = ToggleCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Toggle Case"
            
        Case "TitleCase"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, TitleCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, TitleCase)
            End If
            varChangeTo = TitleCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Title Case"
            
        Case "VaryCaseLower"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, VaryCase, vbLowerCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, VaryCase, vbLowerCase)
            End If
            varChangeTo = VaryCase
            varFirstCharCase = LowerCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Vary Case - Lower First"
            
        Case "VaryCaseUpper"
            If tbsCodeTabs.SelectedItem.Key = "Code" Then
                rtfCode.SelText = ChangeCase(rtfCode.SelText, VaryCase, vbUpperCase)
            ElseIf tbsCodeTabs.SelectedItem.Key = "Notes" Then
                rtfNotes.SelText = ChangeCase(rtfNotes.SelText, VaryCase, vbUpperCase)
            End If
            varChangeTo = VaryCase
            varFirstCharCase = UpperCase
            ButtonMenu.Parent.ToolTipText = "Change Case to Vary Case - Upper First"
            
    End Select
    
End Sub

Private Sub tbRTF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(FormatToolbar)

End Sub

Private Sub tbsCodeTabs_Click()

    On Error Resume Next 'Goto next line on error

    Select Case tbsCodeTabs.SelectedItem.Key
        Case "Code"
            rtfCode.Visible = True
            rtfNotes.Visible = False
            lvBookmarks.Visible = False
            If mnuPop_Format.Checked = True Then tbRTF.Visible = True
        Case "Notes"
            rtfCode.Visible = False
            rtfNotes.Visible = True
            lvBookmarks.Visible = False
            If mnuPop_Format.Checked = True Then tbRTF.Visible = True
        Case "Bookmarks"
            rtfCode.Visible = False
            rtfNotes.Visible = False
            lvBookmarks.Visible = True
            tbRTF.Visible = False
    End Select
    Call Resize
    
End Sub

Private Sub tbsCodeTabs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(Form)
    
End Sub

Private Sub tbStandard_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next 'Goto next line on error
    
    Select Case Button.Key
        Case "New"
            Call mnuFile_New_Click
        Case "Open"
            Call mnuFile_Open_Click
        Case "Save"
            Call mnuFile_Save_Click
        Case "SaveAs"
            Call mnuFile_SaveAs_Click
        Case "SaveAll"
            Call mnuFile_SaveAll_Click
            
        Case "PrintPreview"
            Call mnuFile_PrintPreview_Click
        Case "Print"
            Call mnuFile_Print_Click
        Case "SpellCheck"
            Call mnuTools_SpellCheck_Click
        Case "Find"
            Call mnuEdit_Find_Click
            
        Case "Cut"
            Call mnuEdit_Cut_Click
        Case "Copy"
            Call mnuEdit_Copy_Click
        Case "Paste"
            Call mnuEdit_Paste_Click
        Case "Notes"
            Call mnuTools_CopyToNotes_Click
        Case "Bookmarks"
            Call mnuBookmarks_Add_Click
        
        Case "Undo"
            Call mnuEdit_Undo_Click
        Case "Redo"
            Call mnuEdit_Redo_Click
        
        Case "Help"
            Call mnuHelp_Help_Click
            
        Case Else
            MsgBox Button.Key & " is not finished yet - Sorry"
    End Select
    
End Sub

Private Sub tbStandard_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    On Error Resume Next
    
    Select Case Button.Key
        Case "NewCode"
        Case "NewNotes"
        Case "OpenCode"
        Case "OpenNotes"
    End Select
    
End Sub

Private Sub tbStandard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(StandardToolbar)
        
End Sub

Private Sub tmrHideTree_Timer()
    
    On Error Resume Next
    
    If varTreeLeft > Int("-" & tvCodeType.Width) Then
        varTreeLeft = varTreeLeft - 50
        Call Resize
    Else
        'Enable the show buttons
        mnuTree_Show.Enabled = True
        mnuPop_Show.Enabled = True
        tmrHideTree.Enabled = False
    End If
        
End Sub

Private Sub tmrShowTree_Timer()

    On Error Resume Next
    
    If varTreeLeft < 0 Then
        varTreeLeft = varTreeLeft + 50
        Call Resize
    Else
        'Enable the hide buttons
        mnuTree_Hide.Enabled = True
        mnuPop_Hide.Enabled = True
        'Enable the drag line
        picDragLine.Enabled = True
        picDragLine.Visible = True
        'Set the tree hidden variable
        varTreeHidden = False
        tmrShowTree.Enabled = False
    End If
    
End Sub

Private Sub tmrVisualSettings_Timer()
    
    On Error Resume Next 'Goto next line on error
    Static LastCursorPos As POINTAPI 'Var to hold the last cursor pos
    Dim CurrentCursorPos As POINTAPI 'Var to hold the current cursor pos
    
    'Find which tab is selected
    Select Case tbsCodeTabs.SelectedItem.Key
        'Code
        Case "Code"
            'If some code is selected and the copy to notes or cut button is disabled
            If rtfCode.SelLength > 0 And (tbStandard.Buttons("Notes").Enabled = False Or tbStandard.Buttons("Cut").Enabled = False) Then
                'Enable Cut, Copy & Notes Buttons
                tbStandard.Buttons("Cut").Enabled = True
                tbStandard.Buttons("Copy").Enabled = True
                tbStandard.Buttons("Notes").Enabled = True
            ElseIf rtfCode.SelLength = 0 And (tbStandard.Buttons("Notes").Enabled = True Or tbStandard.Buttons("Cut").Enabled = True) Then
                tbStandard.Buttons("Cut").Enabled = False
                tbStandard.Buttons("Copy").Enabled = False
                tbStandard.Buttons("Notes").Enabled = False
            End If
            'Enabled the paste if needed
            If tbStandard.Buttons("Paste").Enabled = False And (Clipboard.GetText(ccCFRTF) <> "" Or Clipboard.GetText(ccCFText) <> "") Then
                tbStandard.Buttons("Paste").Enabled = True
            'Disable the paste button if needed
            ElseIf tbStandard.Buttons("Paste").Enabled = True And Clipboard.GetText(ccCFRTF) = "" And Clipboard.GetText(ccCFText) = "" Then
                tbStandard.Buttons("Paste").Enabled = False
            End If
            'Enable the print button if needed
            If tbStandard.Buttons("Print").Enabled = False Then tbStandard.Buttons("Print").Enabled = True
        
        'Notes
        Case "Notes"
            'If some code is selected and the cut button is disabled
            If rtfNotes.SelLength > 0 And tbStandard.Buttons("Cut").Enabled = False Then
                'Enable Cut & Copy
                tbStandard.Buttons("Cut").Enabled = True
                tbStandard.Buttons("Copy").Enabled = True
            ElseIf rtfNotes.SelLength = 0 And tbStandard.Buttons("Cut").Enabled = True Then
                tbStandard.Buttons("Cut").Enabled = False
                tbStandard.Buttons("Copy").Enabled = False
            End If
            'Enabled the paste button if needed
            If tbStandard.Buttons("Paste").Enabled = False And (Clipboard.GetText(ccCFRTF) <> "" Or Clipboard.GetText(ccCFText) <> "") Then
                tbStandard.Buttons("Paste").Enabled = True
            'Disable the paste button if needed
            ElseIf tbStandard.Buttons("Paste").Enabled = True And Clipboard.GetText(ccCFRTF) = "" And Clipboard.GetText(ccCFText) = "" Then
                tbStandard.Buttons("Paste").Enabled = False
            End If
            'Disable the copy to notes button if needed
            If tbStandard.Buttons("Notes").Enabled = True Then tbStandard.Buttons("Notes").Enabled = False
            'Enable the print button if needed
            If tbStandard.Buttons("Print").Enabled = False Then tbStandard.Buttons("Print").Enabled = True
            
            
        'Bookmarks
        Case "Bookmarks"
            'Disable Cut, Copy, Paste & Notes buttons if needed
            If tbStandard.Buttons("Cut").Enabled = True Or _
            tbStandard.Buttons("Copy").Enabled = True Or _
            tbStandard.Buttons("Paste").Enabled = True Or _
            tbStandard.Buttons("Notes").Enabled = True Then
                tbStandard.Buttons("Cut").Enabled = False
                tbStandard.Buttons("Copy").Enabled = False
                tbStandard.Buttons("Paste").Enabled = False
                tbStandard.Buttons("Notes").Enabled = False
            End If
            'Enable the print button if needed
            If tbStandard.Buttons("Print").Enabled = True Then tbStandard.Buttons("Print").Enabled = False
    End Select
    
End Sub

Private Sub tvCodeType_Collapse(ByVal Node As MSComctlLib.Node)

    On Error Resume Next 'Goto next line on error

    'Play the sound
    Call PlaySound(varTreeShrinkSound)
    'Select the node
    Call tvCodeType_NodeClick(tvCodeType.SelectedItem)
    
End Sub

Private Sub tvCodeType_Expand(ByVal Node As MSComctlLib.Node)

    On Error Resume Next 'Goto next line on error

    'Play the sound
    Call PlaySound(varTreeExpandSound)
    
End Sub

Private Sub tvCodeType_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next 'Goto next line on error

    'If the right button is presses show the popupmenu
    If Button = vbRightButton Then Call ShowPopUpMenu(TreeView)
    
End Sub

Private Sub tvCodeType_NodeClick(ByVal Node As MSComctlLib.Node)

    On Error Resume Next
    
    On Error Resume Next 'Goto next line on error
    Dim Counter As Integer
    Dim SeperatorPos As Integer
    Dim SeperatorPos2 As Integer
    
    'Play the sound
    Call PlaySound(varTreeClickSound)
    
    'Find the langauge
    For Counter = LBound(varLanguageName) To UBound(varLanguageName)
        'If the left part of the key = the Language's Key
        If Left(Node.Key, Len(varLanguageName(Counter))) = varLanguageName(Counter) Then
            'Set the status bar to that nodes text
            sbSectionInfo.Panels("Language").Text = "Language: " & varLanguageName(Counter)
            'Exit loop for effeciency
            Exit For
        'If the language is unknown
        Else
            'Set the status bar if we are on the last counter (i.e. Language is unknown)
            If Counter = UBound(varLanguageName) Then sbSectionInfo.Panels("Language").Text = "Language: None"
        End If
        'Let form repaint
        DoEvents
    Next Counter
    
    'Find the section
    'Find the _ in the nodes key
    SeperatorPos = InStr(1, Node.Key, "_", vbTextCompare)
    SeperatorPos2 = InStrRev(Node.Key, "_", , vbTextCompare)
    'If we are in a section but no code is selected
    If SeperatorPos > 0 And SeperatorPos2 = SeperatorPos Then
        'Set the Section to None in the status bar
        sbSectionInfo.Panels("Section").Text = "Section: " & Node.Text
        'Set the code's name to None in the status bar
        sbSectionInfo.Panels("Code").Text = "Code: None"
        'Load the default text the text
        If varLoadDefaultFile = True Then rtfCode.LoadFile varDefaultFile
    
    'If we are in a section and code is selected
    ElseIf SeperatorPos > 0 And SeperatorPos2 > SeperatorPos Then
        'Set the Section to None in the status bar
        sbSectionInfo.Panels("Section").Text = "Section: " & Node.Parent
        'Set the code's name to None in the status bar
        sbSectionInfo.Panels("Code").Text = "Code: " & Node.Text
        'Load the code
        Call LoadCode(Node.Key)
    
    'If there is no code or section
    ElseIf SeperatorPos = 0 Then
        'Set the Section to None in the status bar
        sbSectionInfo.Panels("Section").Text = "Section: None"
        'Set the code's name to None in the status bar
        sbSectionInfo.Panels("Code").Text = "Code: None"
        'Load the default text the text
        If varLoadDefaultFile = True Then rtfCode.LoadFile varDefaultFile
    End If
    For Counter = UBound(varHistory) To LBound(varHistory) + 1 Step -1
        'Move them all back one
        varHistory(Counter) = varHistory(Counter - 1)
    Next Counter
    'Set the current node index
    varHistory(UBound(varHistory)) = Node.Index
    varLastHistory = 0
    
End Sub

Private Sub GetCodes(ByVal Path As String, ByVal Relative As String, ByVal Section As String)

    On Error Resume Next
    Dim Counter As Integer
    
    'Add an image
    ilTreeView.ListImages.Add , Relative & "_" & Section, LoadPicture(Path & "\" & varDefaultPicture)
    'If the big file exisits
    If DoesFileExist(Path & "\" & varDefaultLargePicture) = True Then
        'Add it to the image list
        ilListLarge.ListImages.Add , Relative & "_" & Section, LoadPicture(Path & "\" & varDefaultLargePicture)
    'If it doesn't
    Else
        'Add the small one
        ilListLarge.ListImages.Add , Relative & "_" & Section, LoadPicture(Path & "\" & varDefaultPicture)
    End If
    'Add a node for the item
    tvCodeType.Nodes.Add Relative, tvwChild, Relative & "_" & Section, Section, Relative & "_" & Section
    
    'Set the file list path to the specified path Methods
    fileList.Path = Path
    
    'Get the RTFs/codes
    For Counter = 0 To fileList.ListCount - 1
        'Add a node
        tvCodeType.Nodes.Add Relative & "_" & Section, tvwChild, Relative & "_" & Section & "_Code" & Counter, fileList.List(Counter), Relative & "_" & Section
        'Increment the code count
        varCodeCount = varCodeCount + 1
    Next Counter

End Sub

Private Sub LoadCode(ByVal NodeKey As String)

    On Error GoTo ErrorHandler
    
    'Load the code
    rtfCode.LoadFile GetFileName(NodeKey)
    'Set the code's filename
    varCodeFileName = GetFileName(NodeKey)
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.LoadCode(" & NodeKey & ")", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub GetRegSettings()
 
    On Error Resume Next 'Goto next line on error
    
    varLoadDefaultFile = GetSetting(App.EXEName, "Options", "Load Default File", True)
    varDefaultFile = GetSetting(App.EXEName, "Options", "Default File", App.Path & "\Default.rtf")
    varSeperatingChar = GetSetting(App.EXEName, "Options", "Seperating Char", ";")
    varDefaultPicture = GetSetting(App.EXEName, "Options", "Default Picture", "Default.bmp")
    varDefaultLargePicture = GetSetting(App.EXEName, "Options", "Default Large Picture", "DefaultLarge.bmp")
    varBookmarksFileName = GetSetting(App.EXEName, "Options", "Bookmarks File", "Bookmarks.ini")
    varSaveBookmarks = GetSetting(App.EXEName, "Options", "Save Bookmarks", True)
    varShowDragLine = GetSetting(App.EXEName, "Options", "Save Drag Line", True)
    varAddMenuIcons = GetSetting(App.EXEName, "Options", "Menu Icons", True)
    varAutoSwitch = GetSetting(App.EXEName, "Options", "Auto Switch", True)
    tmrHideTree.Interval = GetSetting(App.EXEName, "Options", "Tree Speed", 1)
    tmrShowTree.Interval = tmrHideTree.Interval
    ilTreeView.MaskColor = GetSetting(App.EXEName, "Options", "Mask Colour", &HFF00FF)
    tvCodeType.SingleSel = GetSetting(App.EXEName, "Options", "Multi Code View", True)
    tvCodeType.HotTracking = GetSetting(App.EXEName, "Options", "Hot Tracking", True)
    tvCodeType.Style = GetSetting(App.EXEName, "Options", "Tree Style", 7)
    varUseConfig = GetSetting(App.EXEName, "Options", "Use Config", False)
    varSaveSettings = GetSetting(App.EXEName, "Options", "Save Settings", True)
    varSaveInReg = GetSetting(App.EXEName, "Options", "Save In Reg", True)
    varSeperator = GetSetting(App.EXEName, "Options", "Seperator", vbNewLine)
    varUseSound = GetSetting(App.EXEName, "Options", "Use Sounds", True)
    varTreeClickSound = GetSetting(App.EXEName, "Options", "Tree Click Sound", App.Path & "\Sounds\Tree Click.wav")
    varTreeExpandSound = GetSetting(App.EXEName, "Options", "Tree Expand Sound", App.Path & "\Sounds\Tree Expand.wav")
    varTreeShrinkSound = GetSetting(App.EXEName, "Options", "Tree Shrink Sound", App.Path & "\Sounds\Tree Shrink.wav")
    varNotesSound = GetSetting(App.EXEName, "Options", "Notes Sound", App.Path & "\Sounds\Notes.wav")
    varBookmarksSound = GetSetting(App.EXEName, "Options", "Bookmark Sound", App.Path & "\Sounds\Bookmark.wav")
    varAskQuitQuestion = GetSetting(App.EXEName, "Options", "Ask Quit Question", True)
    mnuBookmarks_View(GetSetting(App.EXEName, "Options", "Bookmarks View", 3)).Checked = True
    Call mnuBookmarks_View_Click(GetSetting(App.EXEName, "Options", "Bookmarks View", 3))
    MnuTools_OnTop.Checked = GetSetting(App.EXEName, "Options", "On Top", False)
    'Set the ontop value
    Call OnTop(Me, MnuTools_OnTop.Checked)
    varCopyFormat = GetSetting(App.EXEName, "Options", "Copy Format", 1)
    'Toolbars
    tbStandard.RestoreToolbar "Toolbars", "Standard", tbStandard
    'Window Pos
    Me.Left = GetSetting(App.EXEName, "Start-Up", "Left", 0)
    Me.Top = GetSetting(App.EXEName, "Start-Up", "Top", 0)
    Me.Width = GetSetting(App.EXEName, "Start-Up", "Width", 8000)
    Me.Height = GetSetting(App.EXEName, "Start-Up", "Height", 6000)
    Me.WindowState = GetSetting(App.EXEName, "Start-Up", "Window State", 0)
    
    'Make sure it is not off the screen
    If Me.Left < Int("-" & Me.Width) Or Me.Left > Screen.Width Then Me.Left = 0
    If Me.Top < Int("-" & Me.Height) Or Me.Top > Screen.Height Then Me.Top = 0
    
End Sub

Private Sub SaveRegSettings()
 
    On Error Resume Next 'Goto next line on error
    
    SaveSetting App.EXEName, "Options", "Save Settings", varSaveSettings
    'If we are to save the settings
    If varSaveSettings = True Then
        'Options
        SaveSetting App.EXEName, "Options", "Save In Reg", varSaveInReg
        SaveSetting App.EXEName, "Options", "Load Default File", varLoadDefaultFile
        SaveSetting App.EXEName, "Options", "Default File", varDefaultFile
        SaveSetting App.EXEName, "Options", "Seperating Char", varSeperatingChar
        SaveSetting App.EXEName, "Options", "Default Picture", varDefaultPicture
        SaveSetting App.EXEName, "Options", "Default Large Picture", varDefaultLargePicture
        SaveSetting App.EXEName, "Options", "Bookmarks File", varBookmarksFileName
        SaveSetting App.EXEName, "Options", "Save Bookmarks", varSaveBookmarks
        SaveSetting App.EXEName, "Options", "Save Drag Line", varShowDragLine
        SaveSetting App.EXEName, "Options", "Menu Icons", varAddMenuIcons
        SaveSetting App.EXEName, "Options", "Mask Colour", ilTreeView.MaskColor
        SaveSetting App.EXEName, "Options", "Auto Switch", varAutoSwitch
        SaveSetting App.EXEName, "Options", "Tree Speed", tmrHideTree.Interval
        SaveSetting App.EXEName, "Options", "Multi Code View", tvCodeType.SingleSel
        SaveSetting App.EXEName, "Options", "Hot Tracking", tvCodeType.HotTracking
        SaveSetting App.EXEName, "Options", "Tree Style", tvCodeType.Style
        SaveSetting App.EXEName, "Options", "Use Config", varUseConfig
        SaveSetting App.EXEName, "Options", "Seperator", varSeperator
        SaveSetting App.EXEName, "Options", "Use Sounds", varUseSound
        SaveSetting App.EXEName, "Options", "Tree Click Sound", varTreeClickSound
        SaveSetting App.EXEName, "Options", "Tree Expand Sound", varTreeExpandSound
        SaveSetting App.EXEName, "Options", "Tree Shrink Sound", varTreeShrinkSound
        SaveSetting App.EXEName, "Options", "Notes Sound", varNotesSound
        SaveSetting App.EXEName, "Options", "Bookmark Sound", varBookmarksSound
        SaveSetting App.EXEName, "Options", "Ask Quit Question", varAskQuitQuestion
        Select Case True
            Case mnuBookmarks_View(0).Checked
                SaveSetting App.EXEName, "Options", "Bookmarks View", 0
            Case mnuBookmarks_View(1).Checked
                SaveSetting App.EXEName, "Options", "Bookmarks View", 1
            Case mnuBookmarks_View(2).Checked
                SaveSetting App.EXEName, "Options", "Bookmarks View", 2
            Case mnuBookmarks_View(3).Checked
                SaveSetting App.EXEName, "Options", "Bookmarks View", 3
        End Select
        SaveSetting App.EXEName, "Options", "On Top", MnuTools_OnTop.Checked
        SaveSetting App.EXEName, "Options", "Copy Format", varCopyFormat
        'Toolbars
        tbStandard.SaveToolbar "Toolbars", "Standard", tbStandard


        'Window Pos
        SaveSetting App.EXEName, "Start-Up", "Left", Me.Left
        SaveSetting App.EXEName, "Start-Up", "Top", Me.Top
        If Me.WindowState = 0 Then SaveSetting App.EXEName, "Start-Up", "Width", Me.Width
        If Me.WindowState = 0 Then SaveSetting App.EXEName, "Start-Up", "Height", Me.Height
        SaveSetting App.EXEName, "Start-Up", "Window State", Me.WindowState
    'If we aren't delete them
    Else
        'Options
        SaveSetting App.EXEName, "Options", "Save In Reg", varSaveInReg
        DeleteSetting App.EXEName, "Options", "Load Default File"
        DeleteSetting App.EXEName, "Options", "Default File"
        DeleteSetting App.EXEName, "Options", "Seperating Char"
        DeleteSetting App.EXEName, "Options", "Default Picture"
        DeleteSetting App.EXEName, "Options", "Bookmarks File"
        DeleteSetting App.EXEName, "Options", "Save Bookmarks"
        DeleteSetting App.EXEName, "Options", "Save Drag Line"
        DeleteSetting App.EXEName, "Options", "Menu Icons"
        DeleteSetting App.EXEName, "Options", "Mask Colour"
        DeleteSetting App.EXEName, "Options", "Auto Switch"
        DeleteSetting App.EXEName, "Options", "Tree Speed"
        DeleteSetting App.EXEName, "Options", "Multi Code View"
        DeleteSetting App.EXEName, "Options", "Hot Tracking"
        DeleteSetting App.EXEName, "Options", "Tree Style"
        DeleteSetting App.EXEName, "Options", "Use Config"
        DeleteSetting App.EXEName, "Options", "Seperator"
        DeleteSetting App.EXEName, "Options", "Use Sounds"
        DeleteSetting App.EXEName, "Options", "Tree Expand Sound"
        DeleteSetting App.EXEName, "Options", "Tree Expand Sound"
        DeleteSetting App.EXEName, "Options", "Tree Shrink Sound"
        DeleteSetting App.EXEName, "Options", "Notes Sound"
        DeleteSetting App.EXEName, "Options", "Bookmark Sound"
        DeleteSetting App.EXEName, "Options", "Bookmarks View"
        DeleteSetting App.EXEName, "Options", "On Top"
        DeleteSetting App.EXEName, "Options", "Copy Format"
        'Toolbars
        DeleteSetting App.EXEName, "Toolbars", "Standard"
        'Window Pos
        DeleteSetting App.EXEName, "Start-Up", "Left"
        DeleteSetting App.EXEName, "Start-Up", "Top"
        DeleteSetting App.EXEName, "Start-Up", "Window State"
    End If
    
End Sub

Private Function GetFileName(ByVal NodeKey As String) As String
    
    On Error Resume Next 'Goto next line on error
    
    'Find the language
    GetFileName = App.Path & "\" & Mid(sbSectionInfo.Panels("Language").Text, 11)
    'Find the section
    GetFileName = GetFileName & "\" & Mid(sbSectionInfo.Panels("Section").Text, 10)
    'Find the code
    GetFileName = GetFileName & "\" & Mid(sbSectionInfo.Panels("Code").Text, 7)
    
End Function

Public Sub SelectNode(NodeNo As Integer)
    
    On Error Resume Next 'Goto next line on error

    'Select it in the tree view
    tvCodeType.Nodes(NodeNo).Selected = True
    'Make out like it has actually been click so as to load the code
    Call tvCodeType_NodeClick(tvCodeType.Nodes(NodeNo))
    
End Sub

Private Sub SaveBookmarks()

    On Error GoTo ErrorHandler
    Dim Counter As Integer
    Dim FileNumber As Byte
    Dim Language As String
    Dim Section As String
    Dim Code As String
    
    'If the bookmarks are to be saved
    If varSaveBookmarks = True Then
        'Find a free file number
        FileNumber = FreeFile
        'Open the bookmarks file for output
        Open App.Path & "\" & varBookmarksFileName For Output As #FileNumber
        'If there are bookmarks
        If lvBookmarks.ListItems.Count > 0 Then
            'Loop for all bookmarks
            For Counter = 1 To lvBookmarks.ListItems.Count
                'Add the langauge, section and code if not empty
                If lvBookmarks.ListItems(Counter).Text <> "None" Then Language = lvBookmarks.ListItems(Counter).Text
                If lvBookmarks.ListItems(Counter).ListSubItems(1).Text <> "None" Then Section = "\" & lvBookmarks.ListItems(Counter).ListSubItems(1).Text
                If lvBookmarks.ListItems(Counter).ListSubItems(2).Text <> "None" Then Code = "\" & lvBookmarks.ListItems(Counter).ListSubItems(2).Text
                'Output the key, language, section and code to the file
                Print #FileNumber, Language & Section & Code & ";"
                'Let form repaint
                DoEvents
            'Onto next bookmark
            Next Counter
            
        'If there are no bookmarks
        Else
            'Output nothing
            Print #FileNumber, ""
        
        End If
        'Close the file
        Close #FileNumber
        
    End If
    'Exit the sub routine as there is no error
    Exit Sub
    
ErrorHandler:
    Close #FileNumber
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.SaveBookmarks", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Private Sub GetBookmarks()

    On Error GoTo ErrorHandler
    Dim Counter As Integer
    Dim LineStart As Long
    Dim LineEnd As Long
    Dim FileText As String
    Dim Language As String
    Dim Section As String
    Dim Code As String
    Dim LanguageEnd As Long
    Dim SectionEnd As Long
    Dim IconNo As Integer
    Dim TempString  As String
    
    LineStart = 1
    
    'If the bookmarks are to be saved
    If varSaveBookmarks = True Then
        'Get the text of the file
        FileText = OpenText(App.Path & "\" & varBookmarksFileName)
        'Loop for all lines in file
        Do While LineStart > 0
            'Find the end of the line
            LineEnd = InStr(LineStart, FileText, ";", vbTextCompare)
            
            'If there is another line
            If LineEnd > LineStart Then
                'Find the end of the language
                If InStr(LineStart, FileText, "\", vbTextCompare) < LineEnd Then
                    LanguageEnd = InStr(LineStart, FileText, "\", vbTextCompare)
                    If InStr(LanguageEnd + 1, FileText, "\", vbTextCompare) < LineEnd Then
                        SectionEnd = InStr(LanguageEnd + 1, FileText, "\", vbTextCompare)
                    Else
                        SectionEnd = LineEnd
                    End If
                    
                Else
                    LanguageEnd = LineEnd
                    SectionEnd = LineEnd
                End If
                
                
                'Get the language, section and code
                Language = Mid(FileText, LineStart, LanguageEnd - LineStart)
                If SectionEnd > LanguageEnd Then
                    Section = Mid(FileText, LanguageEnd + 1, SectionEnd - LanguageEnd)
                Else
                    Section = ""
                    Code = ""
                End If
                If SectionEnd <> LineEnd Then
                    Code = Mid(FileText, SectionEnd + 1, LineEnd - SectionEnd)
                Else
                    Code = ""
                End If
                'Trim the last char if needed
                If Len(Section) > 0 Then Section = Left(Section, Len(Section) - 1)
                If Len(Code) > 0 Then Code = Left(Section, Len(Code) - 1)
                For Counter = 0 To tvCodeType.Nodes.Count
                    TempString = Language & "_" & Section
                    If Right(tvCodeType.Nodes(Counter).Key, InStrRev(tvCodeType.Nodes(Counter).Key, "_", , vbTextCompare) + 1) = Code Then _
                    TempString = TempString & "_Code" & Right(tvCodeType.Nodes(Counter).Key, InStrRev(tvCodeType.Nodes(Counter).Key, "Code", , vbTextCompare) + 4)
                Next Counter
                Debug.Print Language & " " & Section & " " & Code
                lvBookmarks.ListItems.Add , "BM" & lvBookmarks.ListItems.Count, Language ', tvCodeType.Nodes(TempString).Image, tvCodeType.Nodes(TempString).Image
                lvBookmarks.ListItems("BM" & lvBookmarks.ListItems.Count - 1).ListSubItems.Add , , Section
                lvBookmarks.ListItems("BM" & lvBookmarks.ListItems.Count - 1).ListSubItems.Add , , Code
                
                LineStart = LineEnd + Len(vbNewLine & ";")
            'If there isn't
            Else
                'Make line start = 0 to stop the loop
                LineStart = 0
            End If
        'Onto next line
        Loop
    End If
    'Exit the sub routine as there is no error
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "frmMain.GetBookmarks", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Public Sub AddMenuIcons()
    
    On Error Resume Next 'Goto next line on error
    
    'Put the icons in the menus
    With ilMenus
    'File
    Call MenuIcons(Me, 0, 0, .ListImages("New").Picture)  'imgMenuFileIcon(0).Picture)
    Call MenuIcons(Me, 0, 1, .ListImages("Open").Picture)
    Call MenuIcons(Me, 0, 3, .ListImages("Save").Picture)
    Call MenuIcons(Me, 0, 4, .ListImages("SaveAs").Picture)
    Call MenuIcons(Me, 0, 5, .ListImages("SaveAll").Picture)
    Call MenuIcons(Me, 0, 7, .ListImages("PrintPreview").Picture)
    Call MenuIcons(Me, 0, 8, .ListImages("Print").Picture)
    Call MenuIcons(Me, 0, 10, .ListImages("NewWin").Picture)
    Call MenuIcons(Me, 0, 11, .ListImages("Exit").Picture)
    'Edit
    Call MenuIcons(Me, 1, 0, .ListImages("Undo").Picture)
    Call MenuIcons(Me, 1, 1, .ListImages("Redo").Picture)
    Call MenuIcons(Me, 1, 3, .ListImages("Cut").Picture)
    Call MenuIcons(Me, 1, 4, .ListImages("Copy").Picture)
    Call MenuIcons(Me, 1, 5, .ListImages("Paste").Picture)
    Call MenuIcons(Me, 1, 6, .ListImages("Clear").Picture)
    Call MenuIcons(Me, 1, 7, .ListImages("SelectAll").Picture)
    Call MenuIcons(Me, 1, 9, .ListImages("Find").Picture)
    Call MenuIcons(Me, 1, 10, .ListImages("Goto").Picture)
    'View
    Call MenuIcons(Me, 2, 0, .ListImages("Refresh").Picture)
    'Bookmarks
    Call MenuIcons(Me, 3, 0, .ListImages("Bookmark").Picture)
    Call MenuIcons(Me, 3, 1, .ListImages("ClearBookmarks").Picture)
    Call MenuIcons(Me, 3, 3, .ListImages("IconLarge").Picture, .ListImages("IconLarge").Picture)
    Call MenuIcons(Me, 3, 4, .ListImages("IconSmall").Picture, .ListImages("IconSmall").Picture)
    Call MenuIcons(Me, 3, 5, .ListImages("IconList").Picture, .ListImages("IconList").Picture)
    Call MenuIcons(Me, 3, 6, .ListImages("IconReport").Picture, .ListImages("IconReport").Picture)
    'Tree
    Call MenuIcons(Me, 4, 0, .ListImages("Rebuild").Picture)
    Call MenuIcons(Me, 4, 2, .ListImages("Expand").Picture)
    Call MenuIcons(Me, 4, 3, .ListImages("Shrink").Picture)
    Call MenuIcons(Me, 4, 4, .ListImages("Code").Picture)
    Call MenuIcons(Me, 4, 6, .ListImages("Cross").Picture)
    Call MenuIcons(Me, 4, 7, .ListImages("Restore").Picture)
    'Tools
    Call MenuIcons(Me, 5, 0, .ListImages("Notes").Picture)
    Call MenuIcons(Me, 5, 1, .ListImages("SpellCheck").Picture)
    Call MenuIcons(Me, 5, 3, .ListImages("Properties").Picture)
    Call MenuIcons(Me, 5, 4, .ListImages("Error").Picture)
    'Help
    Call MenuIcons(Me, 6, 0, .ListImages("Help").Picture)
    Call MenuIcons(Me, 6, 1, .ListImages("ReadMe").Picture)
    Call MenuIcons(Me, 6, 3, .ListImages("About").Picture)
    
    End With
End Sub

Public Sub RemoveMenuIcons()
    
    On Error Resume Next 'Goto next line on error
    
    'Take the icons out of the menus
    'File
    Call MenuIcons(Me, 0, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 1, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 3, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 4, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 5, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 7, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 8, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 0, 10, imgBlankMenuIcon.Picture)
    'Edit
    Call MenuIcons(Me, 1, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 1, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 3, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 4, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 5, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 6, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 7, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 9, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 10, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 1, 11, imgBlankMenuIcon.Picture)
    'View
    Call MenuIcons(Me, 2, 0, imgBlankMenuIcon.Picture)
    'Bookmarks
    Call MenuIcons(Me, 3, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 3, 1, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 3, 3, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 3, 4, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 3, 5, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 3, 6, imgBlankMenuIcon.Picture)
    'Tree
    Call MenuIcons(Me, 4, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 4, 2, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 4, 3, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 4, 4, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 4, 6, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 4, 7, imgBlankMenuIcon.Picture)
    'Tools
    Call MenuIcons(Me, 5, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 5, 1, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 5, 3, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 5, 4, imgBlankMenuIcon.Picture)
    'Help
    Call MenuIcons(Me, 6, 0, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 6, 1, imgBlankMenuIcon.Picture)
    Call MenuIcons(Me, 6, 3, imgBlankMenuIcon.Picture)
    
End Sub

Private Sub ShowPopUpMenu(MenuType As PopUpType)
    
    On Error Resume Next 'Goto next line on error

    'Hide all popups
    mnuPop_Standard.Visible = True 'This one must be visible as at _
                                    least one menu must be visible at all times
    mnuPop_Format.Visible = False
    mnuPop_Bookmark.Visible = False
    mnuPop_Clear.Visible = False
    mnuPop_Copy.Visible = False
    mnuPop_CustomizeStandard.Visible = False
    mnuPop_CustomizeFormat.Visible = False
    mnuPop_Cut.Visible = False
    mnuPop_Delete.Visible = False
    mnuPop_Hide.Visible = False
    mnuPop_Notes.Visible = False
    mnuPop_Paste.Visible = False
    mnuPop_Refresh.Visible = False
    mnuPop_SelectAll.Visible = False
    mnuPop_SelectAll.Visible = False
    mnuPop_Seperator1.Visible = False
    mnuPop_Seperator2.Visible = False
    mnuPop_Seperator3.Visible = False
    mnuPop_Seperator4.Visible = False
    mnuPop_Show.Visible = False
    'Set whether the standard should be checked
    mnuPop_Standard.Checked = tbStandard.Visible
    
    'Find which one is to be seen
    Select Case MenuType
        'Format Toolbar
        Case FormatToolbar
            'mnuPop_Standard.Visible = True
            mnuPop_Format.Visible = True
            mnuPop_Seperator1.Visible = True
            mnuPop_CustomizeFormat.Visible = True
        
        'Standard Toolbar
        Case StandardToolbar
            'mnuPop_Standard.Visible = True
            mnuPop_Format.Visible = True
            mnuPop_Seperator1.Visible = True
            mnuPop_CustomizeStandard.Visible = True
            
        'Tree view
        Case TreeView
            mnuPop_Bookmark.Visible = True
            mnuPop_Seperator3.Visible = True
            mnuPop_Refresh.Visible = True
            mnuPop_Seperator4.Visible = True
            mnuPop_Hide.Visible = True
            mnuPop_Show.Visible = True
            mnuPop_Standard.Visible = False
            
        'Code
        Case Code
            mnuPop_Cut.Visible = True
            mnuPop_Copy.Visible = True
            mnuPop_Paste.Visible = True
            mnuPop_SelectAll.Visible = True
            mnuPop_Clear.Visible = True
            mnuPop_Seperator2.Visible = True
            mnuPop_Notes.Visible = True
            mnuPop_Bookmark.Visible = True
            mnuPop_Seperator3.Visible = True
            mnuPop_Refresh.Visible = True
            mnuPop_Standard.Visible = False
            
        'Notes
        Case Notes
            mnuPop_Cut.Visible = True
            mnuPop_Copy.Visible = True
            mnuPop_Clear.Visible = True
            mnuPop_Paste.Visible = True
            mnuPop_Seperator2.Visible = True
            mnuPop_Bookmark.Visible = True
            mnuPop_Seperator3.Visible = True
            mnuPop_Refresh.Visible = True
            mnuPop_Standard.Visible = False
            
        'Bookmarks
        Case Bookmark
            mnuPop_SelectAll.Visible = True
            mnuPop_Delete.Visible = True
            mnuPop_Seperator3.Visible = True
            mnuPop_Refresh.Visible = True
            mnuPop_Standard.Visible = False
            
        'Form
        Case Form
            'mnuPop_Standard.Visible = True
            mnuPop_Format.Visible = True
            mnuPop_Seperator1.Visible = True
            mnuPop_Bookmark.Visible = True
            mnuPop_Seperator3.Visible = True
            mnuPop_Refresh.Visible = True
            
    End Select
    
    'Set what things should be enabled
    If lvBookmarks.ListItems.Count > 0 Then
        mnuPop_Delete.Enabled = True
    Else
        mnuPop_Delete.Enabled = False
    End If
    
    mnuPop_Copy.Enabled = tbStandard.Buttons("Copy").Enabled
    mnuPop_Cut.Enabled = tbStandard.Buttons("Cut").Enabled
    mnuPop_Clear.Enabled = tbStandard.Buttons("Cut").Enabled
    mnuPop_Notes.Enabled = tbStandard.Buttons("Notes").Enabled
    mnuPop_Paste.Enabled = tbStandard.Buttons("Paste").Enabled
    If tbsCodeTabs.SelectedItem.Key = "Bookmarks" Then
        mnuPop_Format.Enabled = False
    Else
        mnuPop_Format.Enabled = True
    End If
    'Show the menu
    PopupMenu mnuPop
    
End Sub

Private Sub SaveFiles(ByVal TabKey As TabKeys, Optional ByVal Filename As String = "")

    On Error GoTo ErrorHandler
    Dim DialogTitle As String
    
    If TabKey = CodeWindow Then
        DialogTitle = "Save Code As"
    ElseIf TabKey = NotesWindow Then
        DialogTitle = "Save Notes As"
    ElseIf TabKey = BookmarksWindow Then
        DialogTitle = "Save Bookmarks As"
    End If
    
    'If there is no file name
    If Filename = "" And TabKey <> BookmarksWindow Then
        With CD1
            'Set the dialog title
            .DialogTitle = DialogTitle
            'Cause an error when user chooses cancel
            .CancelError = True
            'Set the Filter
            .Filter = "Rich Text Format (*.rtf)|*.rtf" '|All Files|*.*"
            'Show the save dialog
            .ShowSave
            'Make the filename = the user's choosen name
            Filename = .Filename
        End With
    End If
    
    'Find which tab we need to save
    Select Case TabKey
        'Code
        Case CodeWindow
            'save the code
            rtfCode.SaveFile Filename
        'notes
        Case NotesWindow
            varNotesFileName = Filename
            'Save the notes
            rtfNotes.SaveFile varNotesFileName
        'Boomarks
        Case BookmarksWindow
            Call SaveBookmarks
    End Select
    'Exit sub so as not cause an error
    Exit Sub
    
ErrorHandler:
    Call frmErrorDialog.ShowDialog(Err.Number, "Code Tree: frmMain - SaveFiles(" & TabKey & ", " & Filename & ")", _
        "rickbull@rickmusic.co.uk", "E-Mail the Author", "Click to E-Mail", , "Error", _
        edfIncludeReport Or edfSoundBeep Or edfWriteErrorLog)
End Sub

Public Sub SelectNone()

    On Error Resume Next
    
    'Makes no node selected
    tvCodeType.Nodes(tvCodeType.SelectedItem.Key).Selected = False
    sbSectionInfo.Panels("Language").Text = "Language: None"
    sbSectionInfo.Panels("Section").Text = "Section: None"
    sbSectionInfo.Panels("Code").Text = "Code: None"
    
End Sub

Private Sub History(Optional ByVal HowFar As Integer = -1)

    On Error Resume Next
    Dim Counter As Integer
    
    Call SelectNode(varHistory(HowFar + 1)) 'varLastHistory + (HowFar - 1)))
    For Counter = LBound(varHistory) To UBound(varHistory) - 1
        varHistory(Counter) = varHistory(Counter + 1)
        'Let form repaint
        DoEvents
    Next Counter
    
End Sub

Private Sub SaveINISettings()
 
    On Error Resume Next 'Goto next line on error
    
    SaveSetting App.EXEName, "Options", "Save Settings", varSaveSettings
    SaveSetting App.EXEName, "Options", "Save In Reg", varSaveInReg
    'If we are to save the settings
    If varSaveSettings = True Then
        'Options
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Ask Quit Question", varAskQuitQuestion)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Auto Switch", varAutoSwitch)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Bookmarks File", varBookmarksFileName)
        Select Case True
            Case mnuBookmarks_View(0).Checked
                Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Bookmarks View", 0)
            Case mnuBookmarks_View(1).Checked
                Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Bookmarks View", 1)
            Case mnuBookmarks_View(2).Checked
                Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Bookmarks View", 2)
            Case mnuBookmarks_View(3).Checked
                Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Bookmarks View", 3)
        End Select
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Copy Format", varCopyFormat)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Default File", varDefaultFile)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Default Large Picture", varDefaultLargePicture)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Default Picture", varDefaultPicture)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Hot Tracking", tvCodeType.HotTracking)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Load Default File", varLoadDefaultFile)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Mask Colour", ilTreeView.MaskColor)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Menu Icons", varAddMenuIcons)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Multi Code View", tvCodeType.SingleSel)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "On Top", MnuTools_OnTop.Checked)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Save Bookmarks", varSaveBookmarks)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Save Drag Line", varShowDragLine)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Seperating Char", varSeperatingChar)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Seperator", varSeperator)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Tree Speed", tmrHideTree.Interval)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Tree Style", tvCodeType.Style)
        Call SaveINI(App.Path & "\INIConfig.ini", "Options", "Use Config", varUseConfig)
        
        'Sounds
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Bookmark Sound", varBookmarksSound)
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Notes Sound", varNotesSound)
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Tree Click Sound", varTreeClickSound)
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Tree Expand Sound", varTreeExpandSound)
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Tree Shrink Sound", varTreeShrinkSound)
        Call SaveINI(App.Path & "\INIConfig.ini", "Sounds", "Use Sounds", varUseSound)
        
        'Window Pos
        If Me.WindowState = 0 Then Call SaveINI(App.Path & "\INIConfig.ini", "Start-Up", "Height", Me.Height)
        Call SaveINI(App.Path & "\INIConfig.ini", "Start-Up", "Left", Me.Left)
        Call SaveINI(App.Path & "\INIConfig.ini", "Start-Up", "Top", Me.Top)
        If Me.WindowState = 0 Then Call SaveINI(App.Path & "\INIConfig.ini", "Start-Up", "Width", Me.Width)
        Call SaveINI(App.Path & "\INIConfig.ini", "Start-Up", "Window State", Me.WindowState)
    End If
    
End Sub

Private Sub SetToolBarButtons(ByVal RTFName As RichTextBox)

    On Error Resume Next
    
    cmbFontName.Text = RTFName.SelFontName
    cmbFontSize.Text = RTFName.SelFontSize
    ColourButton1.Colour = RTFName.SelColor
    With tbRTF
        If RTFName.SelBold = True Then
            .Buttons("Bold").Value = tbrPressed
        Else
            .Buttons("Bold").Value = tbrUnpressed
        End If
        If RTFName.SelItalic = True Then
            .Buttons("Italic").Value = tbrPressed
        Else
            .Buttons("Italic").Value = tbrUnpressed
        End If
        If RTFName.SelUnderline = True Then
            .Buttons("Underline").Value = tbrPressed
        Else
            .Buttons("Underline").Value = tbrUnpressed
        End If
        
        Select Case RTFName.SelAlignment
            Case rtfLeft
                .Buttons("Left").Value = tbrPressed
            Case rtfCenter
                .Buttons("Center").Value = tbrPressed
            Case rtfRight
                .Buttons("Right").Value = tbrPressed
        End Select
    End With
    
End Sub
