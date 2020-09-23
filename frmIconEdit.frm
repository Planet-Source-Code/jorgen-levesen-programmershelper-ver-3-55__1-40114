VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIconEdit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Icon editor (32x32)"
   ClientHeight    =   6120
   ClientLeft      =   1290
   ClientTop       =   2415
   ClientWidth     =   8910
   Icon            =   "frmIconEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   Begin VB.CommandButton cmdExit 
      Height          =   375
      Left            =   120
      Picture         =   "frmIconEdit.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Exit"
      Top             =   240
      Width           =   375
   End
   Begin VB.PictureBox picBrushSize 
      Height          =   900
      Left            =   150
      Picture         =   "frmIconEdit.frx":0454
      ScaleHeight     =   56
      ScaleMode       =   0  'User
      ScaleWidth      =   13
      TabIndex        =   53
      Top             =   930
      Width           =   255
   End
   Begin VB.PictureBox PicRGBMix 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5820
      ScaleHeight     =   165
      ScaleWidth      =   315
      TabIndex        =   52
      Top             =   2400
      Width           =   345
   End
   Begin VB.PictureBox picMaskTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   4320
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   51
      Top             =   3330
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picImageTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   585
      Left            =   3600
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   50
      Top             =   3330
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   4320
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   49
      Top             =   2580
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picImage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   585
      Left            =   3600
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   48
      Top             =   2580
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Frame fraPaint 
      Height          =   645
      Left            =   450
      TabIndex        =   47
      Top             =   840
      Width           =   555
      Begin VB.Image imgPaint 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   390
         Left            =   120
         Picture         =   "frmIconEdit.frx":0F16
         ToolTipText     =   "Paint"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.PictureBox PicIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2040
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   45
      Top             =   240
      Width           =   510
   End
   Begin VB.Frame fraIconsContainer 
      Height          =   825
      Left            =   3210
      TabIndex        =   37
      Top             =   30
      Width           =   2265
      Begin VB.PictureBox PanelIcons 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   2
         Left            =   1620
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   43
         Top             =   210
         Width           =   510
      End
      Begin VB.PictureBox PanelIcons 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   1
         Left            =   870
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   42
         Top             =   210
         Width           =   510
      End
      Begin VB.PictureBox PanelIcons 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         ForeColor       =   &H80000008&
         Height          =   510
         Index           =   0
         Left            =   120
         Picture         =   "frmIconEdit.frx":1598
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   41
         Top             =   210
         Width           =   510
      End
      Begin VB.Label lblIcons 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   2
         Left            =   1500
         TabIndex        =   40
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblIcons 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   1
         Left            =   750
         TabIndex        =   39
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblIcons 
         BackColor       =   &H00FFFFFF&
         Height          =   675
         Index           =   0
         Left            =   30
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.Frame FraRegion 
      Height          =   645
      Left            =   450
      TabIndex        =   36
      Top             =   1380
      Width           =   555
      Begin VB.Image imgRegion 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":19DA
         ToolTipText     =   "Region"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdNew 
      Height          =   375
      Left            =   480
      Picture         =   "frmIconEdit.frx":205C
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "New file"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdSave 
      Height          =   375
      Left            =   1200
      Picture         =   "frmIconEdit.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Save to disk"
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   375
      Left            =   840
      Picture         =   "frmIconEdit.frx":27C8
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Open file"
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame fraFill 
      Height          =   645
      Left            =   450
      TabIndex        =   32
      Top             =   1920
      Width           =   555
      Begin VB.Image imgFill 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":28CA
         ToolTipText     =   "Fill"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame fraLine 
      Height          =   645
      Left            =   450
      TabIndex        =   31
      Top             =   2460
      Width           =   555
      Begin VB.Image imgLine 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":2F4C
         ToolTipText     =   "Line (Shift if for 45 or 90 degree)"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame fraSquare 
      Height          =   645
      Left            =   450
      TabIndex        =   30
      Top             =   3000
      Width           =   555
      Begin VB.Image imgSquare 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":35CE
         ToolTipText     =   "Rectangle (Shift key if for square)"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame fraFilledSquare 
      Height          =   645
      Left            =   450
      TabIndex        =   29
      Top             =   3540
      Width           =   555
      Begin VB.Image imgFilledSquare 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":3C50
         ToolTipText     =   "Filled Rectangle (Shift key if for square)"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.PictureBox picAuto 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   1410
      ScaleHeight     =   37
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   28
      Top             =   4470
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1410
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   123
      TabIndex        =   27
      Top             =   1230
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Frame fraCircle 
      Height          =   645
      Left            =   450
      TabIndex        =   26
      Top             =   4080
      Width           =   555
      Begin VB.Image imgCircle 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":42D2
         ToolTipText     =   "Oval (Shift key if for circle)"
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.PictureBox picUndo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   4530
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   1215
      Left            =   5640
      TabIndex        =   21
      Top             =   870
      Width           =   2955
      Begin VB.PictureBox PicColorPalette 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   120
         ScaleHeight     =   885
         ScaleWidth      =   2715
         TabIndex        =   22
         Top             =   210
         Width           =   2745
      End
   End
   Begin VB.Frame fraColorToUse 
      Height          =   495
      Left            =   6810
      TabIndex        =   19
      ToolTipText     =   "Color to paint"
      Top             =   150
      Width           =   585
      Begin VB.PictureBox picColorToUse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   315
         TabIndex        =   20
         Top             =   210
         Width           =   345
      End
   End
   Begin VB.PictureBox picBigContainer 
      BackColor       =   &H00E0E0E0&
      Height          =   4365
      Left            =   1080
      ScaleHeight     =   287
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   15
      Top             =   930
      Width           =   4455
      Begin VB.PictureBox picBig 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   3870
         Left            =   270
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   254
         TabIndex        =   16
         Top             =   270
         Width           =   3840
      End
   End
   Begin VB.Frame fraColorMixer 
      Caption         =   "Color mixer"
      Height          =   3195
      Left            =   5640
      TabIndex        =   0
      Top             =   2100
      Width           =   2955
      Begin VB.PictureBox pixMixContainer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2985
         Left            =   2340
         ScaleHeight     =   2985
         ScaleWidth      =   525
         TabIndex        =   13
         Top             =   150
         Width           =   525
         Begin VB.PictureBox picMixGradient 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2745
            Left            =   150
            ScaleHeight     =   2715
            ScaleWidth      =   285
            TabIndex        =   14
            Top             =   120
            Width           =   315
         End
      End
      Begin VB.PictureBox PicRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1740
         ScaleHeight     =   165
         ScaleWidth      =   195
         TabIndex        =   12
         Top             =   2460
         Width           =   225
      End
      Begin VB.PictureBox PicRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   1740
         ScaleHeight     =   165
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   1500
         Width           =   225
      End
      Begin VB.PictureBox PicRGB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   1710
         ScaleHeight     =   165
         ScaleWidth      =   195
         TabIndex        =   10
         Top             =   600
         Width           =   225
      End
      Begin VB.HScrollBar HsrRed 
         Height          =   225
         Left            =   150
         Max             =   255
         TabIndex        =   9
         Top             =   930
         Width           =   2175
      End
      Begin VB.HScrollBar HsrBlue 
         Height          =   225
         Left            =   120
         Max             =   255
         TabIndex        =   8
         Top             =   2790
         Width           =   2175
      End
      Begin VB.HScrollBar HsrGreen 
         Height          =   225
         Left            =   150
         Max             =   255
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblRed 
         BackStyle       =   0  'Transparent
         Caption         =   "Red:"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblBlue 
         BackStyle       =   0  'Transparent
         Caption         =   "Blue:"
         Height          =   255
         Left            =   570
         TabIndex        =   6
         Top             =   2430
         Width           =   495
      End
      Begin VB.Label lblGreen 
         BackStyle       =   0  'Transparent
         Caption         =   "Green:"
         Height          =   255
         Left            =   570
         TabIndex        =   5
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label lblBlueDegree 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1290
         TabIndex        =   3
         Top             =   2430
         Width           =   375
      End
      Begin VB.Label lblGreenDegree 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1260
         TabIndex        =   2
         Top             =   1500
         Width           =   375
      End
      Begin VB.Label lblRedDegree 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1230
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.Frame fraFilledCircle 
      Height          =   675
      Left            =   450
      TabIndex        =   25
      Top             =   4620
      Width           =   555
      Begin VB.Image imgFilledCircle 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   390
         Left            =   90
         Picture         =   "frmIconEdit.frx":49D4
         ToolTipText     =   "Filled Oval (Shift key if for circle)"
         Top             =   180
         Width           =   375
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5580
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":50D6
            Key             =   "icoBlank"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":53F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":5852
            Key             =   "icoSetUp1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":5CAA
            Key             =   "curHandPoint"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":5FC6
            Key             =   "curArrow8"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "HCL Applications"
      FromPage        =   1
      Max             =   1000
      Min             =   1
      ToPage          =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6180
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   56
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":62EA
            Key             =   "Brushsize1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":6DC2
            Key             =   "Brushsize2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":789A
            Key             =   "Brushsize3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIconEdit.frx":8372
            Key             =   "Brushsize4"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblOffsetXY 
      Alignment       =   1  'Right Justify
      Caption         =   "lblOffsetXY"
      Height          =   195
      Left            =   4020
      TabIndex        =   46
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Label lblIcon 
      BackColor       =   &H80000004&
      Height          =   675
      Left            =   1950
      TabIndex        =   44
      Top             =   150
      Width           =   675
   End
   Begin VB.Label lblFileSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblFileSpec"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   510
      TabIndex        =   24
      Top             =   5610
      Width           =   8055
   End
   Begin VB.Label lblMix 
      Alignment       =   1  'Right Justify
      Caption         =   "lblMix"
      Height          =   195
      Left            =   6960
      TabIndex        =   18
      Top             =   5370
      Width           =   1575
   End
   Begin VB.Label lblBigXY 
      Caption         =   "lblBigXY"
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      Top             =   5370
      Width           =   1365
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsPaint 
         Caption         =   "&Paint"
      End
      Begin VB.Menu mnuToolsRegion 
         Caption         =   "&Region"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsFill 
         Caption         =   "&Fill"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsLine 
         Caption         =   "&Line"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsSquare 
         Caption         =   "&Squire"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsFilledSquare 
         Caption         =   "Filled s&quare"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsCircle 
         Caption         =   "&Circle"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuToolsFilledCircle 
         Caption         =   "Filled C&ircle"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuEditsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditClearClipboard 
         Caption         =   "C&lear clipboard"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "E&ffects"
      Begin VB.Menu mnuEffectsFlipHoriz 
         Caption         =   "Flip &Horizontal"
      End
      Begin VB.Menu mnuEffectsFlipVert 
         Caption         =   "Flip &Verical"
      End
      Begin VB.Menu mnuEffectsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsRotateLeft 
         Caption         =   "Rotate &left"
      End
      Begin VB.Menu mnuEffectsRotateRight 
         Caption         =   "Rotate &right"
      End
      Begin VB.Menu mnuEffectsSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEffectsInvert 
         Caption         =   "&Invert"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsSolidGridlines 
         Caption         =   "&Solid grid lines"
      End
      Begin VB.Menu mnuOptionsDottedGridlines 
         Caption         =   "&Dotted grid lines"
      End
   End
   Begin VB.Menu mnuExtraction 
      Caption         =   "Ext&raction"
      Begin VB.Menu mnuExtractionPanel 
         Caption         =   "Show extraction panel"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmIconEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' IconEdit.frm
'
' By Herman Liu
'
' IconEdit: A basic 32x32 Icon Editor. For creating new icons or editing existing ones,
' with a color palette, a color mixer, a panel for lining up icons for edit, icon new/
' open/save functions, flip/rotate/invert functions and choices of solid/dotted/blank
' grid lines, etc. The use of APIs is reduced to a minumin for purposes of simplicity
' (Other menus not implemented with source code are disabled.  They are left there
' as a ready framework in case some readers may want to build those functions.

Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
    
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
    ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateIconIndirect Lib "user32" (icoinfo As ICONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC, riid As GUID, ByVal fOwn As Long, IPic As IPicture) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, icoinfo As ICONINFO) As Long

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hBMMask As Long
    hBMColor As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
'    data5(16) As Byte                ' Used for LargeAndSmallIcons only
End Type

Private Type PICTDESC
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type IconType
    IndexRef As Integer
    IconPresent As Boolean
    FileName As String
    Loaded As Boolean
    Dirty As Boolean
End Type

Const PICTYPE_BITMAP = 1
Const PICTYPE_ICON = 3
Const PIXELSPERCELL = 8
Const StdW = 32
Const StdH = 32
Const NumLevels = 100
Const ShiftMask = 7

Dim iGuid As GUID
Dim ofIcon(3) As IconType
Dim hDCMono
Dim bmpMono
Dim bmpMonoTemp
Dim CurrIconIndex As Integer

Dim ValReturn As Long
Dim red As Double
Dim Green As Double
Dim blue As Double
Dim BlueSeries As Double
Dim GreenSeries As Double
Dim Celllist() As Long

Dim PicColorToUse_DoubleClicked As Boolean
Dim PaintFlag As Boolean

Dim ColorToUse As Long
Dim Colors()
Dim rotateDirection

Dim BrushSize As Integer
Dim Xicon As Long
Dim Yicon As Long
Dim X1Icon As Single
Dim X2Icon As Single
Dim Y1Icon As Single
Dim Y2Icon As Single

Dim SrcX
Dim SrcY
Dim DestX
Dim DestY

' Various flags
Dim RegionFlag As Boolean
Dim RegionReadyMoveFlag As Boolean
Dim AllowUndoFlag As Boolean
Dim mresult
Dim mfilespec As String
Dim gcancel As Boolean
' Common Dialog control
Dim gcdg As Object
'--------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Me.Move 0, 0
    hDCMono = CreateCompatibleDC(hdc)
    bmpMono = CreateCompatibleBitmap(hDCMono, StdW, StdH)
    bmpMonoTemp = SelectObject(hDCMono, bmpMono)

    With iGuid
         .Data1 = &H20400
         .Data4(0) = &HC0
         .Data4(7) = &H46
    End With
    
    
    Me.ScaleMode = vbPixels
    Me.PaletteMode = vbPaletteModeHalftone
    
    
    picIcon.ScaleMode = vbPixels
    picIcon.AutoRedraw = True
    picIcon.AutoSize = False
    
    
    picMask.ScaleMode = vbPixels
    picMask.AutoRedraw = True
    picMask.AutoSize = False
    picMask.Visible = False
    
    picImage.ScaleMode = vbPixels
    picImage.AutoRedraw = True
    picImage.AutoSize = False
    picImage.Visible = False
    
    picMaskTemp.ScaleMode = vbPixels
    picMaskTemp.AutoRedraw = True
    picMaskTemp.AutoSize = False
    picMaskTemp.Visible = False
    
    picImageTemp.ScaleMode = vbPixels
    picImageTemp.AutoRedraw = True
    picImageTemp.AutoSize = False
    picImageTemp.Visible = False
    
    Dim i
    For i = 0 To 2
        PanelIcons(i).ScaleMode = vbPixels
        PanelIcons(i).AutoRedraw = True
        PanelIcons(i).AutoSize = False
    Next i
    
    picUndo.ScaleMode = vbPixels
    picUndo.AutoRedraw = True
    picUndo.AutoSize = False
    picUndo.Visible = False
    
    picAuto.ScaleMode = vbPixels
    picAuto.AutoRedraw = True
    picAuto.AutoSize = True            ' AutoSize=True for convenience
    picAuto.Visible = False
    
    
    ' -----------
    picBig.ScaleMode = vbPixels
    picBig.AutoRedraw = True
    picBig.AutoSize = False
    picBig.Visible = True
    
    PicGrid.ScaleMode = vbPixels
    PicGrid.Top = picBig.Top
    PicGrid.Left = picBig.Left
    PicGrid.AutoRedraw = True
    PicGrid.AutoSize = False
    PicGrid.Visible = False
    
      ' Fix the size of picBig
    picBig.Width = (picBig.Width - picBig.ScaleWidth) + (StdW * PIXELSPERCELL)
    picBig.Height = (picBig.Height - picBig.ScaleHeight) + (StdH * PIXELSPERCELL)
    
    PicGrid.Width = picBig.Width
    PicGrid.Height = picBig.Height
    
    '----------------------------------------
    picMixGradient.AutoRedraw = True
    PicColorPalette.AutoRedraw = True
    
    HsrRed.Max = 255
    HsrGreen.Max = 255
    HsrBlue.Max = 255
    
    HsrRed.Value = HsrRed.Max
    HsrGreen.Value = HsrGreen.Max
    HsrBlue.Value = HsrBlue.Max / 2
    
    lblRedDegree.Caption = HsrRed.Value
    lblGreenDegree.Caption = HsrGreen.Value
    lblBlueDegree.Caption = HsrBlue.Value
    
    PicRGB(0).BackColor = RGB(HsrRed.Value, 0, 0)
    PicRGB(1).BackColor = RGB(0, HsrGreen.Value, 0)
    PicRGB(2).BackColor = RGB(0, 0, HsrBlue.Value)
    
    ColorMixGradient
    ColorToUse = picMixGradient.BackColor
    picColorToUse.BackColor = ColorToUse
    
    DisplayColorPalette
    
    mnuOptionsSolidGridlines.Checked = False
    mnuOptionsDottedGridlines.Checked = False
        
       ' No gridXY position etc yet
    CleanseCaptions
    lblFileSpec.Caption = "......"
    
       ' Start new values
    SetInitialValues
    
       ' For icons on icons panel
    For i = 0 To 2
         ofIcon(i).IndexRef = i
         ofIcon(i).IconPresent = False
         ofIcon(i).FileName = ""
         ofIcon(i).Loaded = False
         ofIcon(i).Dirty = False
    Next i
    
       ' Load a default icon to icons panel
    PanelIcons(0).Picture = ImageList1.ListImages(2).ExtractIcon
    PanelIcons(0).Visible = True
    lblIcons(0).Visible = True
    CurrIconIndex = 0
    ofIcon(0).IndexRef = 0
    ofIcon(0).IconPresent = True
    ofIcon(0).FileName = "Untitled"
   
    mnuOptionsSolidGridlines.Checked = True
    SetInitialValues
    ofIcon(0).Loaded = False
    imgPaint.Appearance = vbFlat
    mnuToolsPaint.Checked = False
    picBrushSize.Visible = False
    BrushSize = 1                            ' Default once at start only
    Set picBrushSize.Picture = ImageList2.ListImages(1).Picture
    Set gcdg = Me.CommonDialog1
    
    Exit Sub
End Sub
Private Sub Form_Activate()
    ' Avoid blinking effect if focus is on a scrollbar
    picBig.SetFocus
End Sub
Private Sub SetInitialValues()
    gcancel = False
    
    AllowUndoFlag = False
    PaintFlag = False
    
      ' Deselect any tools, then default to imgPaint
    FlatAllToolsImages
    UncheckAllToolsMenus
    EnableToolsMenus True
    imgPaint.Appearance = vb3D
    mnuToolsPaint.Checked = True
    picBrushSize.Visible = True
    
      ' Set starting mouse pointer
    picBig.MousePointer = vbDefault
    
    lblFileSpec.Caption = ofIcon(CurrIconIndex).FileName
    
    ofIcon(CurrIconIndex).Loaded = True
    ofIcon(CurrIconIndex).Dirty = False
    
    CreateGrid
    DoEvents
End Sub
Private Sub CleanseCaptions()
    lblBigXY.Caption = ""
    lblOffsetXY.Caption = ""
    lblMix.Caption = ""
End Sub
Private Sub CreateGrid()
     On Error GoTo Errhandler
     If ofIcon(CurrIconIndex).Loaded = False Then
          Exit Sub
     End If
     If mnuOptionsDottedGridlines.Checked = False And _
           mnuOptionsSolidGridlines.Checked = False Then
          Exit Sub
     End If
     
     Dim w, h
     Dim i, j
    
     w = PicGrid.ScaleWidth
     h = PicGrid.ScaleHeight
    
     PicGrid.Cls
    
      ' Unless specified otherwise, default it to solid lines
     If mnuOptionsDottedGridlines.Checked = False Then
         For i = 1 To 31
             PicGrid.Line (0, i * PIXELSPERCELL)-(w, i * PIXELSPERCELL)
             PicGrid.Line (i * PIXELSPERCELL, 0)-(i * PIXELSPERCELL, h)
         Next i
     Else
         For i = 1 To 31
             For j = 0 To 31
                  PicGrid.PSet ((j + 1) * PIXELSPERCELL, i * PIXELSPERCELL)
             Next j
         Next i
    End If
    Exit Sub
    
Errhandler:
    ErrMsgProc "CreateGrid"
End Sub
Private Sub MagnifyIcon()
    On Error GoTo Errhandler
    
    Dim SrcX As Long, SrcY As Long
    Dim DestX As Long, DestY As Long
    Dim SrcWidth As Long, SrcHeight As Long
    Dim DestWidth As Long, DestHeight As Long
    Dim SrcHDC As Long, DestHDC As Long
      
    picBig.Cls
    picBig.Picture = LoadPicture()
    
    SrcX = 0: SrcY = 0: DestX = 0: DestY = 0
    
    SrcWidth = picIcon.ScaleWidth
    SrcHeight = picIcon.ScaleHeight
    SrcHDC = picIcon.hdc
    
    DestWidth = picBig.ScaleWidth
    DestHeight = picBig.ScaleHeight
    DestHDC = picBig.hdc
    
    mresult = StretchBlt(DestHDC, DestX, DestY, DestWidth, DestHeight, SrcHDC, _
      SrcX, SrcY, SrcWidth, SrcHeight, vbSrcCopy)
    If mresult = 0 Then
        GoTo Errhandler
    End If
    
    If mnuOptionsSolidGridlines.Checked = True Or _
               mnuOptionsDottedGridlines.Checked = True Then
          mresult = BitBlt(picBig.hdc, DestX, DestY, DestWidth, DestHeight, _
                   PicGrid.hdc, SrcX, SrcY, vbSrcAnd)
          If mresult = 0 Then
               GoTo Errhandler
          End If
    End If
         
    picBig.Picture = picBig.Image
    Exit Sub
    
Errhandler:
    ErrMsgProc "MagnifyIcon"
End Sub
Private Sub SetMousePointer(inX, inY)
    If imgPaint.Appearance = vb3D Then picBig.MousePointer = vbUpArrow
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If ofIcon(CurrIconIndex).Loaded = True And ofIcon(CurrIconIndex).Dirty = True Then
        Dim tmp
        tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
             Cancel = True
        ElseIf tmp = vbYes Then
             mnuFileSave_Click
             If gcancel Then
                 Cancel = True
             End If
        End If
    End If
    SelectObject bmpMono, bmpMonoTemp
    DeleteObject bmpMono
    DeleteDC hDCMono
End Sub
Private Sub mnuFile_Click()
    CleanseCaptions
    mnuFileSave.Enabled = ofIcon(CurrIconIndex).Loaded
End Sub
Private Sub mnuEdit_Click()
    CleanseCaptions
    mnuEditUndo.Enabled = ofIcon(CurrIconIndex).Loaded And _
        AllowUndoFlag = True
End Sub
Private Sub mnuEffects_Click()
    CleanseCaptions
    mnuEffectsFlipHoriz.Enabled = ofIcon(CurrIconIndex).Loaded
    mnuEffectsFlipVert.Enabled = ofIcon(CurrIconIndex).Loaded
    mnuEffectsRotateLeft.Enabled = ofIcon(CurrIconIndex).Loaded
    mnuEffectsRotateRight.Enabled = ofIcon(CurrIconIndex).Loaded
    mnuEffectsInvert.Enabled = ofIcon(CurrIconIndex).Loaded
End Sub
Private Sub mnuOptions_Click()
    CleanseCaptions
End Sub
Private Sub PanelIcons_Click(Index As Integer)
    If Index = CurrIconIndex Then
        If ofIcon(CurrIconIndex).Loaded = True Then
             Exit Sub
        End If
    End If
    
    If ofIcon(CurrIconIndex).Loaded Then
        If ofIcon(CurrIconIndex).Loaded = True And ofIcon(CurrIconIndex).Dirty Then
            Dim tmp
            tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)
            If tmp = vbCancel Then
                Exit Sub
            ElseIf tmp = vbYes Then
                mnuFileSave_Click
                If gcancel Then
                    Exit Sub
                End If
            End If
        End If
        ofIcon(CurrIconIndex).Dirty = False
    End If
    
    CurrIconIndex = Index
    Dim i
    For i = 0 To 2
        lblIcons(i).Visible = False
        If i = CurrIconIndex Then
            lblIcons(i).Visible = True
        End If
    Next i
    
    If ofIcon(CurrIconIndex).IconPresent = True Then
         LoadIconToBig
    Else
         picIcon.Cls
         picIcon.Picture = LoadPicture()
         picBig.Cls
         picBig.Picture = LoadPicture()
        
         ofIcon(CurrIconIndex).Loaded = False
         ofIcon(CurrIconIndex).IconPresent = False
         ofIcon(CurrIconIndex).FileName = ""
         ofIcon(CurrIconIndex).Dirty = False
         
         lblFileSpec.Caption = ""
       
         FlatAllToolsImages
         UncheckAllToolsMenus
         EnableToolsMenus False
    End If
End Sub
Private Sub LoadIconToBig()
    ExtractIconInfo PanelIcons(CurrIconIndex)
    picImageAndpicMask picIcon
     ' Start a backup
    doBackUp

    lblFileSpec.Caption = ofIcon(CurrIconIndex).FileName
    SetInitialValues
    MagnifyIcon
End Sub
' Test both if too big and if too small
Private Function TestIconSize(inPic As PictureBox) As Boolean
    TestIconSize = False
    Dim tmp As String
    If inPic.ScaleWidth <> StdW Or inPic.ScaleHeight <> StdH Then
        tmp = "Image is not of " & CStr(StdW) & " x " & _
           CStr(StdH) & vbCrLf & vbCrLf & _
           CStr(inPic.ScaleWidth) & " x " & CStr(inPic.ScaleHeight)
        MsgBox tmp & vbCrLf
        Exit Function
    End If
    TestIconSize = True
End Function
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Set frmIconEdit = Nothing
End Sub
Private Function BitbltPic(SrcPic As Control, DestPic As Control) As Boolean
    On Error GoTo Errhandler
    BitbltPic = True
    
    SrcX = SrcPic.ScaleWidth
    SrcY = SrcPic.ScaleHeight
    
    DestPic.Cls
    DestPic.Picture = LoadPicture()
    mresult = BitBlt(DestPic.hdc, 0, 0, SrcX, SrcY, SrcPic.hdc, 0, 0, vbSrcCopy)
    DestPic.Picture = DestPic.Image
    
    If mresult = 0 Then
        GoTo Errhandler
    End If
    Exit Function
    
Errhandler:
    BitbltPic = False
    ErrMsgProc "BitbltPic"
End Function
Private Sub mnuEffectsFlipHoriz_Click()
    On Error GoTo Errhandler
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    
     ' Do a backup first
    doBackUp
    
    mresult = StretchBlt(picImageTemp.hdc, (StdW - 1), 0, -StdW, StdH, _
                   picImage.hdc, 0, 0, StdW, StdH, vbSrcCopy)
    If mresult = 0 Then
         GoTo Errhandler
    End If
    mresult = StretchBlt(picMaskTemp.hdc, (StdW - 1), 0, -StdW, StdH, _
                   picMask.hdc, 0, 0, StdW, StdH, vbSrcCopy)
                   
    If mresult = 0 Then
         GoTo Errhandler
    End If
    BitbltPic picImageTemp, picImage
    BitbltPic picMaskTemp, picMask
    
    picImageAndpicMask picIcon
    
    MagnifyIcon
    
    Exit Sub
    
Errhandler:
    ErrMsgProc "cmdEditFlipHoriz_Click"
End Sub
Private Sub mnuEffectsFlipVert_Click()
    On Error GoTo Errhandler
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    
    doBackUp
    
    mresult = StretchBlt(picImageTemp.hdc, 0, (StdH - 1), StdW, _
             -StdH, picImage.hdc, 0, 0, StdW, StdH, vbSrcCopy)
    If mresult = 0 Then
         GoTo Errhandler
    End If
    mresult = StretchBlt(picMaskTemp.hdc, 0, (StdH - 1), StdW, _
             -StdH, picMask.hdc, 0, 0, StdW, StdH, vbSrcCopy)
    If mresult = 0 Then
         GoTo Errhandler
    End If
    BitbltPic picImageTemp, picImage
    BitbltPic picMaskTemp, picMask
    picImageAndpicMask picIcon
    MagnifyIcon
    
    Exit Sub
    
Errhandler:
    ErrMsgProc "cmdEditFlipVert_Click"
End Sub
Private Sub mnuEffectsRotateLeft_Click()
    On Error GoTo Errhandler
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    
    doBackUp
    Dim X, Y
    For Y = 0 To (StdW - 1)
         For X = 0 To (StdW - 1)
              picImageTemp.PSet (X, (StdW - 1 - Y)), picImage.Point(Y, X)
              picMaskTemp.PSet (X, (StdW - 1 - Y)), picMask.Point(Y, X)
         Next X
    Next Y
    BitbltPic picImageTemp, picImage
    BitbltPic picMaskTemp, picMask
    
    picImageAndpicMask picIcon
    MagnifyIcon
    Exit Sub
     
Errhandler:
    ErrMsgProc "cmdEditRotateLeft_Click"
End Sub
Private Sub mnuEffectsRotateRight_Click()
    On Error GoTo Errhandler
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    
    doBackUp
    
    Dim X, Y
    For Y = 0 To (StdW - 1)
         For X = 0 To (StdW - 1)
              picImageTemp.PSet (StdW - 1 - Y, X), picImage.Point(X, Y)
              picMaskTemp.PSet (StdW - 1 - Y, X), picMask.Point(X, Y)
         Next X
    Next Y
    BitbltPic picImageTemp, picImage
    BitbltPic picMaskTemp, picMask
    
    picImageAndpicMask picIcon
    MagnifyIcon
    Exit Sub
    
Errhandler:
    ErrMsgProc "cmdEditRotateRight_Click"
End Sub
Private Sub mnuEffectsInvert_Click()
    On Error GoTo Errhandler
    
    doBackUp
    
    picImage.DrawMode = vbInvert
    picImage.Line (0, 0)-(StdW, StdH), , BF
    picImage.DrawMode = vbCopyPen
    picImageAndpicMask picIcon
    MagnifyIcon
    Exit Sub
Errhandler:
    ErrMsgProc "mnuEffectsInvert_click"
End Sub
Private Sub mnuFileNew_Click()
    If ofIcon(CurrIconIndex).Loaded = True And ofIcon(CurrIconIndex).Dirty Then
        Dim tmp
        tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
            Exit Sub
        ElseIf tmp = vbYes Then
            mnuFileSave_Click
            If gcancel Then
                Exit Sub
            End If
        End If
    End If
    
    PanelIcons(CurrIconIndex).Cls
    PanelIcons(CurrIconIndex).Picture = LoadPicture()
    
    ofIcon(CurrIconIndex).Loaded = True
    ofIcon(CurrIconIndex).IconPresent = True
    ofIcon(CurrIconIndex).FileName = "Untitled"
    ofIcon(CurrIconIndex).Dirty = False
    
    picBig.Cls
    picBig.Picture = LoadPicture()
        
    PanelIcons(CurrIconIndex).Picture = ImageList1.ListImages(1).ExtractIcon
    
    LoadIconToBig
End Sub
Private Sub mnuFileOpen_Click()
    On Error GoTo Errhandler
    If ofIcon(CurrIconIndex).Loaded = True And ofIcon(CurrIconIndex).Dirty Then
        Dim tmp
        tmp = MsgBox("Save current icon?", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
            Exit Sub
        ElseIf tmp = vbYes Then
            mnuFileSave_Click
            If gcancel Then
                Exit Sub
            End If
        End If
    End If
    
    gcdg.Filter = "Icon Files (*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.DefaultExt = "ico"
    gcdg.flags = cdlOFNFileMustExist
    
    gcdg.FileName = ""
    gcdg.CancelError = True
    
    gcdg.ShowOpen
    If gcdg.FileName = "" Then
        Exit Sub
    End If
    
    mfilespec = gcdg.FileName
    
    picAuto.Picture = LoadPicture()
    picAuto.Picture = LoadPicture(mfilespec)
    If TestIconSize(picAuto) = False Then
        Exit Sub
    End If
    
    PanelIcons(CurrIconIndex).Cls
    PanelIcons(CurrIconIndex).Picture = LoadPicture()
    PanelIcons(CurrIconIndex).Picture = LoadPicture(mfilespec)
    
    ofIcon(CurrIconIndex).Loaded = True
    ofIcon(CurrIconIndex).IconPresent = True
    ofIcon(CurrIconIndex).FileName = mfilespec
    ofIcon(CurrIconIndex).Dirty = False
    
    LoadIconToBig
    Exit Sub
    
Errhandler:
    If Err <> 32755 Then
         ErrMsgProc "mnuFileOpen_Click"
    End If
End Sub
Private Sub mnuFileSave_Click()
    On Error GoTo Errhandler
    gcancel = True
    
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    
    mfilespec = ofIcon(CurrIconIndex).FileName
    gcdg.FileName = ofIcon(CurrIconIndex).FileName
    gcdg.Filter = "Icon Files (*.ico)|*.ico|(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.DefaultExt = "ico"
    gcdg.flags = cdlOFNOverwritePrompt
    
    gcdg.CancelError = True
    gcdg.ShowSave
    
    mfilespec = gcdg.FileName
    
    SavePicture picIcon.Picture, mfilespec
    ' Update that in icons panel as well
    PanelIcons(CurrIconIndex).Picture = LoadPicture(mfilespec)
    
    ofIcon(CurrIconIndex).FileName = mfilespec
    ofIcon(CurrIconIndex).Dirty = False
    
    SetInitialValues
    LoadIconToBig
    Exit Sub
    
Errhandler:
    If Err <> 32755 Then
          ErrMsgProc "mnuFileSave_Click"
    End If
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub



Private Sub HsrRed_Change()
    CleanseCaptions
    PicRGBMix.BackColor = RGB(HsrRed.Value, HsrGreen.Value, HsrBlue.Value)
    picMixGradient.BackColor = PicRGBMix.BackColor
    lblRedDegree.Caption = HsrRed.Value
    lblGreenDegree.Caption = HsrGreen.Value
    lblBlueDegree.Caption = HsrBlue.Value
    PicRGB(0).BackColor = RGB(HsrRed.Value, 0, 0)
    ColorMixGradient
End Sub
Private Sub HsrGreen_Change()
    CleanseCaptions
    PicRGBMix.BackColor = RGB(HsrRed.Value, HsrGreen.Value, HsrBlue.Value)
    picMixGradient.BackColor = PicRGBMix.BackColor
    lblRedDegree.Caption = HsrRed.Value
    lblGreenDegree.Caption = HsrGreen.Value
    lblBlueDegree.Caption = HsrBlue.Value
    PicRGB(1).BackColor = RGB(0, HsrGreen.Value, 0)
    ColorMixGradient
End Sub
Private Sub HsrBlue_Change()
    CleanseCaptions
    PicRGBMix.BackColor = RGB(HsrRed.Value, HsrGreen.Value, HsrBlue.Value)
    picMixGradient.BackColor = PicRGBMix.BackColor
    lblRedDegree.Caption = HsrRed.Value
    lblGreenDegree.Caption = HsrGreen.Value
    lblBlueDegree.Caption = HsrBlue.Value
    PicRGB(2).BackColor = RGB(0, 0, HsrBlue.Value)
    ColorMixGradient
End Sub
Private Sub mnuOptionsSolidGridLines_Click()
    mnuOptionsSolidGridlines.Checked = Not mnuOptionsSolidGridlines.Checked
    If mnuOptionsSolidGridlines.Checked Then
        If mnuOptionsDottedGridlines.Checked = True Then
             mnuOptionsDottedGridlines.Checked = False
        End If
        If ofIcon(CurrIconIndex).Loaded = True Then
             CreateGrid
             MagnifyIcon
        End If
    Else
        If ofIcon(CurrIconIndex).Loaded = True Then
             CreateGrid
             MagnifyIcon
        End If
    End If
End Sub
Private Sub mnuOptionsDottedGridLines_Click()
    mnuOptionsDottedGridlines.Checked = Not mnuOptionsDottedGridlines.Checked
    If mnuOptionsDottedGridlines.Checked Then
        If mnuOptionsSolidGridlines.Checked = True Then
             mnuOptionsSolidGridlines.Checked = False
        End If
        If ofIcon(CurrIconIndex).Loaded = True Then
             CreateGrid
             MagnifyIcon
        End If
    Else
        If ofIcon(CurrIconIndex).Loaded = True Then
             CreateGrid
             MagnifyIcon
        End If
    End If
End Sub
Private Sub picBrushSize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim d, i
    d = Int(picBrushSize.ScaleHeight / 4)
    i = Int((Y / d))
    If i > 3 Then
        i = 3
    End If
    BrushSize = i + 1
    Set picBrushSize.Picture = ImageList2.ListImages(i + 1).Picture
End Sub
' To gain the use of picColorToUse_MouseUp only
Private Sub picColorToUse_DblClick()
    CleanseCaptions
    PicColorToUse_DoubleClicked = True
End Sub
Private Sub picColorToUse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicColorToUse_DoubleClicked Then
         PicColorToUse_DoubleClicked = False
         Dim tmp
         tmp = picColorToUse.Point(X, Y)
         picColorToUse.BackColor = tmp
         blue = Int((tmp / 256) / 256)
         BlueSeries = (blue * 256) * 256
         Green = Int((tmp - BlueSeries) / 256)
         GreenSeries = Green * 256
         red = Int(tmp - BlueSeries - GreenSeries)
         MsgBox "Currently selected color: " & vbCrLf & vbCrLf & "RGB(" & red & ", " & Green & ", " & blue & ")"
    End If
End Sub
Private Sub picBigContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblBigXY.Caption = ""
    lblOffsetXY.Caption = ""
    PaintFlag = False
    picBig.MousePointer = vbDefault
End Sub
Private Sub PicRGB_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ColorToUse = PicRGB(Index).Point(X, Y)
    picColorToUse.BackColor = ColorToUse
End Sub
Private Sub picRGBMix_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ColorToUse = PicRGBMix.Point(X, Y)
    picColorToUse.BackColor = ColorToUse
End Sub
Private Sub picMixGradient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim w, h
    w = picMixGradient.ScaleWidth
    h = picMixGradient.ScaleHeight
    If (X <= 0) Or (X >= w) Or (Y <= 0) Or (Y > h) Then
         Exit Sub
    End If

    ColorToUse = picMixGradient.Point(X, Y)
    picColorToUse.BackColor = ColorToUse
End Sub
Private Sub pixMixContainer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMix.Caption = ""
End Sub
' ----------------------------------------------------------
' ColorMixGradient() is learnt from Kevino; resulting in a nicer mix
' ----------------------------------------------------------
Private Sub ColorMixGradient()
    On Error Resume Next
    
    Dim R1(1 To NumLevels) As Integer
    Dim G1(1 To NumLevels) As Integer
    Dim B1(1 To NumLevels) As Integer
    
    Dim REnd(1 To 5) As Integer
    Dim GEnd(1 To 5) As Integer
    Dim BEnd(1 To 5) As Integer
    
    Dim i As Integer, End1 As Integer, End2 As Integer, End3 As Integer
    Dim ColorString As String
    Dim Counter As Integer
    
    End1 = 25
    End2 = 50
    End3 = 75
    
    'get RGB values
    For i = 1 To 3
        ColorString = Hex(PicRGB(i - 1).BackColor)
        If Len(ColorString) = 2 Then
            BEnd(i) = 0
            GEnd(i) = 0
            REnd(i) = CLng("&H" & ColorString)
        ElseIf Len(ColorString) = 4 Then
            BEnd(i) = 0
            GEnd(i) = CLng("&H" & Left$(ColorString, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        ElseIf Len(ColorString$) = 6 Then
            BEnd(i) = CLng("&H" & Left$(ColorString, 2))
            GEnd(i) = CLng("&H" & Mid$(ColorString, 3, 2))
            REnd(i) = CLng("&H" & Right$(ColorString, 2))
        End If
    Next i
    
    'Auto calculate mixed "in-between" colors
    For i = 4 To 5 Step 1
        If REnd(i - 2) > REnd(i - 3) Then
            REnd(i) = REnd(i - 2)
        Else
            REnd(i) = REnd(i - 3)
        End If

        If GEnd(i - 2) > GEnd(i - 3) Then
            GEnd(i) = GEnd(i - 2)
        Else
            GEnd(i) = GEnd(i - 3)
        End If

        If BEnd(i - 2) > BEnd(i - 3) Then
            BEnd(i) = BEnd(i - 2)
        Else
            BEnd(i) = BEnd(i - 3)
        End If
    Next i
    
    'set color levels
    For i = 1 To End1
        R1(i) = (i - 1) * (REnd(4) - REnd(1)) / (End1 + 1) + REnd(1)
        G1(i) = (i - 1) * (GEnd(4) - GEnd(1)) / (End1 + 1) + GEnd(1)
        B1(i) = (i - 1) * (BEnd(4) - BEnd(1)) / (End1 + 1) + BEnd(1)
    Next
    Counter = 0

    For i = End1 + 1 To End2
        Counter = Counter + 1
        R1(i) = Counter * (REnd(2) - REnd(4)) / (End2 - End1 + 1) + REnd(4)
        G1(i) = Counter * (GEnd(2) - GEnd(4)) / (End2 - End1 + 1) + GEnd(4)
        B1(i) = Counter * (BEnd(2) - BEnd(4)) / (End2 - End1 + 1) + BEnd(4)
    Next
    Counter = 0

    For i = End2 + 1 To End3
        Counter = Counter + 1
        R1(i) = Counter * (REnd(5) - REnd(2)) / (End3 - End2 + 1) + REnd(2)
        G1(i) = Counter * (GEnd(5) - GEnd(2)) / (End3 - End2 + 1) + GEnd(2)
        B1(i) = Counter * (BEnd(5) - BEnd(2)) / (End3 - End2 + 1) + BEnd(2)
    Next
    Counter = 0

    For i = End3 + 1 To NumLevels
        Counter = Counter + 1
        R1(i) = Counter * (REnd(3) - REnd(5)) / (NumLevels - End3 + 1) + REnd(5)
        G1(i) = Counter * (GEnd(3) - GEnd(5)) / (NumLevels - End3 + 1) + GEnd(5)
        B1(i) = Counter * (BEnd(3) - BEnd(5)) / (NumLevels - End3 + 1) + BEnd(5)
    Next i
    picMixGradient.ScaleHeight = NumLevels
    
    For i = 1 To NumLevels
        picMixGradient.Line (0, i - 1)-(picMixGradient.ScaleWidth, i), RGB(R1(i), G1(i), B1(i)), BF
    Next i
    
    DoEvents
End Sub
Sub DisplayColorPalette()
    PicColorPalette.Scale (0, 0)-(16, 3)
    Dim i
    Colors = Array(16777215, 14737632, 12632319, 12640511, _
                 14745599, 12648384, 16777152, 16761024, _
                 16761087, 192, 16576, 49344, _
                 49152, 12632064, 12582912, 12583104, _
                 12632256, 4210752, 8421631, 8438015, _
                 8454143, 8454016, 16777088, 16744576, _
                 16744703, 128, 16512, 32896, _
                 32768, 8421376, 8388608, 8388736, _
                 8421504, 0, 255, 33023, _
                 65535, 65280, 16776960, 16711680, _
                 16711935, 64, 4210816, 16448, _
                 16384, 4210688, 4194304, 4194368)
    

    For i = 0 To 15
         ' Display a column of colors
        PicColorPalette.Line (i, 0)-(i + 1, 1), Colors(i), BF
        PicColorPalette.Line (i, 1)-(i + 1, 2), Colors(i + 16), BF
        PicColorPalette.Line (i, 2)-(i + 1, 3), Colors(i + 32), BF

        If i > 0 Then
            PicColorPalette.Line (i, 0)-(i, 3)
        End If
    Next i
     ' Horizontal lines dividing the rows.
    PicColorPalette.Line (0, 1)-(16, 1)
    PicColorPalette.Line (0, 2)-(16, 2)
End Sub
Private Sub PicColorPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim w, h
    w = PicColorPalette.ScaleWidth
    h = PicColorPalette.ScaleHeight
    If (X <= 0) Or (X >= w) Or (Y <= 0) Or (Y > h) Then
         Exit Sub
    End If

    Dim c As Long
    c = Fix(X) + Fix(Y) * 16

    ColorToUse = Colors(c)
    picColorToUse.BackColor = ColorToUse
End Sub
Private Sub doBackUp()
    ExtractIconInfo picIcon
    picImageAndpicMask picUndo
    AllowUndoFlag = True
End Sub
Private Sub mnuEditUndo_Click()
    If ofIcon(CurrIconIndex).Loaded = False Then
         MsgBox "No icon loaded yet"
         Exit Sub
    End If
    AllowUndoFlag = False
    RegionFlag = False
    RegionReadyMoveFlag = False
     
       ' Reverse the doBackUp
    ExtractIconInfo picUndo
    picImageAndpicMask picIcon
    
    MagnifyIcon
End Sub
Private Sub mnuEditClear_Click()
    Clipboard.Clear
End Sub
Private Sub cmdNew_Click()
    mnuFileNew_Click
End Sub
Private Sub cmdOpen_Click()
    mnuFileOpen_Click
End Sub
Private Sub cmdSave_Click()
    mnuFileSave_Click
End Sub
' ------Tools menus and images
Private Sub mnuTools_Click()
    If ofIcon(CurrIconIndex).Loaded = False Then
         FlatAllToolsImages
         UncheckAllToolsMenus
         EnableToolsMenus False
    Else
         EnableToolsMenus True
    End If
End Sub
Private Sub mnuToolsPaint_click()
    mnuToolsPaint.Checked = Not mnuToolsPaint.Checked
    SetToolsMenuStatus mnuToolsPaint, imgPaint
End Sub
Private Sub imgPaint_Click()
    mnuToolsPaint_click
End Sub
Private Sub SetToolsMenuStatus(inMenu As Menu, inImage As Image)
    If ofIcon(CurrIconIndex).Loaded = False Then
         FlatAllToolsImages
         UncheckAllToolsMenus
         EnableToolsMenus False
         Exit Sub
    End If
    If inMenu.Checked = True Then
         UncheckAllToolsMenus
         FlatAllToolsImages
         inMenu.Checked = True
         inImage.Appearance = vb3D
    Else
         inMenu.Checked = False
         inImage.Appearance = vbFlat
    End If
    picBrushSize.Visible = mnuToolsPaint.Checked
      ' If mnuFill, we may need to fill region
    If inMenu <> mnuToolsFill Then
         picBig.Cls
    End If
End Sub
Private Sub FlatAllToolsImages()
    imgPaint.Appearance = vbFlat
    picBrushSize.Visible = False
End Sub
Private Sub UncheckAllToolsMenus()
    mnuToolsPaint.Checked = False
    picBrushSize.Visible = False
End Sub
Private Sub EnableToolsMenus(Onoff As Boolean)
    mnuToolsPaint.Enabled = Onoff
End Sub
Private Sub PicBig_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEvents
    If Button <> vbLeftButton Then
        Exit Sub
    End If
    
    If ofIcon(CurrIconIndex).Loaded = False Then
        Exit Sub
    End If

    If imgPaint.Appearance = vbFlat Then
         Exit Sub
    End If
    
    Xicon = Fix(X / PIXELSPERCELL)       ' In picIcon pixel unit now
    Yicon = Fix(Y / PIXELSPERCELL)
    
    PaintFlag = True
         
    doBackUp
    
    X1Icon = Xicon: X2Icon = Xicon: Y1Icon = Yicon: Y2Icon = Yicon
    If BrushSize = 1 Then
        picImage.PSet (Xicon, Yicon), ColorToUse
        picMask.PSet (Xicon, Yicon)
    Else
        X2Icon = Xicon + (BrushSize - 1)
        Y2Icon = Yicon + (BrushSize - 1)
        If X2Icon > (StdW - 1) Then
            X2Icon = StdW - 1
        End If
        If Y2Icon > (StdH - 1) Then
            Y2Icon = StdH - 1
        End If
        picImage.Line (X1Icon, Y1Icon)-(X2Icon, Y2Icon), ColorToUse, BF
        picMask.Line (X1Icon, Y1Icon)-(X2Icon, Y2Icon), , BF
    End If
    picImageAndpicMask picIcon
    MagnifyIcon
End Sub
Private Sub picBig_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ofIcon(CurrIconIndex).Loaded = False Then
        Exit Sub
    End If

    Dim w, h
    w = picBig.ScaleWidth
    h = picBig.ScaleHeight
    If (X <= 0) Or (X >= w) Or (Y <= 0) Or (Y > h) Then
        If RegionReadyMoveFlag Then
             Exit Sub
        End If
    End If
    
      ' Here to effect SetMousePointer
    SetMousePointer X, Y
    
    If ofIcon(CurrIconIndex).Loaded = False Then
         Exit Sub
    End If
    
    Xicon = Fix(X / PIXELSPERCELL)
    Yicon = Fix(Y / PIXELSPERCELL)
    
      ' Display current Big (X,Y) position
    lblBigXY.Caption = "X=" & CStr(Xicon + 1) & "  Y=" & CStr(Yicon + 1)
    
    If imgPaint.Appearance = vbFlat Then
         Exit Sub
    End If
    
    If PaintFlag = False Then
         Exit Sub
    End If
    
    X2Icon = Xicon: Y2Icon = Yicon
    If BrushSize = 1 Then
         ' May also use Line method
        picImage.PSet (Xicon, Yicon), ColorToUse
        picMask.PSet (Xicon, Yicon)
    Else
        X1Icon = Xicon
        Y1Icon = Yicon
        X2Icon = Xicon + (BrushSize - 1)
        Y2Icon = Yicon + (BrushSize - 1)
        If X2Icon > StdW Then
            X2Icon = StdW
        End If
        If Y2Icon > StdH Then
            Y2Icon = StdH
        End If
        picImage.Line (X1Icon, Y1Icon)-(X2Icon, Y2Icon), ColorToUse, BF
        picMask.Line (X1Icon, Y1Icon)-(X2Icon, Y2Icon), , BF
    End If
    picImageAndpicMask picIcon
    MagnifyIcon
End Sub
Private Sub picBig_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ofIcon(CurrIconIndex).Loaded = False Then
        Exit Sub
    End If

    If imgPaint.Appearance = vbFlat Then
       Exit Sub
    End If
    
    ofIcon(CurrIconIndex).Dirty = True
    
    PaintFlag = False
    Exit Sub
End Sub
Private Sub ExtractIconInfo(inPic As PictureBox)
    Dim IPic As IPicture
    Dim icoinfo As ICONINFO
    Dim PDesc As PICTDESC
    Dim hDCWork
    Dim hBMOldWork
    Dim hNewBM
    Dim hBMOldMono
    
    GetIconInfo inPic.Picture, icoinfo
    hDCWork = CreateCompatibleDC(0)
    hNewBM = CreateCompatibleBitmap(Me.hdc, StdW, StdH)
    hBMOldWork = SelectObject(hDCWork, hNewBM)
    hBMOldMono = SelectObject(hDCMono, icoinfo.hBMMask)
    BitBlt hDCWork, 0, 0, StdW, StdH, hDCMono, 0, 0, vbSrcCopy
    SelectObject hDCMono, hBMOldMono
    SelectObject hDCWork, hBMOldWork
    
    With PDesc
        .cbSizeofStruct = Len(PDesc)
        .picType = PICTYPE_BITMAP
        .hImage = hNewBM
    End With
    
    OleCreatePictureIndirect PDesc, iGuid, 1, IPic
    
    picMask = IPic
    Set IPic = Nothing
    
    PDesc.hImage = icoinfo.hBMColor
    OleCreatePictureIndirect PDesc, iGuid, 1, IPic
    picImage = IPic
    
    DeleteObject icoinfo.hBMMask
    DeleteDC hDCWork
    Set hBMOldWork = Nothing
    Set hBMOldMono = Nothing
End Sub
' Update picIcon with its image and mask
Sub picImageAndpicMask(inPic As PictureBox)
    Dim hOldMonoBM
    Dim hDCWork
    Dim hBMOldWork
    Dim hBMWork
    Dim PDesc As PICTDESC
    Dim icoinfo As ICONINFO
    Dim IPic As IPicture

    BitBlt hDCMono, 0, 0, StdW, StdH, Me.picMask.hdc, 0, 0, vbSrcCopy
    
    SelectObject hDCMono, bmpMonoTemp
    
    hDCWork = CreateCompatibleDC(0)
    
    With inPic
        hBMWork = CreateCompatibleBitmap(Me.hdc, StdW, StdH)
    End With
    
    hBMOldWork = SelectObject(hDCWork, hBMWork)
    
    BitBlt hDCWork, 0, 0, StdW, StdH, Me.picImage.hdc, 0, 0, vbSrcCopy
    
    SelectObject hDCWork, hBMOldWork
    
    With icoinfo
        .fIcon = 1
        .xHotspot = 16                         ' Icon hot spot in middle
        .yHotspot = 16
        .hBMMask = bmpMono
        .hBMColor = hBMWork
    End With
    
    With PDesc
        .cbSizeofStruct = Len(PDesc)
        .picType = PICTYPE_ICON
        .hImage = CreateIconIndirect(icoinfo)
    End With
    
    OleCreatePictureIndirect PDesc, iGuid, 1, IPic
    
    inPic.Picture = LoadPicture()
    inPic = IPic
    
    bmpMonoTemp = SelectObject(hDCMono, bmpMono)
    
    DeleteObject icoinfo.hBMMask
    DeleteDC hDCWork
    Set hBMOldWork = Nothing
End Sub
Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub
