VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About..."
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "About Project2"
   Begin VB.CommandButton cmdOK 
      Height          =   375
      Left            =   240
      Picture         =   "frmAbout.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   5640
      Width           =   6255
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "About ...."
      TabPicture(0)   =   "frmAbout.frx":058C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTitle"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblVersion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "picIcon"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Credits"
      TabPicture(1)   =   "frmAbout.frx":05A8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label3(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label3(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3(6)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label3(7)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3(8)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label3(9)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label3(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label3(0)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.PictureBox picIcon 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   720
         Picture         =   "frmAbout.frx":05C4
         ScaleHeight     =   480
         ScaleMode       =   0  'User
         ScaleWidth      =   480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The Print Preview code was written by Morgan Haueisen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   20
         Top             =   4200
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The Database HTML print was written by Joseph B. Surls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The VB Code Spell Checker was written by Greg DeBacker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -74760
         TabIndex        =   17
         Top             =   3840
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The Picture Viewer was written by: www.vb-helper.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   16
         Top             =   3480
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The Project Statistic was written by Eric O'Sullivan "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   15
         Top             =   3120
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The Database Print-routine was written by Joseph B. Surls"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   -74640
         TabIndex        =   14
         Top             =   2760
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The ""Disable MDI Close button""-routine was written by Sean Dittmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   -74880
         TabIndex        =   13
         Top             =   2280
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Thanks to the above - as well as numerous others who have helped to make this system possible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Index           =   4
         Left            =   -74760
         TabIndex        =   12
         Top             =   4680
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "I have used many a programmers various code snippets and  components, I have left the original codes as were if appliable."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   735
         Index           =   3
         Left            =   -74760
         TabIndex        =   11
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The form-resizing is done by unknown author."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   10
         Top             =   1920
         Width           =   5775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "The left hand vertical-menu is programed by Yves Lessard , SevySoft (I think)."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   -74880
         TabIndex        =   9
         Top             =   1440
         Width           =   5775
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Version: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Tag             =   "Version"
         Top             =   1200
         Width           =   3885
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Application Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1680
         TabIndex        =   7
         Tag             =   "Application Title"
         Top             =   960
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "This program is created by Jørgen E. Levesen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   6
         Top             =   1800
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail: jorgen@levesen.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   2040
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "URL: http//www.levesen.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   2280
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Please feel free to use the source code at your wish"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   3
         Top             =   2880
         Width           =   5175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.levesen.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   3240
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    On Error Resume Next
    lblVersion.Caption = lblVersion.Caption & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    m_iFormNo = 1
    Dither Me
    DisableButtons 1
End Sub
Private Sub cmdOK_Click()
        Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_iFormNo = 0
    DisableButtons 2
    Set frmAbout = Nothing
End Sub

Private Sub Label1_Click(Index As Integer)
Dim iRet As Long
    On Error Resume Next
    Select Case Index
    Case 1
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            "jorgen@levesen.com", _
            vbNullString, _
            "c:\", _
            SW_SHOWNORMAL)
    Case 2
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            "http://www.levesen.com", _
            vbNullString, _
            "c:\", _
            SW_SHOWNORMAL)
    Case Else
    End Select
End Sub

Private Sub Label2_Click()
Dim X As Long
    X = Shell("explorer http://www.levesen.com")
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &H80FF80
End Sub

Private Sub Tab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label2.ForeColor = &H0&
End Sub
