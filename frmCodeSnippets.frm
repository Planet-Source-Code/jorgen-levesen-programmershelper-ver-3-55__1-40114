VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeSnippets 
   BackColor       =   &H00404040&
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmCodeSnippets.frx":0000
   ScaleHeight     =   8865
   ScaleWidth      =   11475
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      Picture         =   "frmCodeSnippets.frx":0376
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data rsCodeType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeSnippets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeType"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CodeText"
      DataSource      =   "rsCodeSnippet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   50
      TabIndex        =   24
      Top             =   960
      Width           =   7575
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Author"
      DataSource      =   "rsCodeSnippet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   70
      TabIndex        =   22
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "&Find"
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
      Left            =   6120
      TabIndex        =   21
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "AuthorMail"
      DataSource      =   "rsCodeSnippet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      MaxLength       =   100
      TabIndex        =   19
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton btnChangeDatabase 
      Caption         =   "&Change Database"
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
      Left            =   8760
      Picture         =   "frmCodeSnippets.frx":06EC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "DateInDatabase"
      DataSource      =   "rsCodeSnippet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   70
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "AuthorUrl"
      DataSource      =   "rsCodeSnippet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      MaxLength       =   100
      TabIndex        =   12
      Top             =   2280
      Width           =   4815
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4905
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton btnStatistic 
      Caption         =   "C&ode Statistic"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   8040
      Width           =   1455
   End
   Begin VB.ComboBox cmbCodeType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.ComboBox cmbCodeLanguage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Data rsCodeSnippet 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\For WebPage\CodeSnippets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeSnippet"
      Top             =   -120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "CodeSnippet"
      DataSource      =   "rsCodeSnippet"
      Height          =   5415
      Left            =   3600
      TabIndex        =   10
      Top             =   3120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9551
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCodeSnippets.frx":0836
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   41
      X1              =   8640
      X2              =   8640
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   40
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   39
      X1              =   11280
      X2              =   11280
      Y1              =   2880
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   38
      X1              =   3480
      X2              =   3480
      Y1              =   2880
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   37
      X1              =   3360
      X2              =   3360
      Y1              =   2880
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   36
      X1              =   120
      X2              =   120
      Y1              =   2880
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   35
      X1              =   11280
      X2              =   11280
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   34
      X1              =   11280
      X2              =   11280
      Y1              =   1440
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   33
      X1              =   11280
      X2              =   11280
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   32
      X1              =   8520
      X2              =   8520
      Y1              =   2160
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   31
      X1              =   3480
      X2              =   3480
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   30
      X1              =   3480
      X2              =   3480
      Y1              =   1440
      Y2              =   1920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   29
      X1              =   3480
      X2              =   3480
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   28
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   27
      X1              =   3360
      X2              =   3360
      Y1              =   1920
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   26
      X1              =   120
      X2              =   120
      Y1              =   1920
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   25
      X1              =   120
      X2              =   120
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   24
      X1              =   3360
      X2              =   3360
      Y1              =   840
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   23
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   22
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Text:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3600
      TabIndex        =   25
      Top             =   720
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   21
      X1              =   3480
      X2              =   11280
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   20
      X1              =   4680
      X2              =   11280
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   3600
      TabIndex        =   23
      Top             =   1320
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   19
      X1              =   4440
      X2              =   6720
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   18
      X1              =   3480
      X2              =   6720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Author Email Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6960
      TabIndex        =   20
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   17
      X1              =   8640
      X2              =   11280
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   16
      X1              =   6840
      X2              =   11280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   18
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   15
      X1              =   3480
      X2              =   11280
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   14
      X1              =   4680
      X2              =   11280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Date stored in Database:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   8760
      TabIndex        =   15
      Top             =   2040
      Width           =   1770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   13
      X1              =   8640
      X2              =   11280
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   12
      X1              =   10680
      X2              =   11280
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Author Internet Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   3600
      TabIndex        =   13
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   11
      X1              =   3480
      X2              =   8520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   5520
      X2              =   8520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   3600
      TabIndex        =   11
      Top             =   2760
      Width           =   420
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   9
      X1              =   3480
      X2              =   11280
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   4320
      X2              =   11280
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Text:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   120
      X2              =   3360
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   1080
      X2              =   3360
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   120
      X2              =   3360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1200
      X2              =   3360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Code Language:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   3360
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1560
      X2              =   3360
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Records in Database:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1920
      X2              =   3360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   3360
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmCodeSnippets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodeBook() As Variant, bNewRecord As Boolean, iList1Index As Integer
Dim rsLanguage As Recordset
Dim rsCodeLanguage As Recordset
Dim rsUser As Recordset
Dim boolOwnCode As Boolean, boolWrite As Boolean
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 10
        Label3(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 41
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub


Private Sub LoadCodeLanguage()
    cmbCodeLanguage.Clear
    With rsCodeLanguage
        .MoveFirst
        .Index = "PrimaryKey"
        Do While Not .EOF
            cmbCodeLanguage.AddItem Trim(.Fields("Language"))
        .MoveNext
        Loop
    End With
    If Not IsNull(rsUser.Fields("PrefferedLanguage")) Then
        cmbCodeLanguage.Text = rsUser.Fields("PrefferedLanguage")
    Else
        cmbCodeLanguage.ListIndex = 0
    End If
End Sub
Private Sub ReadText()
Dim sHelp As String
    On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("Label3(2)")) Then
                    .Fields("Label3(2)") = Label3(2).Caption
                Else
                    Label3(2).Caption = .Fields("Label3(2)")
                End If
                If IsNull(.Fields("Label3(3)")) Then
                    .Fields("Label3(3)") = Label3(3).Caption
                Else
                    Label3(3).Caption = .Fields("Label3(3)")
                End If
                If IsNull(.Fields("Label3(4)")) Then
                    .Fields("Label3(4)") = Label3(4).Caption
                Else
                    Label3(4).Caption = .Fields("Label3(4)")
                End If
                If IsNull(.Fields("Label3(10)")) Then
                    .Fields("Label3(10)") = Label3(10).Caption
                Else
                    Label3(10).Caption = .Fields("Label3(10)")
                End If
                If IsNull(.Fields("Label3(9)")) Then
                    .Fields("Label3(9)") = Label3(9).Caption
                Else
                    Label3(9).Caption = .Fields("Label3(9)")
                End If
                If IsNull(.Fields("Label3(8)")) Then
                    .Fields("Label3(8)") = Label3(8).Caption
                Else
                    Label3(8).Caption = .Fields("Label3(8)")
                End If
                If IsNull(.Fields("Label3(5)")) Then
                    .Fields("Label3(5)") = Label3(5).Caption
                Else
                    Label3(5).Caption = .Fields("Label3(5)")
                End If
                If IsNull(.Fields("Label3(6)")) Then
                    .Fields("Label3(6)") = Label3(6).Caption
                Else
                    Label3(6).Caption = .Fields("Label3(6)")
                End If
                If IsNull(.Fields("Label3(1)")) Then
                    .Fields("Label3(1)") = Label3(1).Caption
                Else
                    Label3(1).Caption = .Fields("Label3(1)")
                End If
                If IsNull(.Fields("Label3(0)")) Then
                    .Fields("Label3(0)") = Label3(0).Caption
                Else
                    Label3(0).Caption = .Fields("Label3(0)")
                End If
                If IsNull(.Fields("btnChangeDatabase")) Then
                    .Fields("btnChangeDatabase") = btnChangeDatabase.Caption
                Else
                    btnChangeDatabase.Caption = .Fields("btnChangeDatabase")
                End If
                
                If IsNull(.Fields("btnStatistic")) Then
                    .Fields("btnStatistic") = btnStatistic.Caption
                Else
                    btnStatistic.Caption = .Fields("btnStatistic")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        'this language was not found, make it. Find the English text first
        sHelp = " "
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then
                If Not IsNull(.Fields("Help")) Then
                    sHelp = .Fields("Help")
                    Exit Do
                End If
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Label3(2)") = Label3(2).Caption
        .Fields("Label3(3)") = Label3(3).Caption
        .Fields("Label3(4)") = Label3(4).Caption
        .Fields("Label3(10)") = Label3(10).Caption
        .Fields("Label3(9)") = Label3(9).Caption
        .Fields("Label3(8)") = Label3(8).Caption
        .Fields("Label3(5)") = Label3(5).Caption
        .Fields("Label3(6)") = Label3(6).Caption
        .Fields("Label3(7)") = Label3(7).Caption
        .Fields("Label3(1)") = Label3(1).Caption
        .Fields("Label3(0)") = Label3(0).Caption
        .Fields("btnChangeDatabase") = btnChangeDatabase.Caption
        .Fields("btnStatistic") = btnStatistic.Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    boolWrite = True
    With rsCodeSnippet.Recordset
        ReDim vCodeBook(0 To .RecordCount) As Variant
        For i = 0 To .RecordCount - 1
            List1.AddItem .Fields("CodeText")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vCodeBook(List1.ListCount - 1) = .Bookmark
            .MoveNext
        Next
    End With
    boolWrite = False
End Sub


Private Sub SelectAll()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeSnippet"
    With rsCodeSnippet
        .RecordSource = Sql
        .Refresh
        If Not .Recordset.EOF And Not .Recordset.BOF Then
            .Recordset.MoveFirst
            .Recordset.MoveLast
            Label1(0).Caption = "Records: " & .Recordset.RecordCount
            Label1(0).ForeColor = rsUser.Fields("LabelColor")
            .Recordset.MoveFirst
        Else
            Label1(0).Caption = "Records: " & 0
            Label1(0).ForeColor = rsUser.Fields("LabelColor")
        End If
    End With
End Sub

Public Sub SelectAllCode()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeSnippet"
    rsCodeSnippet.RecordSource = Sql
    rsCodeSnippet.Refresh
End Sub


Private Sub SelectCodeType()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeType WHERE Trim(CodeLanguage) ="
    Sql = Sql & Chr(34) & Trim(cmbCodeLanguage.Text) & Chr(34)
    
    With rsCodeType
        .RecordSource = Sql
        .Refresh
        If Not .Recordset.EOF And Not .Recordset.BOF Then
            .Recordset.MoveFirst
        End If
    End With
End Sub

Public Sub SelectRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeSnippet WHERE Trim(CodeType) ="
    Sql = Sql & Chr(34) & Trim(cmbCodeType.Text) & Chr(34)
    Sql = Sql & " AND Trim(CodeLanguage) ="
    Sql = Sql & Chr(34) & Trim(cmbCodeLanguage.Text) & Chr(34)
    Sql = Sql & " ORDER BY CodeText"
    
    With rsCodeSnippet
        .RecordSource = Sql
        .Refresh
        If Not .Recordset.EOF And Not .Recordset.BOF Then
            .Recordset.MoveFirst
            .Recordset.MoveLast
            Label1(1).Caption = "Records: " & .Recordset.RecordCount
            Label1(1).ForeColor = rsUser.Fields("LabelColor")
            .Recordset.MoveFirst
        Else
            Label1(1).Caption = "Records: " & 0
            Label1(1).ForeColor = rsUser.Fields("LabelColor")
        End If
    End With
End Sub

Private Sub LoadCodeType()
    On Error Resume Next
    SelectCodeType
    cmbCodeType.Clear
    With rsCodeType.Recordset
        .MoveFirst
        .Index = "PrimaryKey"
        Do While Not .EOF
            cmbCodeType.AddItem Trim(.Fields("CodeType"))
        .MoveNext
        Loop
    End With
End Sub
Public Sub CopySnippToClip()
    On Error Resume Next
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Clipboard.Clear
    Clipboard.SetText Text1.Text
End Sub

Public Sub DeleteRecord()
Dim DgDef, Msg, response, Title
    If bNewRecord Then Exit Sub
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    Title = "Delete Record"
    Msg = "Do you really want to delete this Code Snippet ?"
    Beep
    response = MsgBox(Msg, DgDef, Title)
    If response = IdNo Then
        Exit Sub
    End If
    On Error Resume Next
    rsCodeSnippet.Recordset.Delete
    List1.RemoveItem (iList1Index)
    List1.ListIndex = 0
End Sub

Public Sub NewRecord()
    On Error Resume Next
    If Len(cmbCodeType.Text) = 0 Then Exit Sub
    rsCodeSnippet.Recordset.AddNew
    bNewRecord = True
    Text2.SetFocus
End Sub

Public Sub PrintRecord()
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    cPrint.FontSize = 12
    cPrint.FontBold = True
    cPrint.pBox , , , 1.1, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pPrint Label3(10).Caption, 0.2, True
    cPrint.FontBold = False
    cPrint.pPrint Text2.Text, 2.4, False 'code text
    cPrint.FontBold = True
    cPrint.pPrint Label3(9).Caption, 0.2, True
    cPrint.FontBold = False
    cPrint.pPrint Text3.Text, 2.4, False 'author
    cPrint.FontBold = True
    cPrint.pPrint Label3(8).Caption, 0.2, True
    cPrint.FontBold = False
    cPrint.pPrint Text4.Text, 2.4, False 'author mail address
    cPrint.FontBold = True
    cPrint.pPrint Label3(5).Caption, 0.2, True
    cPrint.FontBold = False
    cPrint.pPrint Text5.Text, 2.4, False 'internet address
    cPrint.FontBold = True
    cPrint.pPrint Label3(6).Caption, 0.2, True
    cPrint.FontBold = False
    cPrint.pPrint Text6.Text, 2.4, False 'date added to database
    cPrint.pPrint
    cPrint.FontSize = 10
    cPrint.pDoubleLine
    cPrint.pPrint
    cPrint.BackColor = -1
    cPrint.FontBold = True
    cPrint.pPrint Label3(4).Caption & ":", 0.2
    cPrint.FontBold = False
    cPrint.pMultiline Text1.Text, , , , False, True
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Public Sub SelectRemote()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeSnippet WHERE Trim(CodeType) ="
    Sql = Sql & Chr(34) & Trim(cmbCodeType.Text) & Chr(34)
    rsCodeSnippet.RecordSource = Sql
    rsCodeSnippet.Refresh
End Sub

Public Sub ShowAuthor()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeSnippet WHERE CLng(CodeNo) ="
    Sql = Sql & Chr(34) & CLng(m_lSnippet) & Chr(34)
    rsCodeSnippet.RecordSource = Sql
    rsCodeSnippet.Refresh
End Sub

Private Sub btnChangeDatabase_Click()
    On Error Resume Next
    If boolOwnCode Then
        rsCodeSnippet.DatabaseName = m_strCodeSnippet
        rsCodeSnippet.Refresh
        rsCodeType.DatabaseName = m_strCodeSnippet
        rsCodeType.Refresh
        SelectAll
        Label2.Caption = ExtractFileName(m_strPrograming)
        LoadCodeType
        DoEvents
        cmbCodeType.ListIndex = 0
        cmbCodeType_Click
        List1.ListIndex = 0
        btnSearch.Enabled = True
        boolOwnCode = False
    Else
        rsCodeSnippet.DatabaseName = m_strMyCodeSnippet
        rsCodeType.DatabaseName = m_strMyCodeSnippet
        rsCodeSnippet.Refresh
        rsCodeType.Refresh
        SelectAll
        Label2.Caption = ExtractFileName(m_strMyCodeSnippet)
        LoadCodeType
        DoEvents
        cmbCodeType.ListIndex = 0
        cmbCodeType_Click
        List1.ListIndex = 0
        btnSearch.Enabled = False
        boolOwnCode = True
    End If
End Sub

Private Sub btnSearch_Click()
    With frmShowAuthors
        .Text1.Text = Text3.Text
        .Show vbModal
    End With
End Sub
Private Sub btnStatistic_Click()
    m_boolSnippet = True
    frmCodeStatistic.Show 1
End Sub

Private Sub cmbCodeLanguage_Click()
    On Error Resume Next
    LoadCodeType
    cmbCodeType.ListIndex = 0
End Sub

Private Sub cmbCodeType_Click()
    On Error Resume Next
    SelectRecords
    LoadList1
    List1.ListIndex = 0
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsCodeSnippet.Refresh
    rsCodeType.Refresh
    
    With rsCodeSnippet.Recordset
        .MoveLast
        .MoveFirst
        Label1(0).Caption = .RecordCount
        Label1(0).ForeColor = rsUser.Fields("LabelColor")
    End With
    LoadCodeLanguage
    LoadCodeType
    cmbCodeType.ListIndex = 0
    List1.ListIndex = 0
    ReadText
    DisableButtons 2
    
    frmMDI.Toolbar1.Buttons(15).Enabled = True
    frmMDI.Toolbar1.Buttons(16).Enabled = True
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
Dim sName As String
    On Error GoTo errForm_Load
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    
    If CBool(rsUser.Fields("OwnSnippetAsDefault")) Then
        rsCodeSnippet.DatabaseName = m_strMyCodeSnippet  'first choice
        rsCodeType.DatabaseName = m_strMyCodeSnippet
        Label2.Caption = ExtractFileName(m_strMyCodeSnippet)
        boolOwnCode = True
    Else
        rsCodeSnippet.DatabaseName = m_strCodeSnippet
        rsCodeType.DatabaseName = m_strCodeSnippet
        Set rsCodeLanguage = m_dbCodeSnippet.OpenRecordset("Language")
        Label2.Caption = ExtractFileName(m_strCodeSnippet)
        boolOwnCode = False
    End If
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCodeSnippets")
    m_iFormNo = 2
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Form Load"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
    LoadBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCodeSnippet.Recordset.Close
    rsCodeType.Recordset.Close
    rsCodeLanguage.Close
    rsLanguage.Close
    rsUser.Close
    m_iFormNo = 0
    DisableButtons 1
    frmMDI.Toolbar1.Buttons(15).Enabled = False
    frmMDI.Toolbar1.Buttons(16).Enabled = False
    Set frmCodeSnippets = Nothing
End Sub


Private Sub List1_Click()
    On Error Resume Next
    If boolWrite Then Exit Sub
    iList1Index = List1.ListIndex
    rsCodeSnippet.Recordset.Bookmark = vCodeBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ItemHeight As Long
Dim NewIndex As Long
    Static OldIndex As Long
    On Error Resume Next
    With List1
       ItemHeight = SendMessage(.hWnd, LB_GETITEMHEIGHT, 0, ByVal 0&)
       ItemHeight = .Parent.ScaleY(ItemHeight, vbPixels, vbTwips)
       NewIndex = .TopIndex + (Y \ ItemHeight)
       If NewIndex <> OldIndex Then
          If NewIndex < .ListCount Then
             .ToolTipText = .List(NewIndex)
          Else
             .ToolTipText = vbNullString
          End If
          OldIndex = NewIndex
      End If
    End With
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        Text1.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      frmMDI.PopupMenu frmMDI.mnuFormat
   End If
End Sub

Private Sub Text2_LostFocus()
    On Error Resume Next
    If bNewRecord Then
        With rsCodeSnippet.Recordset
            .Fields("CodeLanguage") = cmbCodeLanguage.Text
            .Fields("CodeType") = cmbCodeType.Text
            .Fields("CodeText") = Trim(Text2.Text)
            .Fields("DateInDatabase") = Format(Now, "dd.mm.yyyy")
            .Update
            LoadList1
            .Bookmark = .LastModified
            bNewRecord = False
        End With
    End If
End Sub
