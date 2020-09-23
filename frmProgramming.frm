VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProgramming 
   BackColor       =   &H00404040&
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10860
   ControlBox      =   0   'False
   Icon            =   "frmProgramming.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   10860
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   6855
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   36
      Top             =   360
      Width           =   1815
   End
   Begin VB.Data rsProgramming 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProjectTime"
      Top             =   7080
      Visible         =   0   'False
      Width           =   1140
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   7215
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Project Times"
      TabPicture(0)   =   "frmProgramming.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Wish List"
      TabPicture(1)   =   "frmProgramming.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   8175
         Begin VB.CommandButton btnUpdate2 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   0
            Picture         =   "frmProgramming.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "Update wish list"
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox cmbCustomer 
            BackColor       =   &H00FFFFC0&
            DataField       =   "WishedByCustomerName"
            DataSource      =   "rsWishList"
            Height          =   315
            Left            =   3240
            Sorted          =   -1  'True
            TabIndex        =   31
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H80000004&
            Caption         =   "Change Done"
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   1320
            TabIndex        =   22
            Top             =   1680
            Width           =   3855
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00C0C0C0&
               Height          =   495
               Left            =   120
               Picture         =   "frmProgramming.frx":0706
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   240
               Width           =   615
            End
            Begin MSMask.MaskEdBox MaskEdBox1 
               Height          =   375
               Index           =   4
               Left            =   2400
               TabIndex        =   23
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777152
               PromptInclude   =   0   'False
               AllowPrompt     =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   10
               Mask            =   "##.##.####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Change done Date:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   720
               TabIndex        =   24
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3240
            MaxLength       =   50
            TabIndex        =   18
            Top             =   360
            Width           =   1935
         End
         Begin VB.ListBox List5 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5880
            Left            =   6960
            TabIndex        =   17
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   3855
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   2640
            Width           =   5055
         End
         Begin VB.ListBox List4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5880
            Left            =   5280
            TabIndex        =   15
            Top             =   480
            Width           =   1695
         End
         Begin VB.Data rsWishList 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   300
            Left            =   1080
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "WishList"
            Top             =   6360
            Visible         =   0   'False
            Width           =   1140
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   3
            Left            =   3240
            TabIndex        =   19
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Mask            =   "##.##.####"
            PromptChar      =   "_"
         End
         Begin VB.PictureBox Picture1 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   33
            Top             =   0
            Width           =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Wished by Customer:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   1200
            TabIndex        =   32
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   6960
            TabIndex        =   27
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Project"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Input Date:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   1560
            TabIndex        =   21
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ProjectName:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   1560
            TabIndex        =   20
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   8175
         Begin VB.CommandButton btnUpdate1 
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            Picture         =   "frmProgramming.frx":0A10
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Update project"
            Top             =   840
            Width           =   1215
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H8000000A&
            Caption         =   "Total Time Used:"
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   5400
            TabIndex        =   28
            Top             =   5760
            Width           =   2775
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Hours"
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
               Left            =   2040
               TabIndex        =   30
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Left            =   360
               TabIndex        =   29
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.Timer Timer1 
            Interval        =   60
            Left            =   1320
            Top             =   120
         End
         Begin VB.TextBox TextBox1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   4095
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   2280
            Width           =   4935
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   3120
            MaxLength       =   50
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000A&
            Caption         =   "Project Dates"
            ForeColor       =   &H00000000&
            Height          =   5535
            Left            =   5400
            TabIndex        =   2
            Top             =   120
            Width           =   2775
            Begin VB.ListBox List3 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   5100
               Index           =   0
               Left            =   1440
               TabIndex        =   4
               Top             =   240
               Width           =   1095
            End
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   5100
               Index           =   0
               Left            =   120
               TabIndex        =   3
               ToolTipText     =   "Click to show the record"
               Top             =   240
               Width           =   1335
            End
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   0
            Left            =   3120
            TabIndex        =   7
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   10
            Mask            =   "##.##.####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   1
            Left            =   3120
            TabIndex        =   8
            Top             =   1200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   375
            Index           =   2
            Left            =   3120
            TabIndex        =   9
            Top             =   1680
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777152
            PromptInclude   =   0   'False
            AllowPrompt     =   -1  'True
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Time To:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   1440
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Time From:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   12
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Project Date:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   11
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ProjectName:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projects:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   0
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   120
      X2              =   2160
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   960
      X2              =   2160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2160
      X2              =   2160
      Y1              =   120
      Y2              =   7440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   7440
   End
End
Attribute VB_Name = "frmProgramming"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer, i As Integer, bookmarksPro() As Variant, bookmarksWish() As Variant
Dim bNewRecord As Boolean, bFirst As Boolean
Dim vCustomBook() As Variant
Dim bookmarks1() As Variant
Dim dSumTime As Double, vTime As Variant
Dim lngFormWidth As Long, lngFormHeight As Long
Dim dbTemp As Database
Dim rsUser As Recordset
Dim rsCustomer As Recordset
Dim rsLanguage As Recordset
Dim rsProjects As Recordset
Private Sub LoadBackground()
    Picture2.Visible = False
    Picture2.AutoRedraw = True
    Picture2.AutoSize = True
    Picture2.BorderStyle = 0
    Picture2.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture2
    Label4.ForeColor = rsUser.Fields("LabelColor")
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
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
                For i = 0 To 9
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4.Caption
                Else
                    Label4.Caption = .Fields("label4")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("Frame5")) Then
                    .Fields("Frame5") = Frame5.Caption
                Else
                    Frame5.Caption = .Fields("Frame5")
                End If
                If IsNull(.Fields("Frame6")) Then
                    .Fields("Frame6") = Frame6.Caption
                Else
                    Frame6.Caption = .Fields("Frame6")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                If IsNull(.Fields("btnUpdate1")) Then
                    .Fields("btnUpdate1") = btnUpdate1.Caption
                Else
                    btnUpdate1.Caption = .Fields("btnUpdate1")
                End If
                If IsNull(.Fields("btnUpdate2")) Then
                    .Fields("btnUpdate2") = btnUpdate2.Caption
                Else
                    btnUpdate2.Caption = .Fields("btnUpdate2")
                End If
                Tab1.Tab = 0
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
        For i = 0 To 9
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("label4") = Label4.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame5") = Frame5.Caption
        .Fields("Frame6") = Frame6.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        .Fields("btnUpdate1") = btnUpdate1.Caption
        .Fields("btnUpdate2") = btnUpdate2.Caption
        .Fields("Help") = sHelp
        .Update
        Tab1.Tab = 0
    End With
End Sub

Public Sub DeletePrograming()
    On Error Resume Next
    rsProgramming.Recordset.Delete
    LoadList2
End Sub


Public Sub DeleteWish()
    On Error Resume Next
    rsWishList.Recordset.Delete
    LoadList4
End Sub

Private Sub LoadcmbCustomer()
    On Error Resume Next
    cmbCustomer.Clear
    With rsCustomer
        .MoveLast
        .MoveFirst
        ReDim vCustomBook(.RecordCount)
        Do While Not .EOF
            cmbCustomer.AddItem .Fields("CustomerName")
            cmbCustomer.ItemData(cmbCustomer.NewIndex) = cmbCustomer.ListCount - 1
            vCustomBook(cmbCustomer.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub ClearFields2()
    MaskEdBox1(3).Text = " "
    MaskEdBox1(4).Text = " "
    Text2.Text = " "
    Text3.Text = " "
    TextBox1.Text = " "
    Check1.Value = 0
End Sub


Private Sub LoadList4()
    On Error Resume Next
    If bNewRecord Then Exit Sub
    List4.Clear
    List5.Clear
    ClearFields2
    With rsWishList.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarksWish(.RecordCount)
        Do While Not .EOF
            If Not CBool(.Fields("ChangeDone")) Then
                List4.AddItem .Fields("ProjectID")
                If IsDate(.Fields("RegisterDate")) Then
                    List5.AddItem .Fields("RegisterDate")
                Else
                    List5.AddItem "0"
                End If
                List4.ItemData(List4.NewIndex) = List4.ListCount - 1
                bookmarksWish(List4.ListCount - 1) = .Bookmark
            End If
        .MoveNext
        Loop
    End With
End Sub

Private Sub ClearFields()
    MaskEdBox1(0).Text = " "
    MaskEdBox1(1).Text = " "
    MaskEdBox1(2).Text = " "
    TextBox1.Text = " "
    Label2.Caption = " "
End Sub


Private Sub LoadList1()
    On Error Resume Next
    If bNewRecord Then Exit Sub
    With rsProjects
        .MoveLast
        .MoveFirst
        ReDim bookmarks1(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("ProjectID")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarks1(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub LoadList2()
    On Error Resume Next
    If bNewRecord Then Exit Sub
    List2(0).Clear
    List3(0).Clear
    ClearFields
    dSumTime = 0
    With rsProgramming.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarksPro(.RecordCount)
        Do While Not .EOF
            List2(0).AddItem .Fields("Date")
            List3(0).AddItem .Fields("TimeFrom")
            List2(0).ItemData(List2(0).NewIndex) = List2(0).ListCount - 1
            bookmarksPro(List2(0).ListCount - 1) = .Bookmark
            vTime = DateDiff("n", CDate(.Fields("TimeFrom")), CDate(.Fields("TimeTo")))
            dSumTime = dSumTime + vTime
        .MoveNext
        Loop
    End With
    dSumTime = dSumTime / 60    'calculate in hours
    Label2.Caption = Format(dSumTime, "0.00")
End Sub

Public Sub NewPrograming()
    ClearFields
    bNewRecord = True
    MaskEdBox1(0).SetFocus
End Sub

Public Sub NewWish()
    ClearFields2
    bNewRecord = True
    Text3.Text = List1.List(List1.ListIndex)
    MaskEdBox1(3).SetFocus
End Sub

Private Sub SelectProject()
Dim Sql As String
    On Error Resume Next
    'find all the project records
    Sql = "SELECT * FROM ProjectTime WHERE Trim(ProjectID) ="
    Sql = Sql & Chr(34) & Trim(rsProjects.Fields("ProjectID")) & Chr(34)
    Sql = Sql & "ORDER BY Date"
    rsProgramming.RecordSource = Sql
    rsProgramming.Refresh
End Sub

Private Sub ShowRecord()
    On Error Resume Next
    With rsProgramming.Recordset
        Frame2.Caption = .Fields("ProjectID")
        Text1.Text = .Fields("ProjectID")
        MaskEdBox1(0).Text = CDate(.Fields("Date"))
         MaskEdBox1(1).Text = .Fields("TimeFrom")
        MaskEdBox1(2).Text = .Fields("TimeTo")
        If Not IsNull(.Fields("DoneWhat")) Then
            TextBox1.Text = .Fields("DoneWhat")
        End If
    End With
End Sub

Private Sub ShowWish()
    On Error Resume Next
    ClearFields2
    With rsWishList.Recordset
        Text3.Text = .Fields("ProjectID")
        MaskEdBox1(3).Text = CDate(.Fields("RegisterDate"))
        Text2.Text = .Fields("ChangeWish")
    End With
End Sub

Public Sub WriteAllProjects()
Dim sProject As String, dblSumLine As Double, boolFirstRead As Boolean
Dim Sql As String, boolWrite As Boolean, sString As String
Dim sFromDate As String, sToDate As String, sDate As String
    'On Error Resume Next
    
    'find all the project records
    Sql = "SELECT * FROM ProjectTime ORDER BY ProjectID, Date"
    rsProgramming.RecordSource = Sql
    rsProgramming.Refresh
    
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter3:
    Screen.MousePointer = vbHourglass
    DoEvents
    cPrint.pStartDoc
    
    dSumTime = 0
    sProject = " "
    dblSumLine = 0
    boolFirstRead = True
    
    WriteHead
    
    With rsProgramming.Recordset
        .MoveFirst
        Do While Not .EOF
            frmWriteProgRep.Text1(2).Text = CLng(frmWriteProgRep.Text1(2).Text) + 1
            DoEvents
            
            'only print records within from- and to-dates
            If IsDateBetween(CDate(dateFromDate), CDate(dateToDate), CDate(.Fields("Date"))) Then
                If .Fields("ProjectID") = sProject Then
                    cPrint.pPrint Format(CDate(.Fields("Date")), "dd.mm.yyyy"), 0.3, True
                    cPrint.pPrint Format(CDate(.Fields("TimeFrom")), "hh:mm"), 1.5, True
                    cPrint.pPrint Format(CDate(.Fields("TimeTo")), "hh:mm"), 2.5, True
                    cPrint.pMultiline cPrint.GetRemoveCRLF(.Fields("DoneWhat")), 3.5, , , , True
                    
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        WriteHead
                    End If
                    
                    boolWrite = True
                    frmWriteProgRep.Text1(3).Text = CLng(frmWriteProgRep.Text1(3).Text) + 1
                    DoEvents
                    
                    vTime = DateDiff("n", CDate(.Fields("TimeFrom")), CDate(.Fields("TimeTo")))
                    dSumTime = dSumTime + vTime
                    dblSumLine = dblSumLine + vTime
                Else
                    If boolFirstRead Then
                        sProject = .Fields("ProjectID")
                        cPrint.FontBold = True
                        cPrint.pPrint sProject, 0.3
                        cPrint.FontBold = False
                        
                        cPrint.pPrint Format(CDate(.Fields("Date")), "dd.mm.yyyy"), 0.3, True   'programming date
                        cPrint.pPrint Format(CDate(.Fields("TimeFrom")), "hh:mm"), 1.5, True
                        cPrint.pPrint Format(CDate(.Fields("TimeTo")), "hh:mm"), 2.5, True
                        cPrint.pMultiline cPrint.GetRemoveCRLF(.Fields("DoneWhat")), 3.5, , , , True
                        
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            WriteHead
                        End If
                        
                        boolWrite = True
                        frmWriteProgRep.Text1(3).Text = CLng(frmWriteProgRep.Text1(3).Text) + 1
                        DoEvents
                        vTime = DateDiff("n", CDate(.Fields("TimeFrom")), CDate(.Fields("TimeTo")))
                        dSumTime = vTime
                        dblSumLine = vTime
                        boolFirstRead = False
                    Else
                        If boolWrite Then
                            dblSumLine = dblSumLine / 60    'calculate in hours
                            cPrint.pPrint "Sum time Project:", 0.3, True
                            cPrint.FontBold = True
                            cPrint.pPrint Format(dblSumLine, "###,###.00"), 1.5
                            cPrint.FontBold = False
                            cPrint.pPrint
                            boolWrite = False
                        End If
                        
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            WriteHead
                        End If
                        
                        sProject = Trim(.Fields("ProjectID"))
                        vTime = DateDiff("n", CDate(.Fields("TimeFrom")), CDate(.Fields("TimeTo")))
                        dSumTime = dSumTime + vTime
                        dblSumLine = vTime
                        
                        cPrint.FontBold = True
                        cPrint.pPrint sProject, 0.3
                        cPrint.FontBold = False
                        
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            WriteHead
                        End If
                        
                        cPrint.pPrint Format(CDate(.Fields("Date")), "dd.mm.yyyy"), 0.3, True   'programming date
                        cPrint.pPrint Format(CDate(.Fields("TimeFrom")), "hh:mm"), 1.5, True
                        cPrint.pPrint Format(CDate(.Fields("TimeTo")), "hh:mm"), 2.5, True
                        cPrint.pMultiline cPrint.GetRemoveCRLF(.Fields("DoneWhat")), 3.5, , , , True
                        
                        boolWrite = True
                        frmWriteProgRep.Text1(3).Text = CLng(frmWriteProgRep.Text1(3).Text) + 1
                        DoEvents
                        
                        If cPrint.pEndOfPage Then
                            cPrint.pFooter
                            cPrint.pNewPage
                            WriteHead
                        End If
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If boolWrite Then
        'write the last sum time line
        dblSumLine = dblSumLine / 60    'calculate in hours
        cPrint.pPrint
        cPrint.pPrint "Sum time Project:", 0.3, True
        cPrint.FontBold = True
        cPrint.pPrint Format(dblSumLine, "###,###.00"), 1.5
        cPrint.FontBold = False
        
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            WriteHead
        End If
    End If
    
    'write the sum times this report
    dSumTime = dSumTime / 60    'calculate in hours
    cPrint.pPrint
    cPrint.pPrint "Sum total time:", 0.3, True
    cPrint.FontBold = True
    cPrint.pPrint Format(dSumTime, "###,###.00"), 1.5
    cPrint.FontBold = False
    
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter3
    Set cPrint = Nothing
End Sub

Private Sub WriteHead()
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 18
    cPrint.pBox 0.25, cPrint.CurrentY, cPrint.GetPaperWidth - 0.25, 0.7, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pCenter List1.List(List1.ListIndex) & " - " & Frame2.Caption
    cPrint.BackColor = -1
    cPrint.FontSize = 10
    cPrint.pPrint "Time", 1.5, True
    cPrint.pPrint "Time", 2.5
    cPrint.pPrint "Date", 0.3, True
    cPrint.pPrint "From", 1.5, True
    cPrint.pPrint "To", 2.5, True
    cPrint.pPrint "Done What", 3.5
    cPrint.pLine 0.3
    cPrint.FontBold = False
    cPrint.pPrint
End Sub

Private Sub WriteHeadWish()
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 18
    cPrint.pBox 0.25, cPrint.CurrentY, cPrint.GetPaperWidth - 0.25, 0.7, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pCenter "Project wish list"
    cPrint.BackColor = -1
    cPrint.FontSize = 10
    cPrint.pPrint "Date", 1.5, True
    cPrint.pPrint "Pri-", 2.5
    cPrint.pPrint "Project", 0.3, True
    cPrint.pPrint "Registred", 1.5, True
    cPrint.pPrint "ority", 2.5, True
    cPrint.pPrint "Change wish", 3.5
    cPrint.pLine 0.3
    cPrint.FontBold = False
    cPrint.pPrint
End Sub

Public Sub WriteProject()
    On Error Resume Next
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
    dSumTime = 0
    WriteHead
    With rsProgramming.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("ProjectID") = List1.List(List1.ListIndex) Then
                cPrint.pPrint Format(CDate(.Fields("Date")), "dd.mm.yyyy"), 0.3, True   'programming date
                cPrint.pPrint Format(CDate(.Fields("TimeFrom")), "hh:mm"), 1.5, True
                cPrint.pPrint Format(CDate(.Fields("TimeTo")), "hh:mm"), 2.5, True
                cPrint.pMultiline cPrint.GetRemoveCRLF(.Fields("DoneWhat")), 3.5, , , , True
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    WriteHead
                End If
            End If
            vTime = DateDiff("n", CDate(.Fields("TimeFrom")), CDate(.Fields("TimeTo")))
            dSumTime = dSumTime + vTime
        .MoveNext
        Loop
    End With
    
    'write sum time line
    dSumTime = dSumTime / 60    'calculate in hours
    cPrint.pPrint
    cPrint.pPrint "Sum time:", 0.3, True
    cPrint.pPrint Format(dSumTime, "###,###.00"), 1.5
    
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Public Sub WriteWishList()
    On Error Resume Next
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter2:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    WriteHeadWish
    With rsWishList.Recordset
        .MoveFirst
        Do While Not .EOF
            If Not CBool(.Fields("ChangeDone")) Then
                cPrint.pPrint
                cPrint.FontBold = True
                cPrint.pPrint .Fields("ProjectID"), 0.3
                cPrint.FontBold = False
                cPrint.pPrint CDate(.Fields("RegisterDate")), 1.5, True
                cPrint.pPrint .Fields("Priority"), 2.5, True
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    WriteHeadWish
                End If
                cPrint.pMultiline cPrint.GetRemoveCRLF(.Fields("ChangeWish")), 3.5, , , , True
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    WriteHeadWish
                End If
            End If
        .MoveNext
        Loop
    End With
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter2
    Set cPrint = Nothing
End Sub

Private Sub btnUpdate1_Click()
    On Error GoTo errUpdate
    With rsProgramming.Recordset
        If bNewRecord Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("ProjectID") = CStr(Text1.Text)
        .Fields("Date") = CDate(MaskEdBox1(0).FormattedText)
        .Fields("TimeFrom") = MaskEdBox1(1).FormattedText
        .Fields("TimeTo") = MaskEdBox1(2).FormattedText
        If Len(TextBox1.Text) <> 0 Then
            .Fields("DoneWhat") = TextBox1.Text
        End If
        .Update
        .Bookmark = .LastModified
    End With
    
    bNewRecord = False
    LoadList2
    Exit Sub
    
errUpdate:
    Beep
    MsgBox Err.Description, vbCritical, "Record Update"
    Resume errUpdate2
errUpdate2:
End Sub

Private Sub btnUpdate2_Click()
    On Error Resume Next
    With rsWishList.Recordset
        If bNewRecord Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("ProjectID") = CStr(Text3.Text)
        .Fields("RegisterDate") = CDate(MaskEdBox1(3).FormattedText)
        .Fields("ChangeWish") = Text2.Text
        If Check1.Value <> 0 Then
            .Fields("ChangeDone") = True
            If IsDate(MaskEdBox1(4).FormattedText) Then
                .Fields("ChangeDoneDate") = MaskEdBox1(4).FormattedText
            End If
        End If
        .Update
        .Bookmark = .LastModified
    End With
    
    bNewRecord = False
    LoadList4
End Sub

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        MaskEdBox1(4).Text = " "
        MaskEdBox1(4).Enabled = False
    Else
        MaskEdBox1(4).Enabled = True
        MaskEdBox1(4).SetFocus
    End If
End Sub

Private Sub cmbCustomer_Click()
    On Error Resume Next
    rsCustomer.Bookmark = vCustomBook(cmbCustomer.ItemData(cmbCustomer.ListIndex))
    With rsWishList.Recordset
        .Edit
        .Fields("WishedByCustomerCounter") = CLng(rsCustomer.Fields("AutoLine"))
        .Update
    End With
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    If Not bFirst Then Exit Sub
    rsProgramming.Refresh
    rsWishList.Refresh
    LoadList1
    List1.ListIndex = 0
    LoadcmbCustomer
    ReadText
    DisableButtons 2
    frmMDI.Toolbar1.Buttons(8).Enabled = False
    frmMDI.Toolbar1.Buttons(9).Enabled = False
    bFirst = False
    DBEngine.Idle dbFreeLocks
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsProgramming.DatabaseName = m_strPrograming
    rsWishList.DatabaseName = m_strPrograming
    Set rsProjects = m_dbPrograming.OpenRecordset("Projects")
    Set rsCustomer = m_dbPrograming.OpenRecordset("Customer")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmProgramming")
    bFirst = True
    m_iFormNo = 18
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
    LoadBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsProgramming.UpdateRecord
    rsProgramming.Recordset.Close
    rsWishList.Recordset.Close
    rsProjects.Close
    rsCustomer.Close
    rsUser.Close
    rsLanguage.Close
    dbTemp.Close
    m_iFormNo = 0
    DisableButtons 1
    Erase bookmarksPro
    Erase bookmarksWish
    Erase bookmarks1
    Erase vCustomBook
    Set frmProgramming = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    If bNewRecord Then Exit Sub
    rsProjects.Bookmark = bookmarks1(List1.ItemData(List1.ListIndex))
    SelectProject
    List2(0).Clear
    List3(0).Clear
    LoadList2
    Frame2.Caption = rsProjects.Fields("ProjectText")
    Text1.Text = rsProjects.Fields("ProjectID")
    Text3.Text = rsProjects.Fields("ProjectID")
End Sub

Private Sub List2_Click(Index As Integer)
    On Error Resume Next
    If bNewRecord Then Exit Sub
    n = List2(0).ListIndex
    List3(0).ListIndex = n
    rsProgramming.Recordset.Bookmark = bookmarksPro(List2(0).ItemData(List2(0).ListIndex))
    ShowRecord
End Sub

Private Sub List4_Click()
    On Error Resume Next
    n = List4.ListIndex
    List5.ListIndex = n
    If bNewRecord Then Exit Sub
    rsWishList.Recordset.Bookmark = bookmarksWish(List4.ItemData(List4.ListIndex))
    ShowWish
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        LoadList4
    End If
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    List3(0).TopIndex = List2(0).TopIndex
    List5.TopIndex = List4.TopIndex
End Sub
