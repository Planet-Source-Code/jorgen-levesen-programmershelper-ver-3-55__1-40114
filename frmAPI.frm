VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAPI 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox cboAPIType 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   1320
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   26
      Top             =   600
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5295
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   2415
   End
   Begin TabDlg.SSTab Tab1 
      Height          =   8175
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "API"
      TabPicture(0)   =   "frmAPI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Example"
      TabPicture(1)   =   "frmAPI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   7575
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   7575
         Begin VB.Frame Frame9 
            ForeColor       =   &H00000000&
            Height          =   975
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   7335
            Begin VB.Data rsAPI 
               Caption         =   "Data1"
               Connect         =   "Access 2000;"
               DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeSnippets.mdb"
               DefaultCursorType=   0  'DefaultCursor
               DefaultType     =   2  'UseODBC
               Exclusive       =   0   'False
               Height          =   345
               Left            =   120
               Options         =   0
               ReadOnly        =   0   'False
               RecordsetType   =   1  'Dynaset
               RecordSource    =   "API"
               Top             =   600
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.ComboBox cboAPIType 
               BackColor       =   &H00FFFFC0&
               DataField       =   "APIType"
               DataSource      =   "rsAPI"
               Height          =   315
               Index           =   1
               Left            =   3240
               TabIndex        =   17
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "APIName"
               DataSource      =   "rsAPI"
               Height          =   285
               Index           =   6
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   16
               Top             =   240
               Width           =   2295
            End
            Begin VB.Label DateInDatabase 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               BorderStyle     =   1  'Fixed Single
               DataField       =   "DateInDatabase"
               DataSource      =   "rsAPI"
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   6120
               TabIndex        =   20
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "API Type:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   1440
               TabIndex        =   19
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "API Name:"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   1440
               TabIndex        =   18
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Author Information"
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
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   7335
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "AuthorInternet"
               DataSource      =   "rsAPI"
               Height          =   285
               Index           =   5
               Left            =   3240
               LinkTimeout     =   100
               TabIndex        =   11
               Top             =   720
               Width           =   3975
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "AuthorEmail"
               DataSource      =   "rsAPI"
               Height          =   285
               Index           =   4
               Left            =   3240
               LinkTimeout     =   70
               TabIndex        =   10
               Top             =   480
               Width           =   3975
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "AuthorName"
               DataSource      =   "rsAPI"
               Height          =   285
               Index           =   3
               Left            =   3240
               MaxLength       =   50
               TabIndex        =   9
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Author Internet:"
               Height          =   255
               Index           =   2
               Left            =   1440
               TabIndex        =   14
               Top             =   720
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Author Email:"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   13
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Author Name:"
               Height          =   255
               Index           =   0
               Left            =   1440
               TabIndex        =   12
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Parameters"
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
            Height          =   3015
            Left            =   120
            TabIndex        =   6
            Top             =   4440
            Width           =   7335
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "APIParameters"
               DataSource      =   "rsAPI"
               Height          =   2655
               Index           =   1
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   7
               Top             =   240
               Width           =   7095
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Explanation"
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
            Height          =   1935
            Left            =   120
            TabIndex        =   4
            Top             =   2400
            Width           =   7335
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "APIExplanation"
               DataSource      =   "rsAPI"
               Height          =   1575
               Index           =   0
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   5
               Top             =   240
               Width           =   7095
            End
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Example"
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
         Height          =   7455
         Left            =   -74760
         TabIndex        =   1
         Top             =   480
         Width           =   7455
         Begin RichTextLib.RichTextBox RichText1 
            DataField       =   "APIExample"
            DataSource      =   "rsAPI"
            Height          =   7095
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   12515
            _Version        =   393217
            BackColor       =   16777152
            Enabled         =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmAPI.frx":0038
         End
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   3120
      X2              =   240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   3120
      X2              =   240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   3120
      X2              =   240
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   3120
      X2              =   3120
      Y1              =   240
      Y2              =   8400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   8400
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Alpha sorted"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   840
      TabIndex        =   30
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Show API sorted"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   840
      TabIndex        =   29
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "API Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   28
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stored API's"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   24
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   23
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   1800
      Width           =   2655
   End
End
Attribute VB_Name = "frmAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolWrite As Boolean, vCodeBook() As Variant, bNewRecord As Boolean
Dim boolFirst As Boolean
Dim rsAPIType As Recordset
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 5
        Label2(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 4
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
                For i = 0 To 4
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
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
                If IsNull(.Fields("Frame7")) Then
                    .Fields("Frame7") = Frame7.Caption
                Else
                    Frame7.Caption = .Fields("Frame7")
                End If
                If IsNull(.Fields("Frame8")) Then
                    .Fields("Frame8") = Frame8.Caption
                Else
                    Frame8.Caption = .Fields("Frame8")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab1(0)")) Then
                    .Fields("Tab1(0)") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab1(0)")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab1(1)")) Then
                    .Fields("Tab1(1)") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab1(1)")
                End If
                If IsNull(.Fields("label2(3)")) Then
                    .Fields("label2(3)") = Label2(3).Caption
                Else
                    Label2(3).Caption = .Fields("label2(3)")
                End If
                If IsNull(.Fields("label2(4)")) Then
                    .Fields("label2(4)") = Label2(4).Caption
                Else
                    Label2(4).Caption = .Fields("label2(4)")
                End If
                If IsNull(.Fields("label2(5)")) Then
                    .Fields("label2(5)") = Label2(5).Caption
                Else
                    Label2(5).Caption = .Fields("label2(5)")
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
        For i = 0 To 5
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("Frame5") = Frame5.Caption
        .Fields("Frame6") = Frame6.Caption
        .Fields("Frame7") = Frame7.Caption
        .Fields("Frame8") = Frame8.Caption
        Tab1.Tab = 0
        .Fields("Tab1(0)") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab1(1)") = Tab1.Caption
        Tab1.Tab = 0
        .Fields("label2(3)") = Label2(3).Caption
        .Fields("label2(4)") = Label2(4).Caption
        .Fields("label2(5)") = Label2(5).Caption
        .Fields("Help") = sHelp
        .Update
    End With
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
    cPrint.pPrint Label1(5).Caption, 0.2, True
    cPrint.pPrint cboAPIType(1).Text, 2.4, False 'api type
    cPrint.pPrint Label1(4).Caption, 0.2, True
    cPrint.pPrint Text1(6).Text, 2.4, False 'api name
    cPrint.pPrint Label1(1).Caption, 0.2, True
    cPrint.pPrint Text1(3).Text, 2.4, False 'author name
    cPrint.pPrint Label1(2).Caption, 0.2, True
    cPrint.pPrint Text1(4).Text, 2.4, False 'author mail
    cPrint.pPrint Label1(3).Caption, 0.2, True
    cPrint.pPrint Text1(5).Text, 2.4, False 'author internet
    cPrint.FontBold = False
    cPrint.pPrint
    cPrint.BackColor = &HFFFFFF
    cPrint.pPrint Frame5.Caption, 0.2, True
    cPrint.pMultiline Text1(0).Text, 2.4, , , , True 'api explanation
    cPrint.pPrint
    cPrint.pPrint Frame6.Caption, 0.2, True
    cPrint.pMultiline Text1(1).Text, 2.4, , , , True 'api parameter
    cPrint.pPrint
    cPrint.pPrint Frame7.Caption, 0.2, True
    cPrint.pMultiline RichText1.Text, 2.4, , , , True 'api explanation
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Public Sub SelectAllAPI()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM API"
    rsAPI.RecordSource = Sql
    rsAPI.Refresh
    With rsAPI.Recordset
        .MoveLast
        .MoveFirst
        Label2(0).Caption = "Records: " & .RecordCount
    End With
End Sub
Public Sub SelectRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM API WHERE Trim(APIType) ="
    Sql = Sql & Chr(34) & Trim(cboAPIType(0).Text) & Chr(34)
    
    With rsAPI
        .RecordSource = Sql
        .Refresh
        If Not .Recordset.EOF And Not .Recordset.BOF Then
            .Recordset.MoveFirst
            .Recordset.MoveLast
            Label2(1).Caption = "Records: " & .Recordset.RecordCount
            .Recordset.MoveFirst
        Else
            Label2(1).Caption = "Records: " & 0
        End If
    End With
End Sub

Public Sub NewRecord()
    On Error Resume Next
    If Len(cboAPIType(0).Text) = 0 Then Exit Sub
    Tab1.Tab = 0
    rsAPI.Recordset.AddNew
    cboAPIType(1).Text = cboAPIType(0).Text
    bNewRecord = True
    Text1(6).SetFocus
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
    rsAPI.Recordset.Delete
    List1.RemoveItem (List1.List(List1.ListIndex))
    List1.ListIndex = 0
End Sub
Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    boolWrite = True
    With rsAPI.Recordset
        ReDim vCodeBook(0 To .RecordCount) As Variant
        For i = 0 To .RecordCount - 1
            List1.AddItem .Fields("APIName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vCodeBook(List1.ListCount - 1) = .Bookmark
            .MoveNext
        Next
    End With
    boolWrite = False
End Sub
Private Sub LoadAPIType()
    On Error Resume Next
    cboAPIType(0).Clear
    cboAPIType(1).Clear
    With rsAPIType
        .MoveFirst
        .Index = "PrimaryKey"
        Do While Not .EOF
            cboAPIType(0).AddItem .Fields("APIType")
            cboAPIType(1).AddItem .Fields("APIType")
        .MoveNext
        Loop
    End With
End Sub

Private Sub SelectSortedAPI()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM API"
    Sql = Sql & " ORDER BY APIName"
    
    With rsAPI
        .RecordSource = Sql
        .Refresh
        If Not .Recordset.EOF And Not .Recordset.BOF Then
            .Recordset.MoveFirst
            .Recordset.MoveLast
            Label2(1).Caption = "Records: " & .Recordset.RecordCount
            .Recordset.MoveFirst
        Else
            Label2(1).Caption = "Records: " & 0
        End If
    End With
End Sub
Private Sub cboAPIType_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        SelectRecords
        LoadList1
        List1.ListIndex = 0
    Case Else
    End Select
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    If Not boolFirst Then Exit Sub
    rsAPI.Refresh
    SelectAllAPI
    LoadAPIType
    cboAPIType(0).ListIndex = 0
    SelectRecords
    LoadList1
    List1.ListIndex = 0
    ReadText
    DisableButtons 2
    
    frmMDI.Toolbar1.Buttons(15).Enabled = True
    frmMDI.Toolbar1.Buttons(16).Enabled = True
    Me.WindowState = vbMaximized
    boolFirst = False
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsAPI.DatabaseName = m_strCodeSnippet
    Set rsAPIType = m_dbCodeSnippet.OpenRecordset("APIType")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmAPI")
    boolFirst = True
    m_iFormNo = 37
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbExclamation, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
    LoadBackground
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsAPIType.Close
    rsAPI.Recordset.Close
    rsUser.Close
    rsLanguage.Close
    Erase vCodeBook
    m_iFormNo = 0
    
    DisableButtons 1
    frmMDI.Toolbar1.Buttons(15).Enabled = False
    frmMDI.Toolbar1.Buttons(16).Enabled = False
    Set frmAPI = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    If boolWrite Then Exit Sub
    rsAPI.Recordset.Bookmark = vCodeBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Option1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0  'API type
        Label1(0).Visible = True
        cboAPIType(0).Visible = True
        DoEvents
        SelectRecords
        LoadList1
        List1.ListIndex = 0
    Case 1  'sorted alphanumerical
        Label1(0).Visible = False
        cboAPIType(0).Visible = False
        DoEvents
        SelectSortedAPI
        LoadList1
        List1.ListIndex = 0
    Case Else
    End Select
End Sub

Private Sub RichText1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichText1.SelText = vbTab
        KeyCode = 0
    End If
End Sub


Private Sub RichText1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      frmMDI.PopupMenu frmMDI.mnuFormat
   End If
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 6
    If bNewRecord Then
        With rsAPI.Recordset
            .Fields("APIName") = Text1(6).Text
            .Fields("DateInDatabase") = Format(Now, "dd.mm.yyyy")
            .Update
            LoadList1
            .Bookmark = .LastModified
            bNewRecord = False
        End With
    End If
    Case Else
    End Select
End Sub


