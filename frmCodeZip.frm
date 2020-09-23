VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCodeZip 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   25
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnReadFromFile 
      Caption         =   "Load Zip-file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Picture         =   "frmCodeZip.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton btnPaste 
      Caption         =   "Paste Zip-file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Picture         =   "frmCodeZip.frx":06C2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CodeText"
      DataSource      =   "rsCodeZip"
      Height          =   285
      Index           =   0
      Left            =   5400
      MaxLength       =   70
      TabIndex        =   15
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "DateInDatabase"
      DataSource      =   "rsCodeZip"
      Height          =   285
      Index           =   1
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   70
      TabIndex        =   14
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Author"
      DataSource      =   "rsCodeZip"
      Height          =   285
      Index           =   2
      Left            =   5400
      MaxLength       =   70
      TabIndex        =   13
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "AuthorMail"
      DataSource      =   "rsCodeZip"
      Height          =   285
      Index           =   3
      Left            =   5400
      MaxLength       =   70
      TabIndex        =   12
      Top             =   1320
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "AuthorUrl"
      DataSource      =   "rsCodeZip"
      Height          =   285
      Index           =   4
      Left            =   5400
      MaxLength       =   70
      TabIndex        =   11
      Top             =   1680
      Width           =   5295
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
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5100
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.CommandButton btnStatistic 
      Caption         =   "Code Statistic"
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
      Left            =   1920
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin VB.ComboBox cboCodeType 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
   End
   Begin VB.ComboBox cboCodeLanguage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   360
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Data rsCodeType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\CodeSnippets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeType"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   3960
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data rsCodeZip 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeZip.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeZip"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox Text2 
      DataField       =   "CodeDescription"
      DataSource      =   "rsCodeZip"
      Height          =   3015
      Left            =   3840
      TabIndex        =   9
      Top             =   2040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   5318
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmCodeZip.frx":0D84
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Zip File:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   3720
      TabIndex        =   24
      Top             =   5760
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   19
      X1              =   3600
      X2              =   3600
      Y1              =   5880
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   18
      X1              =   10920
      X2              =   10920
      Y1              =   5880
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   17
      X1              =   4560
      X2              =   10920
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   16
      X1              =   3600
      X2              =   10920
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.OLE OLE1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Class           =   "Package"
      DataField       =   "CodeZip"
      DataSource      =   "rsCodeZip"
      Height          =   1575
      Left            =   3720
      OleObjectBlob   =   "frmCodeZip.frx":0E06
      TabIndex        =   23
      Top             =   6120
      Width           =   5415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   15
      X1              =   3600
      X2              =   10920
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   14
      X1              =   3600
      X2              =   10920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   13
      X1              =   10920
      X2              =   10920
      Y1              =   240
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   12
      X1              =   3600
      X2              =   3600
      Y1              =   240
      Y2              =   5640
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date stored in database:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   20
      Top             =   5160
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code Text:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   19
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   18
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author Mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   17
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Author Internet:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Text:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   780
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   11
      X1              =   3480
      X2              =   3480
      Y1              =   2040
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   10
      X1              =   240
      X2              =   240
      Y1              =   2040
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   9
      X1              =   240
      X2              =   3480
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   1440
      X2              =   3480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   825
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   3480
      X2              =   3480
      Y1              =   1320
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   240
      X2              =   3480
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   1440
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Language:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   3480
      X2              =   3480
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   3480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1800
      X2              =   3480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "frmCodeZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodeBook() As Variant, bNewRecord As Boolean, iList1Index As Integer
Dim boolWrite As Boolean
Dim rsCodeLanguage As Recordset
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 8
        Label2(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 19
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Sub LoadCodeLanguage()
    On Error Resume Next
    cboCodeLanguage.Clear
    With rsCodeLanguage
        .MoveFirst
        Do While Not .EOF
            cboCodeLanguage.AddItem .Fields("Language")
        .MoveNext
        Loop
    End With
    If Not IsNull(rsUser.Fields("PrefferedLanguage")) Then
        cboCodeLanguage.Text = rsUser.Fields("PrefferedLanguage")
    Else
        cboCodeLanguage.ListIndex = 0
    End If
End Sub
Private Sub SelectCodeType()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeType WHERE Trim(CodeLanguage) ="
    Sql = Sql & Chr(34) & Trim(cboCodeLanguage.Text) & Chr(34)
    With rsCodeType
        .RecordSource = Sql
        .Refresh
        cboCodeType.Clear
        .Recordset.MoveFirst
        Do While Not .Recordset.EOF
            cboCodeType.AddItem .Recordset.Fields("CodeType")
        .Recordset.MoveNext
        Loop
    End With
End Sub

Public Sub ShowAuthor()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeZip WHERE CLng(CodeAuto) ="
    Sql = Sql & Chr(34) & CLng(m_lSnippet) & Chr(34)
    rsCodeZip.RecordSource = Sql
    rsCodeZip.Refresh
End Sub
Public Sub CopyZipToClip()
    On Error Resume Next
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
    Clipboard.Clear
    Clipboard.SetText Text2.Text
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
                For i = 0 To 8
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label2(i).Caption
                    Else
                        Label2(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("btnReadFromFile")) Then
                    .Fields("btnReadFromFile") = btnReadFromFile.Caption
                Else
                    btnReadFromFile.Caption = .Fields("btnReadFromFile")
                End If
                If IsNull(.Fields("btnPaste")) Then
                    .Fields("btnPaste") = btnPaste.Caption
                Else
                    btnPaste.Caption = .Fields("btnPaste")
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
        For i = 0 To 8
            .Fields(i + 1) = Label2(i).Caption
        Next
        .Fields("btnReadFromFile") = btnReadFromFile.Caption
        .Fields("btnPaste") = btnPaste.Caption
        .Fields("btnStatistic") = btnStatistic.Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Public Function SelectRecords() As Boolean
Dim Sql As String
    On Error GoTo errSelectRecords
    List1.Clear
    Sql = "SELECT * FROM CodeZip WHERE Trim(CodeType) ="
    Sql = Sql & Chr(34) & Trim(cboCodeType.Text) & Chr(34)
    Sql = Sql & " AND Trim(CodeLanguage) ="
    Sql = Sql & Chr(34) & Trim(cboCodeLanguage.Text) & Chr(34)
    Sql = Sql & " ORDER BY CodeType"
    
    With rsCodeZip
        .RecordSource = Sql
        .Refresh
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Label1(1).Caption = "Records: " & .Recordset.RecordCount
        Label1(1).ForeColor = rsUser.Fields("LabelColor")
        
    End With
    SelectRecords = True
    Exit Function
    
errSelectRecords:
    Label1(1).Caption = "Records: " & 0
    Label1(1).ForeColor = rsUser.Fields("LabelColor")
    SelectRecords = False
    Err.Clear
End Function
Public Sub SelectAllCode()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeZip"
    rsCodeZip.RecordSource = Sql
    rsCodeZip.Refresh
End Sub
Public Sub NewRecord()
    On Error Resume Next
    rsCodeZip.Recordset.AddNew
    bNewRecord = True
    Text1(0).SetFocus
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
    rsCodeZip.Recordset.Delete
    List1.RemoveItem (iList1Index)
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    boolWrite = True
    With rsCodeZip.Recordset
        .MoveLast
        .MoveFirst
        ReDim vCodeBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("CodeText")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vCodeBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    boolWrite = False
End Sub

Private Sub btnPaste_Click()
    On Error GoTo errPaste
    OLE1.Paste
    Exit Sub
    
errPaste:
    Beep
    MsgBox Err.Description, vbExclamation, "Paste"
    Err.Clear
End Sub

Private Sub btnReadFromFile_Click()
    On Error GoTo errNewRecord
    With Cmd1
        .FileName = ""
        .DialogTitle = "Load Zip-file from disk"
        .Filter = "Dokument (*.zip)|*.zip"
        .FilterIndex = 1
        .Action = 1
    End With
    OLE1.CreateEmbed Cmd1.FileName
    Exit Sub
    
errNewRecord:
    Beep
    MsgBox Err.Description, vbExclamation, "Read From File"
    Err.Clear
End Sub

Private Sub btnSearch_Click()
    With frmShowAuthors
        .Text1.Text = Text1(2).Text
        .Show vbModal
    End With
End Sub
Private Sub btnStatistic_Click()
    m_boolSnippet = False
    frmCodeStatistic.Show 1
End Sub

Private Sub cboCodeLanguage_Click()
    On Error Resume Next
    SelectCodeType
    cboCodeType.ListIndex = 0
End Sub
Private Sub cboCodeType_Click()
    On Error Resume Next
    If SelectRecords Then
        LoadList1
        List1.ListIndex = 0
    End If
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsCodeType.Refresh
    With rsCodeZip
        .Refresh
        With .Recordset
            .MoveLast
            .MoveFirst
            Label1(0).Caption = "Records: " & .RecordCount
            Label1(0).ForeColor = rsUser.Fields("LabelColor")
        End With
    End With
    LoadCodeLanguage
    DoEvents
    SelectCodeType
    DoEvents
    DoEvents
    cboCodeType.ListIndex = 0
    ReadText
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = True
        .Buttons(6).Enabled = True
        .Buttons(15).Enabled = True
        .Buttons(16).Enabled = True
    End With
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsCodeZip.DatabaseName = m_strCodeZip
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCodeZip")
    rsCodeType.DatabaseName = m_strCodeSnippet
    Set rsCodeLanguage = m_dbCodeSnippet.OpenRecordset("Language")
    m_iFormNo = 33
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
    rsCodeZip.Recordset.Close
    rsCodeType.Recordset.Close
    rsCodeLanguage.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    DisableButtons 1
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(15).Enabled = False
        .Buttons(16).Enabled = False
    End With
    Set frmCodeZip = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    If boolWrite Then Exit Sub
    iList1Index = List1.ListIndex
    rsCodeZip.Recordset.Bookmark = vCodeBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        If bNewRecord Then
            With rsCodeZip.Recordset
                .Fields("CodeLanguage") = cboCodeLanguage.Text
                .Fields("CodeType") = cboCodeType.Text
                .Fields("CodeText") = Text1(0).Text
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
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        Text2.SelText = vbTab
        KeyCode = 0
    End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error Resume Next
   If Button = vbRightButton Then
      frmMDI.PopupMenu frmMDI.mnuFormat
   End If
End Sub
