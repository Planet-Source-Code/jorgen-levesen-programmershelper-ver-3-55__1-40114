VERSION 5.00
Begin VB.Form frmLinks 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Links"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   585
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cboLinkTypes 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   840
      Width           =   3975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5295
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkName"
      DataSource      =   "rsVBLinks"
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   6
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkHyper"
      DataSource      =   "rsVBLinks"
      Height          =   285
      Index           =   1
      Left            =   5040
      MaxLength       =   70
      TabIndex        =   5
      Top             =   1440
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkNote"
      DataSource      =   "rsVBLinks"
      Height          =   3765
      Index           =   2
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3240
      Width           =   6015
   End
   Begin VB.Data rsVBLinks 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VBLinks"
      Top             =   480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkUserName"
      DataSource      =   "rsVBLinks"
      Height          =   285
      Index           =   3
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkPassWord"
      DataSource      =   "rsVBLinks"
      Height          =   285
      Index           =   4
      Left            =   5040
      MaxLength       =   50
      TabIndex        =   2
      Top             =   2640
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkInDatabase"
      DataSource      =   "rsVBLinks"
      Enabled         =   0   'False
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
      Index           =   5
      Left            =   9360
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkLastUsed"
      DataSource      =   "rsVBLinks"
      Enabled         =   0   'False
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
      Index           =   6
      Left            =   9360
      MaxLength       =   50
      TabIndex        =   0
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   600
      TabIndex        =   17
      Top             =   480
      Width           =   750
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   240
      X2              =   11280
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   240
      X2              =   11280
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   11280
      X2              =   11280
      Y1              =   360
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   240
      X2              =   240
      Y1              =   360
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   4800
      X2              =   4800
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   480
      X2              =   480
      Y1              =   600
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   480
      X2              =   4800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1560
      X2              =   4800
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Link Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   14
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link URL:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   13
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Note:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   12
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5040
      TabIndex        =   11
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link User Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Stored In Database:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   9360
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Date Used::"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   9360
      TabIndex        =   8
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "frmLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLinkBook() As Variant
Dim bNewRecord As Boolean
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Dim rsLinkType As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 8
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 4
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub


Public Sub SelectRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM VBLinks WHERE Trim(LinkType) ="
    Sql = Sql & Chr(34) & Trim(cboLinkTypes.Text) & Chr(34)
    
    rsVBLinks.RecordSource = Sql
    rsVBLinks.Refresh
End Sub

Public Sub DeleteLink()
    On Error GoTo errDelete
    rsVBLinks.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
    Exit Sub
    
errDelete:
    Beep
    MsgBox Err.Description, vbCritical, "Delete a link"
    Err.Clear
End Sub


Private Sub LoadcboLinkTypes()
    On Error Resume Next
    cboLinkTypes.Clear
    With rsLinkType
        .MoveFirst
        Do While Not .EOF
            cboLinkTypes.AddItem .Fields("LinkType")
        .MoveNext
        Loop
    End With
End Sub

Public Sub NewLink()
    bNewRecord = True
    rsVBLinks.Recordset.AddNew
    Text1(0).SetFocus
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
                If IsNull("Form") Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                For i = 0 To 8
                    If IsNull(i + 2) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
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
        .Fields("Form") = Me.Caption
        For i = 0 To 8
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsVBLinks.Recordset
        .MoveLast
        .MoveFirst
        ReDim vLinkBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("LinkName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vLinkBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub UpdateLinkDate()
    On Error Resume Next
    With rsVBLinks.Recordset
        .Edit
        .Fields("LinkLastUsed") = Format(Now, "dd.mm.yyyy")
        .Update
        .Bookmark = .LastModified
    End With
End Sub

Private Sub cboLinkTypes_Click()
    On Error Resume Next
    SelectRecords
    LoadList1
    List1.ListIndex = 0
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsVBLinks.Refresh
    ReadText
    LoadcboLinkTypes
    cboLinkTypes.ListIndex = 0
    LoadList1
    List1.ListIndex = 0
    LoadBackground
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = True
        .Buttons(6).Enabled = True
        .Buttons(16).Enabled = True
    End With
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsVBLinks.DatabaseName = m_strPrograming
    Set rsLinkType = m_dbPrograming.OpenRecordset("LinkType")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmLinks")
    m_iFormNo = 34
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsVBLinks.Recordset.Close
    rsLinkType.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    Erase vLinkBook
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(6).Enabled = False
        .Buttons(16).Enabled = False
    End With
    Set frmLinks = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsVBLinks.Recordset.Bookmark = vLinkBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo errNewLink
    Select Case Index
    Case 0
        If bNewRecord Then
            With rsVBLinks.Recordset
                .Fields("LinkType") = Trim(cboLinkTypes.Text)
                .Fields("LinkName") = Trim(Text1(0).Text)
                .Fields("LinkInDatabase") = Format(Now, "dd.mm.yyyy")
                .Update
                LoadList1
                .Bookmark = .LastModified
                bNewRecord = False
            End With
        End If
    Case Else
    End Select
    Exit Sub
    
errNewLink:
    Beep
    MsgBox Err.Description, vbCritical, "New Internet Link"
    Err.Clear
End Sub


