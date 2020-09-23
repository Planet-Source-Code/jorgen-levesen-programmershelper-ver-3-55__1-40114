VERSION 5.00
Begin VB.Form frmPasswords 
   BackColor       =   &H00404040&
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   9450
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6840
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox RichTextBox1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Password"
      DataSource      =   "rsPasswords"
      Height          =   5295
      Left            =   3600
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SiteName"
      DataSource      =   "rsPasswords"
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   6360
      Width           =   5655
   End
   Begin VB.Data rsPasswords 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Passwords"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   6270
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   9240
      X2              =   9240
      Y1              =   120
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   9240
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   9240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Site Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Site"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBook() As Variant, bNewRecord As Boolean
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 1
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub
Private Sub ReadText()
Dim sHelp As String
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("label1(1)")
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
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Public Sub WritePasswords()
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    DoEvents
    cPrint.pStartDoc
    With rsPasswords.Recordset
        .MoveFirst
        Do While Not .EOF
            If cPrint.pEndOfPage Then
                cPrint.pFooter
                cPrint.pNewPage
            End If
            cPrint.FontSize = 12
            cPrint.FontBold = True
            cPrint.pPrint .Fields("SiteName"), 0.3
            cPrint.FontBold = False
            cPrint.pMultiline Format(.Fields("Password")), 0.3, , , False, True
            cPrint.pPrint
            cPrint.pLine
            cPrint.pPrint
        .MoveNext
        Loop
    End With
    
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsPasswords.Recordset
        .MoveLast
        .MoveFirst
        ReDim vBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("SiteName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub DeleteRecord()
    On Error Resume Next
    rsPasswords.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
End Sub

Public Sub NewRecord()
    bNewRecord = True
    rsPasswords.Recordset.AddNew
    Text1.SetFocus
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsPasswords.Refresh
    LoadList1
    ReadText
    DisableButtons 2
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsPasswords.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmPasswords")
    m_iFormNo = 14
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
    rsPasswords.Recordset.Close
    rsUser.Close
    rsLanguage.Close
    DisableButtons 1
    Set frmPasswords = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsPasswords.Recordset.Bookmark = vBook(List1.ItemData(List1.ListIndex))
End Sub


Private Sub Text1_LostFocus()
    On Error GoTo errText1
    If bNewRecord Then
        With rsPasswords.Recordset
            .Fields("SiteName") = Trim(Text1.Text)
            .Update
            LoadList1
            .Bookmark = .LastModified
        End With
        bNewRecord = False
    End If
    Exit Sub
    
errText1:
    Beep
    MsgBox Err.Description, vbCritical, "New Site"
    Resume Next
End Sub
