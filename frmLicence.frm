VERSION 5.00
Begin VB.Form frmLicence 
   BackColor       =   &H00404040&
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   10215
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3000
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cmbCustomer 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerName"
      DataSource      =   "rsLicence"
      Height          =   315
      Left            =   4920
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "ProjectID"
      DataSource      =   "rsLicence"
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
      Index           =   0
      Left            =   4920
      TabIndex        =   12
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "NoOfProgrammes"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   11
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "PriceAProgram"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "PriceAProgramDim"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LiberationKey"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   8
      Top             =   3480
      Width           =   2775
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4905
      Left            =   7920
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Data rsLicence 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Licence"
      Top             =   5400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "ProgrammeVersion"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   6
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "PurchaseDate"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   6
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LastUpdate"
      DataSource      =   "rsLicence"
      Height          =   285
      Index           =   7
      Left            =   4920
      MaxLength       =   10
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5685
      Left            =   9240
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnMassMail 
      Caption         =   "&Mass Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      Picture         =   "frmLicence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   5880
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   3000
      TabIndex        =   24
      Top             =   240
      Width           =   45
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   2880
      X2              =   10080
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   2880
      X2              =   10080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   10080
      X2              =   10080
      Y1              =   360
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2880
      X2              =   2880
      Y1              =   360
      Y2              =   6240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Customer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   22
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Programme:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   21
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Programmes:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   20
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Price a Programme:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   19
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Date of purchase:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   18
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Liberation key:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   17
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Users"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   16
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Programme version no:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last update:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   14
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Projects / Programme"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vProjectBook() As Variant
Dim vLicenceBook() As Variant
Dim vCustomBook() As Variant, vCustomerBook2() As Variant
Dim bNewRecord As Boolean
Dim rsCustomer As Recordset
Dim rsUser As Recordset
Dim rsProjects As Recordset
Dim rsLanguage As Recordset
Dim WClone As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 9
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
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
                    If IsNull(i + 1) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("btnMassMail")) Then
                    .Fields("btnMassMail") = btnMassMail.Caption
                Else
                    btnMassMail.Caption = .Fields("btnMassMail")
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
        For i = 0 To 9
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("btnMassMail") = btnMassMail.Caption
        .Fields("Help") = sHelp
        .Update
    End With
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


Private Sub LoadList1()
    List1.Clear
    With rsProjects
        .MoveLast
        .MoveFirst
        ReDim vProjectBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("ProjectID")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vProjectBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub LoadList2()
    'list all programmes sold this project
    List2.Clear
    List3.Clear
    On Error Resume Next
    Set WClone = rsLicence.Recordset.Clone()
    With WClone
        .MoveLast
        .MoveFirst
        ReDim vLicenceBook(.RecordCount)
        ReDim vCustomerBook2(.RecordCount)
        Do While Not .EOF
            'find the customer name
            rsCustomer.MoveFirst
            Do While Not rsCustomer.EOF
                If CLng(rsCustomer.Fields("AutoLine")) = CLng(.Fields("AutoLineCustomer")) Then
                    List2.AddItem .Fields("CustomerName")
                    List2.ItemData(List2.NewIndex) = List2.ListCount - 1
                    vLicenceBook(List2.ListCount - 1) = .Bookmark
                    If Not IsNull(rsCustomer.Fields("CustomerEMail")) Then
                        If Not IsNull(rsCustomer.Fields("CustomerEMail")) Then
                            List3.AddItem rsCustomer.Fields("CustomerEMail")
                            List3.ItemData(List3.NewIndex) = List3.ListCount - 1
                            vCustomerBook2(List3.ListCount - 1) = .Bookmark
                        Else
                            List3.AddItem "?"
                            List3.ItemData(List3.NewIndex) = List3.ListCount - 1
                            vCustomerBook2(List3.ListCount - 1) = .Bookmark
                        End If
                    End If
                End If
            rsCustomer.MoveNext
            Loop
        .MoveNext
        Loop
    End With
    Set WClone = Nothing
End Sub

Private Sub SelectCustomerRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM Licence WHERE ProjectID ="
    Sql = Sql & Chr(34) & rsProjects.Fields("ProjectID") & Chr(34)
    rsLicence.RecordSource = Sql
    rsLicence.Refresh
End Sub

Public Sub DeleteRecord()
    On Error Resume Next
    rsLicence.Recordset.Delete
    LoadList2
End Sub

Private Sub btnMassMail_Click()
    On Error Resume Next
    If List3.ListCount <> 0 Then
        With frmEmail
            For n = 0 To List3.ListCount - 1
                List3.ListIndex = n
                If List3.List(List3.ListIndex) <> "?" Then
                    .List1.AddItem List3.List(List3.ListIndex)
                End If
            Next
            .Text2.Visible = False
            .cboAdr.Visible = False
            .btnMailTo.Visible = False
            .Show vbModal
        End With
    End If
End Sub

Public Sub NewRecord()
    bNewRecord = True
    rsLicence.Recordset.AddNew
    cmbCustomer.SetFocus
End Sub


Private Sub cmbCustomer_LostFocus()
    On Error GoTo errCustomerNew
    If bNewRecord Then
        With rsLicence.Recordset
            rsCustomer.Bookmark = vCustomBook(cmbCustomer.ItemData(cmbCustomer.ListIndex))
            .Fields("ProjectID") = rsProjects.Fields("ProjectID")
            .Fields("AutoLineCustomer") = CLng(rsCustomer.Fields("AutoLine"))
            .Fields("CustomerName") = CStr(rsCustomer.Fields("CustomerName"))
            .Fields("PurchaseDate") = CDate(Format(Now, "dd.mm.yyyy"))
            .Update
            LoadList2
            .Bookmark = .LastModified
        End With
        bNewRecord = False
    End If
    Exit Sub
    
errCustomerNew:
    Beep
    MsgBox Err.Description, vbCritical, "New Sale"
    Resume Next
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    rsLicence.Refresh
    LoadcmbCustomer
    LoadList1
    List1.ListIndex = 0
    LoadList2
    List2.ListIndex = 0
    ReadText
    DisableButtons 2
    frmMDI.Toolbar1.Buttons(8).Enabled = False
    frmMDI.Toolbar1.Buttons(9).Enabled = False
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsLicence.DatabaseName = m_strPrograming
    Set rsCustomer = m_dbPrograming.OpenRecordset("Customer")
    Set rsProjects = m_dbPrograming.OpenRecordset("Projects")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmLicence")
    m_iFormNo = 10
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
    LoadBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLicence.Recordset.Close
    rsCustomer.Close
    rsProjects.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    Erase vProjectBook
    Erase vLicenceBook
    Erase vCustomBook
    Erase vCustomerBook2
    DisableButtons 1
    Set frmLicence = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsProjects.Bookmark = vProjectBook(List1.ItemData(List1.ListIndex))
    Label1(10).Caption = rsProjects.Fields("ProjectText")
    Label1(10).ForeColor = rsUser.Fields("LabelColor")
    SelectCustomerRecords
    LoadList2
    List2.ListIndex = 0
End Sub
Private Sub List2_Click()
    On Error Resume Next
    rsLicence.Recordset.Bookmark = vLicenceBook(List2.ItemData(List2.ListIndex))
End Sub


