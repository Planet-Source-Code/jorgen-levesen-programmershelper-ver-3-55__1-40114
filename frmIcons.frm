VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIcons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Icons To database"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      DataField       =   "IconPicture"
      DataSource      =   "rsIcons"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1920
      ScaleHeight     =   1065
      ScaleWidth      =   1305
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog Cmd1 
      Left            =   360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   2760
      Left            =   6960
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.Data rsIcons 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programing\Source Code\ProgramIco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Icons"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdReadFromFile 
      Height          =   375
      Left            =   3600
      Picture         =   "frmIcons.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Read Icon from file"
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   375
      Left            =   3600
      Picture         =   "frmIcons.frx":06C2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Copy this Icon to the Clipboard"
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdPaste 
      Height          =   375
      Left            =   3600
      Picture         =   "frmIcons.frx":0D84
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Paste a Icon"
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox cmbIconType 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "IconDescription"
      DataSource      =   "rsIcons"
      Height          =   285
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Icon"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   3480
      X2              =   3480
      Y1              =   1200
      Y2              =   2760
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1680
      X2              =   1680
      Y1              =   1200
      Y2              =   2760
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1680
      X2              =   3480
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1680
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarkIcon() As Variant
Dim boolNew As Boolean
Dim rsUser As Recordset
Dim rsIconType As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture2.Visible = False
    Picture2.AutoRedraw = True
    Picture2.AutoSize = True
    Picture2.BorderStyle = 0
    Picture2.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture2
    For i = 0 To 2
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Sub ReadText()
Dim sHelp As String
    'On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("Label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("Label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("Label1(2)")) Then
                    .Fields("Label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("Label1(2)")
                End If
                If IsNull(.Fields("cmdPaste")) Then
                    .Fields("cmdPaste") = cmdPaste.ToolTipText
                Else
                    cmdPaste.ToolTipText = .Fields("cmdPaste")
                End If
                If IsNull(.Fields("cmdCopy")) Then
                    .Fields("cmdCopy") = cmdCopy.ToolTipText
                Else
                    cmdCopy.ToolTipText = .Fields("cmdCopy")
                End If
                If IsNull(.Fields("cmdReadFromFile")) Then
                    .Fields("cmdReadFromFile") = cmdReadFromFile.ToolTipText
                Else
                    cmdReadFromFile.ToolTipText = .Fields("cmdReadFromFile")
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
        .Fields("Form") = Me.Caption
        .Fields("Label1(0)") = Label1(0).Caption
        .Fields("Label1(1)") = Label1(1).Caption
        .Fields("Label1(2)") = Label1(2).Caption
        .Fields("cmdPaste") = cmdPaste.ToolTipText
        .Fields("cmdCopy") = cmdCopy.ToolTipText
        .Fields("cmdReadFromFile") = cmdReadFromFile.ToolTipText
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Public Sub DeleteIcon()
    On Error Resume Next
    rsIcons.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
End Sub
Private Sub LoadcmbIconType()
    On Error Resume Next
    With rsIconType
        .MoveFirst
        Do While Not .EOF
            cmbIconType.AddItem .Fields("IconType")
        .MoveNext
        Loop
    End With
End Sub


Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsIcons.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarkIcon(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("IconDescription")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarkIcon(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub NewIcon()
    On Error Resume Next
    If boolNew Then Exit Sub
    rsIcons.Recordset.AddNew
    boolNew = True
    Text1.SetFocus
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsIcons.Refresh
    LoadcmbIconType
    cmbIconType.ListIndex = 0
    ReadText
    LoadBackground
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = True
        .Buttons(6).Enabled = True
    End With
End Sub
Private Sub SelectIconType()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM Icons WHERE Trim(IconType) ="
    Sql = Sql & Chr(34) & Trim(cmbIconType.Text) & Chr(34)
    rsIcons.RecordSource = Sql
    rsIcons.Refresh
End Sub

Private Sub cmbIconType_Click()
    On Error Resume Next
    If boolNew Then Exit Sub
    List1.Clear
    SelectIconType
    LoadList1
    List1.ListIndex = 0
End Sub
Private Sub cmdCopy_Click()
    On Error GoTo errCopy
    Clipboard.Clear
    Clipboard.SetData Image1.Image, vbCFBitmap
    Exit Sub
    
errCopy:
    Beep
    MsgBox Err.Description, vbCritical, "Copy Picture to Clipboard"
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    On Error GoTo errPicPaste
    Image1.Picture = Clipboard.GetData(vbCFDIB)
    Exit Sub
    
errPicPaste:
    Beep
    MsgBox Err.Description, vbExclamation, "Paste Picture"
    Err.Clear
End Sub
Private Sub cmdReadFromFile_Click()
    On Error GoTo errPicRead
    With Cmd1
        .FileName = ""
        .DialogTitle = "Load Picture from disk"
        .Filter = "(*.ico)|*.ico"
        .FilterIndex = 1
        .ShowOpen
        'Image1.Image = System.Drawing.Image.FromFile(.FileName)
        Image1.Picture = LoadPicture(.FileName)
    End With
    Exit Sub
    
errPicRead:
    Beep
    MsgBox Err.Description, vbExclamation, "Read Picture from disk"
    Err.Clear
End Sub

Private Sub Form_Load()
Dim sName As String, dbTemp As Database
    On Error GoTo errForm_Load
    Me.Move 0, 0
    sName = App.Path & "\CodeIco.mdb"
    Set dbTemp = OpenDatabase(sName)
    rsIcons.DatabaseName = sName
    Set rsIconType = dbTemp.OpenRecordset("IconType")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmIcons")
    m_iFormNo = 36
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsIcons.Recordset.Close
    rsUser.Close
    rsIconType.Close
    rsLanguage.Close
    Erase bookmarkIcon
    m_iFormNo = 0
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(6).Enabled = False
    End With
    Set frmIcons = Nothing
End Sub
Private Sub List1_Click()
    On Error Resume Next
    rsIcons.Recordset.Bookmark = bookmarkIcon(List1.ItemData(List1.ListIndex))
End Sub


Private Sub Text1_LostFocus()
    On Error GoTo errNew
    If boolNew Then
        With rsIcons.Recordset
            .Fields("IconType") = cmbIconType.Text
            .Fields("IconDescription") = Trim(Text1.Text)
            .Update
            LoadList1
            .Bookmark = .LastModified
            boolNew = False
        End With
    End If
    Exit Sub
    
errNew:
    Beep
    MsgBox Err.Description, vbCritical, "New Icon"
    Err.Clear
End Sub
