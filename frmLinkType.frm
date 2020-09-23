VERSION 5.00
Begin VB.Form frmLinkType 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Link Type"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3480
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   2955
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkType"
      DataSource      =   "rsLinkType"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "LinkDescription"
      DataSource      =   "rsLinkType"
      Height          =   1965
      Index           =   1
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Data rsLinkType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LinkType"
      Top             =   600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   5880
      X2              =   5880
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   5880
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   5880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Link Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Link Type Note:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "frmLinkType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLinkBook() As Variant
Dim bNewRecord As Boolean
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 2
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Public Sub DeleteLinkType()
    On Error GoTo errDelete
    rsLinkType.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
    Exit Sub
    
errDelete:
    Beep
    MsgBox Err.Description, vbCritical, "Delete a link Type"
    Err.Clear
End Sub

Public Sub NewLinkType()
    bNewRecord = True
    rsLinkType.Recordset.AddNew
    Text1(0).SetFocus
End Sub
Private Sub ReadText()
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
                For i = 0 To 2
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
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then
                Exit Do
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        For i = 0 To 2
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Update
    End With
End Sub

Private Sub LoadList1()
    List1.Clear
    With rsLinkType.Recordset
        .MoveLast
        .MoveFirst
        ReDim vLinkBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("LinkType")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vLinkBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsLinkType.Refresh
    ReadText
    LoadList1
    List1.ListIndex = 0
    LoadBackground
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = True
        .Buttons(6).Enabled = True
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsLinkType.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmLinkType")
    m_iFormNo = 35
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLinkType.Recordset.Close
    rsLanguage.Close
    rsUser.Close
    m_iFormNo = 0
    Erase vLinkBook
    With frmMDI.Toolbar1
        .Buttons(4).Enabled = False
        .Buttons(6).Enabled = False
    End With
    Set frmLinkType = Nothing
End Sub


Private Sub List1_Click()
    On Error Resume Next
    rsLinkType.Recordset.Bookmark = vLinkBook(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    On Error GoTo errNewLink
    Select Case Index
    Case 0
        If bNewRecord Then
            With rsLinkType.Recordset
                .Fields("LinkType") = Trim(Text1(0).Text)
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
    MsgBox Err.Description, vbCritical, "New Internet Link Type"
    Err.Clear
End Sub
