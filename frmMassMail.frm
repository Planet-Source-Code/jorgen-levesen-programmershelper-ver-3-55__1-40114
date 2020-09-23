VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMassMail 
   BackColor       =   &H00404040&
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9705
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   6660
      Left            =   7920
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Data rsMassMail 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programing\Source Code\CodeMaster.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "MassMail"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "MailTo"
      DataSource      =   "rsMassMail"
      Height          =   615
      Index           =   0
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   360
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Mailregarding"
      DataSource      =   "rsMassMail"
      Height          =   615
      Index           =   1
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   5655
   End
   Begin VB.CommandButton btnReSend 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      Picture         =   "frmMassMail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Send Mail"
      Top             =   6480
      Width           =   7335
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      DataField       =   "MailContents"
      DataSource      =   "rsMassMail"
      Height          =   4455
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7858
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMassMail.frx":27A2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send Mail"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   7800
      X2              =   9480
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   9000
      X2              =   9480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   9480
      X2              =   9480
      Y1              =   240
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   7800
      X2              =   7800
      Y1              =   240
      Y2              =   7080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   240
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   7560
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   960
      X2              =   7560
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "frmMassMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarksMail() As Variant
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 3
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 7
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
                If IsNull(.Fields("label1(2)")) Then
                    .Fields("label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("label1(2)")
                End If
                If IsNull(.Fields("Label1(3)")) Then
                    .Fields("Frame2") = Label1(3).Caption
                Else
                    Label1(3).Caption = .Fields("Label1(3)")
                End If
                If IsNull(.Fields("btnReSend")) Then
                    .Fields("btnReSend") = btnReSend.ToolTipText
                Else
                    btnReSend.ToolTipText = .Fields("btnReSend")
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
        .Fields("Label1(2)") = Label1(2).Caption
        .Fields("Label1(3)") = Label1(3).Caption
        .Fields("btnReSend") = btnReSend.ToolTipText
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsMassMail.Recordset
        .MoveLast
        ReDim bookmarksMail(.RecordCount)
        .MoveFirst
        Do While Not .EOF
            List1.AddItem Format(.Fields("MailDate"), "dd.mm.yyyy") & " - " & Format(.Fields("MailTime"), "hh:mm")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarksMail(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub btnReSend_Click()
        On Error Resume Next
        Call SendOutlookMail(Text1(1).Text, Text1(0).Text, RichTextBox1.Text)
        With rsMassMail.Recordset
            .AddNew
            .Fields("MailDate") = Format(Now, "dd.mm.yyyy")
            .Fields("MailTime") = Format(Now, "hh:mm")
            .Fields("MailTo") = Trim(Text1(0).Text)
            .Fields("Mailregarding") = Trim(Text1(1).Text)
            .Fields("MailContents") = RichTextBox1.TextRTF
            .Update
        End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsMassMail.Refresh
    LoadList1
    List1.ListIndex = 0
    ReadText
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsMassMail.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmMassMail")
    m_iFormNo = 12
    DisableButtons 1
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Form Load"
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
    rsMassMail.Recordset.Close
    rsUser.Close
    m_iFormNo = 0
    DisableButtons 2
    rsLanguage.Close
    Set frmMassMail = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsMassMail.Recordset.Bookmark = bookmarksMail(List1.ItemData(List1.ListIndex))
End Sub
