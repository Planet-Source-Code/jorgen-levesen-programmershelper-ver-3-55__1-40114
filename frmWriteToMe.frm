VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWriteToMe 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mail To System Responsible"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnSend 
      Height          =   495
      Left            =   3360
      Picture         =   "frmWriteToMe.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Send this e-mail"
      Top             =   5640
      Width           =   4935
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   240
      Picture         =   "frmWriteToMe.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   5640
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393217
      BackColor       =   16777152
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmWriteToMe.frx":28EC
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   8280
      X2              =   8280
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   240
      X2              =   8280
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   8280
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "I would like to have the folowing errors corrected / new facilities added:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   7695
   End
End
Attribute VB_Name = "frmWriteToMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSupplier As Recordset
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    Label1.ForeColor = rsUser.Fields("LabelColor")
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub
Private Sub ReadText()
    On Error Resume Next    'this is only text
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.ToolTipText
                Else
                    btnExit.ToolTipText = .Fields("btnExit")
                End If
                If IsNull(.Fields("btnSend")) Then
                    .Fields("btnSend") = btnSend.ToolTipText
                Else
                    btnSend.ToolTipText = .Fields("btnSend")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("label1") = Label1.Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Fields("btnSend") = btnSend.ToolTipText
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSend_Click()
Dim CarbonCopy As String
    On Error Resume Next
    If Len(RichTextBox1.Text) <> 0 Then
        CarbonCopy = ""
        Call SendOutlookMail("Programming: Suggestions, Errors", rsSupplier.Fields("SupplierEmailySysResponse"), RichTextBox1.Text)
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
    LoadBackground
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Set rsSupplier = m_dbPrograming.OpenRecordset("Supplier")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmWriteToMe")
    m_iFormNo = 23
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    Set frmWriteToMe = Nothing
End Sub
Private Sub RichTextBox1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim mblnTabPressed As Boolean
    mblnTabPressed = (KeyCode = vbKeyTab)
    If mblnTabPressed Then
        RichTextBox1.SelText = vbTab
        KeyCode = 0
    End If
End Sub
