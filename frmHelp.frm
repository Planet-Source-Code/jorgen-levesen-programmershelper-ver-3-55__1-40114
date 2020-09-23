VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00000000&
   Caption         =   "Help"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   ForeColor       =   &H00000000&
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   5640
      Picture         =   "frmHelp.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Data rsLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\MasterEmpW\MasterLang.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "frmCountry"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage"
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmHelp.frx":0454
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   4455
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFormLanguage As Recordset
Private Sub ReadText()
    On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsFormLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.ToolTipText
                Else
                    btnExit.ToolTipText = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Update
    End With
End Sub

Private Sub btnEditText_Click()
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
   rsLanguage.RecordSource = Trim(CStr(Label1.Caption))
   rsLanguage.Refresh
    With rsLanguage.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then Exit Do
        .MoveNext
        Loop
    End With
    ReadText
End Sub
Private Sub Form_Load()
    On Error Resume Next
    rsLanguage.DatabaseName = m_strProgramLng
    Set rsFormLanguage = m_dbLanguage.OpenRecordset("frmHelp")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Recordset.Close
    rsFormLanguage.Close
    Set frmHelp = Nothing
End Sub

Private Sub Label1_Click()
    frmEditHelp.Caption = Label1.Caption
    frmEditHelp.Show
End Sub
