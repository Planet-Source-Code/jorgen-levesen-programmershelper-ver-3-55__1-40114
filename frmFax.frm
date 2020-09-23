VERSION 5.00
Begin VB.Form frmFax 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Fax"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmFax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   240
      Picture         =   "frmFax.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Exit"
      Top             =   5400
      Width           =   3375
   End
   Begin VB.CommandButton btnSend 
      Height          =   495
      Left            =   3720
      Picture         =   "frmFax.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Send fax"
      Top             =   5400
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4590
      Left            =   3720
      TabIndex        =   3
      Top             =   720
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4905
      Left            =   240
      ReadOnly        =   0   'False
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Stored Fax files"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDir As String
Dim rsMyRec As Recordset
Dim rsLanguage As Recordset

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSend_Click()
Dim strFax As String, strNumber As String
    On Error Resume Next
    Select Case m_iFormNo
    Case 5  'customers
        If Len(frmCustomer.Text1(8).Text) = 0 Then Exit Sub
        strFax = File1.Path & "\" & File1.List(File1.ListIndex)
        strNumber = frmCustomer.Text1(8).Text
        'Call SendWinFax(strFax, strNumber)
    Case Else
    End Select
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Set rsMyRec = m_dbPrograming.OpenRecordset("User")
    Drive1.Drive = Left$(rsMyRec.Fields("FaxDirectory"), 3)
    Dir1.Path = Trim(rsMyRec.Fields("FaxDirectory"))
    File1.Path = Dir1.Path
    'File1.Pattern = "*.txt; *.doc"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMyRec.Close
    rsLanguage.Close
    Set frmFax = Nothing
End Sub


