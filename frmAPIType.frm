VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmAPIType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "API Types"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmAPIType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      Picture         =   "frmAPIType.frx":030A
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data rsAPIType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeSnippets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "APIType"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   120
      Picture         =   "frmAPIType.frx":0680
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   6720
      Width           =   4575
   End
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmAPIType.frx":07CA
      Height          =   6015
      Left            =   360
      OleObjectBlob   =   "frmAPIType.frx":07E2
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   4680
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   4680
      X2              =   120
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   4680
      X2              =   4680
      Y1              =   120
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   6600
   End
End
Attribute VB_Name = "frmAPIType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsAPIType.Refresh
    LoadBackground
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsAPIType.DatabaseName = m_strCodeSnippet
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbExclamation, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsAPIType.Recordset.Close
    rsUser.Close
    Set frmAPIType = Nothing
End Sub
