VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDueDate 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Payment Due Date Text"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   Begin VB.Data rsDueDate 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DueDateText"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmDueDate.frx":0000
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frmDueDate.frx":0018
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmDueDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error Resume Next
    rsDueDate.Refresh
    DisableButtons 1
End Sub

Private Sub Form_Load()
Dim sDir As String
    On Error GoTo errForm_Load
    Me.Move 0, 0
    sDir = App.Path & "\Programming.mdb"
    rsDueDate.DatabaseName = sDir
    m_iFormNo = 6
    Me.AutoRedraw = True
    a = (255 / Me.ScaleHeight)
    For i = 0 To Me.ScaleHeight
        Me.Line (0, i)-(Me.ScaleWidth, i), RGB(0, 0, a * i)
    Next
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
    rsDueDate.Recordset.Close
    m_iFormNo = 0
    DisableButtons 2
    Set frmDueDate = Nothing
End Sub
