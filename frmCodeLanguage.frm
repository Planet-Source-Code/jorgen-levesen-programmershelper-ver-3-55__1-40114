VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCodeLanguage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Language"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data rsCodeLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeSnippets.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Language"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   120
      Picture         =   "frmCodeLanguage.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   5040
      Width           =   4335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCodeLanguage.frx":014A
      Height          =   4455
      Left            =   240
      OleObjectBlob   =   "frmCodeLanguage.frx":0167
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   4440
      X2              =   120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   4440
      X2              =   120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   4440
      X2              =   4440
      Y1              =   120
      Y2              =   4800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   4800
   End
End
Attribute VB_Name = "frmCodeLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
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

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsCodeLanguage.Refresh
    ReadText
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsCodeLanguage.DatabaseName = m_strCodeSnippet
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCodeLanguage")
    Dither Me
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbExclamation, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCodeLanguage.Recordset.Close
    rsLanguage.Close
    Set frmCodeLanguage = Nothing
End Sub
