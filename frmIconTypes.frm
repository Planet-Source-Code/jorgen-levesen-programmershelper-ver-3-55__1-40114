VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmIconTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Types"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   120
      Picture         =   "frmIconTypes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   5280
      Width           =   4335
   End
   Begin VB.Data rsIconType 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\ProgramIco.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "IconType"
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmIconTypes.frx":014A
      Height          =   5055
      Left            =   120
      OleObjectBlob   =   "frmIconTypes.frx":0163
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmIconTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Private Sub ReadText()
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
    rsIconType.Refresh
    ReadText
    TileForm Me, Picture1
End Sub

Private Sub Form_Load()
Dim sName As String, dbTemp As Database
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Dither Me
    sName = App.Path & "\CodeIco.mdb"
    Set dbTemp = OpenDatabase(sName)
    rsIconType.DatabaseName = sName
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmIconTypes")
    m_iFormNo = 37
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsIconType.Recordset.Close
    rsLanguage.Close
    Set frmIconTypes = Nothing
End Sub
