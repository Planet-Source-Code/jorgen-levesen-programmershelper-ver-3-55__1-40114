VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmCodeType 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Types"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox cboLanguage 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cboLanguage 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   1
      Left            =   7560
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.Data rsMyCodeTypes 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeType"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsCodeTypes 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CodeType"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmCodeType.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   6360
      Width           =   9615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCodeType.frx":014A
      Height          =   5295
      Index           =   0
      Left            =   240
      OleObjectBlob   =   "frmCodeType.frx":0164
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmCodeType.frx":0B3A
      Height          =   5295
      Index           =   1
      Left            =   5160
      OleObjectBlob   =   "frmCodeType.frx":0B56
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   9120
      TabIndex        =   8
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   9720
      X2              =   9720
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   4920
      X2              =   4920
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   9720
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1680
      X2              =   7560
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code Language:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Code Language:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "frmCodeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCodeLanguage As Recordset
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
    For i = 0 To 4
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Sub LoadCodeLanguage()
    With rsCodeLanguage
        .MoveFirst
        Do While Not .EOF
            cboLanguage(0).AddItem .Fields("Language")
            cboLanguage(1).AddItem .Fields("Language")
        .MoveNext
        Loop
    End With
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
                If IsNull(.Fields("Form")) Then
                    .Fields("Form") = Me.Caption
                Else
                    Me.Caption = .Fields("Form")
                End If
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1")
                    Label1(1).Caption = .Fields("label1")
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
        .Fields("label1") = Label1(0).Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cboLanguage_Click(Index As Integer)
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM CodeType WHERE Trim(CodeLanguage) ="
    Sql = Sql & Chr(34) & Trim(cboLanguage(Index).Text) & Chr(34)
    
    Select Case Index
    Case 0
        With rsCodeTypes
            .RecordSource = Sql
            .Refresh
            If Not .Recordset.EOF And Not .Recordset.BOF Then
                .Recordset.MoveFirst
            End If
        End With
    Case 1
        With rsMyCodeTypes
            .RecordSource = Sql
            .Refresh
            If Not .Recordset.EOF And Not .Recordset.BOF Then
                .Recordset.MoveFirst
            End If
        End With
    Case Else
    End Select
End Sub


Private Sub DBGrid1_OnAddNew(Index As Integer)
    DBGrid1(Index).Columns(0).Text = cboLanguage(Index).Text
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsCodeTypes.Refresh
    rsMyCodeTypes.Refresh
    LoadCodeLanguage
    cboLanguage(0).ListIndex = 0
    cboLanguage(1).ListIndex = 0
    DisableButtons 1
    ReadText
    LoadBackground
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsCodeTypes.DatabaseName = m_strCodeSnippet
    Label1(2).Caption = ExtractFileName(m_strCodeSnippet)
    rsMyCodeTypes.DatabaseName = m_strMyCodeSnippet
    Label1(3).Caption = ExtractFileName(m_strMyCodeSnippet)
    Set rsCodeLanguage = m_dbCodeSnippet.OpenRecordset("Language")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCodeType")
    m_iFormNo = 3
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Error$, vbCritical, "Form Load"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCodeTypes.Recordset.Close
    rsMyCodeTypes.Recordset.Close
    rsCodeLanguage.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    DisableButtons 2
    Set frmCodeType = Nothing
End Sub


