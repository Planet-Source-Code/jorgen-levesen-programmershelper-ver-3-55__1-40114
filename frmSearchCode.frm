VERSION 5.00
Begin VB.Form frmSearchCode 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnShowSelection 
      Caption         =   "&Show Selection"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5640
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4905
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton btnSearch 
      Height          =   375
      Left            =   4800
      Picture         =   "frmSearchCode.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Start Search"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Found Entries:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Search String:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSearchCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBokmark() As Variant
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("label2")
                End If
                If IsNull(.Fields("btnShowSelection")) Then
                    .Fields("btnShowSelection") = btnShowSelection.Caption
                Else
                    btnShowSelection.Caption = .Fields("btnShowSelection")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
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
        .Fields("label2") = Label2.Caption
        .Fields("btnShowSelection") = btnShowSelection.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSearch_Click()
Dim vComp As Variant
    On Error Resume Next
    List1.Clear
    With frmCodeSnippets.rsCodeSnippet.Recordset
        .MoveLast
        .MoveFirst
        ReDim vBokmark(.RecordCount)
        Do While Not .EOF
            vComp = InStr(CStr(.Fields("CodeSnippet")), CStr(txtSearch.Text))
             If vComp > 0 Then
                List1.AddItem .Fields("CodeText")
                List1.ItemData(List1.NewIndex) = List1.ListCount - 1
                vBokmark(List1.ListCount - 1) = .Bookmark
            End If
        .MoveNext
        Loop
    End With
End Sub
Private Sub btnShowSelection_Click()
    On Error Resume Next
    frmCodeSnippets.rsCodeSnippet.Recordset.Bookmark = _
        vBokmark(List1.ItemData(List1.ListIndex))
    frmCodeSnippets.cmbCodeType.Text = frmCodeSnippets.rsCodeSnippet.Recordset.Fields("CodeType")
    DoEvents
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
    frmCodeSnippets.SelectAllCode
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmSearchCode")
    txtSearch.SetFocus
    m_iFormNo = 21
    DisableButtons 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    m_iFormNo = 0
    DisableButtons 2
    rsLanguage.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    frmCodeSnippets.List1.Clear
    Erase vBokmark
    Set frmSearchCode = Nothing
End Sub


