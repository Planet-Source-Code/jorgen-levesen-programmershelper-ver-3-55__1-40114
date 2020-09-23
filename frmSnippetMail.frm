VERSION 5.00
Begin VB.Form frmSnippetMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Snippet Mail"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmSnippetMail.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton btnSendMail 
      Height          =   495
      Left            =   2640
      Picture         =   "frmSnippetMail.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Send this Mail"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Option1 
         Caption         =   "Send this Code Snippet"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Send Mail to Snippet Author"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmSnippetMail"
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
                If IsNull(.Fields("Option1(0)")) Then
                    .Fields("Option1(0)") = Option1(0).Caption
                Else
                    Option1(0).Caption = .Fields("Option1(0)")
                End If
                If IsNull(.Fields("Option1(1)")) Then
                    .Fields("Option1(1)") = Option1(1).Caption
                Else
                    Option1(1).Caption = .Fields("Option1(1)")
                End If
                If IsNull(.Fields("btnSendMail")) Then
                    .Fields("btnSendMail") = btnSendMail.Caption
                Else
                    btnSendMail.Caption = .Fields("btnSendMail")
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
        .Fields("Option1(0)") = Option1(0).Caption
        .Fields("Option1(1)") = Option1(1).Caption
        .Fields("btnSendMail") = btnSendMail.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSendMail_Click()
Dim IsValid As Boolean
Dim InvalidReason As String
    
    On Error GoTo errSendMail
    If Option1(0).Value = True Then
        IsValid = IsEMailAddress(Trim(frmCodeSnippets.Text4.Text), InvalidReason)
        If Not IsValid Then
            MsgBox "Invalid mail address, the reason given is: " & InvalidReason
            Exit Sub
        End If
        With frmEmail
            .List1.AddItem Trim(frmCodeSnippets.Text4.Text)
            .Show vbModal
        End With
    Else
        With frmEmail
            .Text1.Text = frmCodeSnippets.Text2.Text    'the code type
            .RichText1.Text = frmCodeSnippets.Text1.Text    'the code snippet
            .List1.Visible = False
            .Text2.Visible = True
            .Show vbModal
        End With
    End If
    Unload Me
    Exit Sub
    
errSendMail:
    Beep
    MsgBox Err.Description, vbExclamation, "Send Mail"
    Err.Clear
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmSnippetMail")
    ReadText
    If Len(frmCodeSnippets.Text4.Text) = 0 Then
        Option1(0).Enabled = False
        Option1(1).Value = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmSnippetMail = Nothing
End Sub
