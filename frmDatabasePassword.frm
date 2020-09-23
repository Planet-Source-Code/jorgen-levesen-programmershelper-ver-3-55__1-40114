VERSION 5.00
Begin VB.Form frmDatabasePassword 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password?"
   ClientHeight    =   1065
   ClientLeft      =   3525
   ClientTop       =   2070
   ClientWidth     =   3360
   ControlBox      =   0   'False
   Icon            =   "frmDatabasePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2790
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "This Database requires a password. Please enter below..."
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   600
      TabIndex        =   1
      Top             =   75
      Width           =   2565
   End
End
Attribute VB_Name = "frmDatabasePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Public pstrPassword As String
Public pblnCancel As Boolean
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
                    .Fields(Label1) = Label1.Caption
                Else
                    Label1.Caption = .Fields(Label1)
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
        .Update
    End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmDatabasePassword")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmDatabasePassword = Nothing
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Not Trim(txtPassword) = "" Then
        pstrPassword = Trim(txtPassword)
        txtPassword = ""
        pblnCancel = False
        Me.Hide
    End If
    If KeyAscii = vbKeyEscape Then
        pstrPassword = ""
        txtPassword = ""
        pblnCancel = True
        Unload Me
    End If
End Sub
