VERSION 5.00
Begin VB.Form frmKeyAscii 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show keybord combinations"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Picture         =   "frmKeyAscii.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   4815
   End
End
Attribute VB_Name = "frmKeyAscii"
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
        .Fields("Msg1") = "The ASCII combination is: "
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Label1.Caption = rsLanguage.Fields("Msg1") & " " & KeyAscii
    Label2.Caption = Chr(KeyAscii)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmKeyAscii")
    m_iFormNo = 9
    DisableButtons 1
End Sub
Private Sub Form_Terminate()
    On Error Resume Next
    rsLanguage.Close
    Set frmKeyAscii = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_iFormNo = 0
    DisableButtons 2
    Set frmKeyAscii = Nothing
End Sub
