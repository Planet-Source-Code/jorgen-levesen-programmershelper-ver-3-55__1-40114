VERSION 5.00
Begin VB.Form frmWriteProgRep 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Program Times"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   8
      Text            =   "0"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   2160
      Width           =   2055
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Just Print Shown Project"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Print All Projects"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmWriteProgRep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton btnPrint 
      Height          =   495
      Left            =   1920
      Picture         =   "frmWriteProgRep.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   5
      X1              =   3240
      X2              =   3240
      Y1              =   2880
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   120
      X2              =   120
      Y1              =   2880
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   5
      X1              =   120
      X2              =   3240
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   4
      X1              =   120
      X2              =   3240
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Read Records:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Write Records:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   3240
      X2              =   3240
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   120
      X2              =   120
      Y1              =   1200
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   3
      X1              =   120
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   2
      X1              =   120
      X2              =   3240
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "From date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   120
      X2              =   3240
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   120
      X2              =   3240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   1
      X1              =   3240
      X2              =   3240
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   960
   End
End
Attribute VB_Name = "frmWriteProgRep"
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
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("label1(1)")
                End If
                If IsNull(.Fields("label1(2)")) Then
                    .Fields("label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("label1(2)")
                End If
                If IsNull(.Fields("label1(3)")) Then
                    .Fields("label1(3)") = Label1(3).Caption
                Else
                    Label1(3).Caption = .Fields("label1(3)")
                End If
                'If IsNull(.Fields("btnExit")) Then
                    '.Fields("btnExit") = btnExit.Caption
                'Else
                    'btnExit.Caption = .Fields("btnExit")
                'End If
                'If IsNull(.Fields("btnPrint")) Then
                    '.Fields("btnPrint") = btnPrint.Caption
                'Else
                    'btnPrint.Caption = .Fields("btnPrint")
                'End If
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
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("label1(2)") = Label1(2).Caption
        .Fields("label1(3)") = Label1(3).Caption
        '.Fields("btnExit") = btnExit.Caption
        '.Fields("btnPrint") = btnPrint.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    On Error Resume Next
    If Option1(0).Value = True Then
        dateFromDate = Format(CDate(Text1(0).Text), "dd.mm.yyyy")
        dateToDate = Format(CDate(Text1(1).Text), "dd.mm.yyyy")
        frmProgramming.WriteAllProjects
    Else
        frmProgramming.WriteProject
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
    Text1(0).Text = "01.01." & Year(Now)
    Text1(1).Text = Format(Now, "dd.mm.yyyy")
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmWriteProgRep")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmWriteProgRep = Nothing
End Sub
