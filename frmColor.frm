VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmColor 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colors"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   ControlBox      =   0   'False
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   Begin VB.Data rsColor 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Color"
      Top             =   3480
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid Grid1 
      Bindings        =   "frmColor.frx":0442
      Height          =   3735
      Left            =   120
      OleObjectBlob   =   "frmColor.frx":0458
      TabIndex        =   12
      Top             =   120
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Mix a color"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   7455
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   5280
         ScaleHeight     =   945
         ScaleWidth      =   2025
         TabIndex        =   11
         Top             =   240
         Width           =   2055
      End
      Begin VB.HScrollBar H3 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   4
         Top             =   960
         Width           =   2895
      End
      Begin VB.HScrollBar H2 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   3
         Top             =   600
         Width           =   2895
      End
      Begin VB.HScrollBar H1 
         Height          =   255
         Left            =   1080
         Max             =   255
         TabIndex        =   2
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Blue:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Green:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Red:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   5880
      ScaleHeight     =   3705
      ScaleWidth      =   1665
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecordBookmark() As Variant
Dim rsLanguage As Recordset

Private Sub ShowText()
Dim strHelp As String
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
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
            
            'this language was not found, make it. Find the English text first
            strHelp = " "
            .MoveFirst
            Do While Not .EOF
                If .Fields("Language") = "ENG" Then
                    If Not IsNull(.Fields("Help")) Then
                        strHelp = .Fields("Help")
                        Exit Do
                    End If
                End If
            .MoveNext
            Loop
            
        .MoveFirst
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("label1(2)") = Label1(2).Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Help") = strHelp
        .Update
    End With
End Sub

Private Sub Grid1_Click()
    With rsColor.Recordset
        Picture1.BackColor = RGB(CLng(.Fields("RedValue")), CLng(.Fields("GreenValue")), CLng(.Fields("BlueValue")))
        H1.Value = CLng(.Fields("RedValue"))
        H2.Value = CLng(.Fields("GreenValue"))
        H3.Value = CLng(.Fields("BlueValue"))
    End With
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsColor.Refresh
    ShowText
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsColor.DatabaseName = m_strPrograming
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmColor")
    m_iFormNo = 35
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LOadForm"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsColor.Recordset.Close
    rsLanguage.Close
    m_iFormNo = 0
    Erase RecordBookmark
    Set frmColor = Nothing
End Sub
Private Sub H1_Change()
Label2(0).Caption = H1.Value
Picture2.BackColor = RGB(H1.Value, H2.Value, H3.Value)
End Sub

Private Sub H2_Change()
Label2(1).Caption = H2.Value
Picture2.BackColor = RGB(H1.Value, H2.Value, H3.Value)
End Sub

Private Sub H3_Change()
Label2(2).Caption = H3.Value
Picture2.BackColor = RGB(H1.Value, H2.Value, H3.Value)
End Sub

