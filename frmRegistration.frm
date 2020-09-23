VERSION 5.00
Begin VB.Form frmRegistration 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Registration"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   14
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   12
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   6
      Left            =   3240
      TabIndex        =   8
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   7
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   8
      Left            =   3240
      TabIndex        =   6
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Check1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   12
      Left            =   3240
      MaxLength       =   16
      TabIndex        =   3
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   13
      Left            =   3240
      TabIndex        =   2
      Top             =   4320
      Width           =   2655
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmRegistration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit"
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton btnSendEmail 
      Height          =   495
      Left            =   4920
      Picture         =   "frmRegistration.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Send Mail"
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your data:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   120
      X2              =   6720
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   6720
      X2              =   6720
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   5280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   1320
      X2              =   6720
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Your Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   25
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   24
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   23
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Town:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   22
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   21
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   20
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   19
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Programme Installation Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   18
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Send Me regular E-mail information:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   480
      TabIndex        =   17
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Program Version No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   16
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Version:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   480
      TabIndex        =   15
      Top             =   4320
      Width           =   2655
   End
End
Attribute VB_Name = "frmRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32s = 0


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'Maintenance string For PSS usage
    End Type

Dim rsMyRecord As Recordset
Dim rsSupplier As Recordset
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 11
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Function FindMyOS()
'by:  Brandon9 13

    Dim MsgEnd As String
    Dim junk
    Dim osvi As OSVERSIONINFO
    osvi.dwOSVersionInfoSize = 148
    junk = GetVersionEx(osvi)

    If junk <> 0 Then

        Select Case osvi.dwPlatformId
            Case VER_PLATFORM_WIN32s '0
            Text1(13).Text = "Microsoft Win32s"
            Case VER_PLATFORM_WIN32_WINDOWS '1
            If ((osvi.dwMajorVersion > 4) Or _
            ((osvi.dwMajorVersion = 4) And (osvi.dwMinorVersion > 0))) Then
            Text1(13).Text = "Microsoft Windows 98"
        Else
            Text1(13).Text = "Microsoft Windows 95"
        End If
        Case VER_PLATFORM_WIN32_NT '2
        If osvi.dwMajorVersion <= 4 Then _
        Text1(13).Text = "Microsoft Windows NT"
        If osvi.dwMajorVersion = 5 Then _
        Text1(13).Text = "Microsoft Windows 2000"
        If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 1 Then _
        Text1(13).Text = "Microsoft Windows XP Build " & osvi.dwBuildNumber
    End Select
End If
FindMyOS = Text1(13).Text
End Function
Private Sub InsertFields()
    On Error Resume Next
    With rsMyRecord
        Text1(0).Text = Format(.Fields("CompanyName"))
        Text1(1).Text = Format(.Fields("CompanyAddress1"))
        Text1(2).Text = Format(.Fields("CompanyAddress2"))
        Text1(4).Text = Format(.Fields("CompanyZip"))
        Text1(5).Text = Format(.Fields("CompanyTown"))
        Text1(6).Text = Format(.Fields("CompanyCountry"))
        Text1(7).Text = Format(.Fields("CompanyPrefixPhone")) & " " & Format(.Fields("CompanyPhoneNo"))
        Text1(8).Text = Format(.Fields("CompanyEMail"))
        Text1(9).Text = Format(Now, "dd.mm.yyyy")
    End With
End Sub


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
                For i = 0 To 11
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("btnSendEmail")) Then
                    .Fields("btnSendEmail") = btnSendEmail.ToolTipText
                Else
                    btnSendEmail.ToolTipText = .Fields("btnSendEmail")
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
        For i = 0 To 11
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnSendEmail") = btnSendEmail.ToolTipText
        .Fields("btnExit") = btnExit.ToolTipText
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSendEmail_Click()
Dim Subject As String, Recipient As String, Message As String
    Subject = "Programming registration"
    Recipient = rsSupplier.Fields("SupplierEmailySysResponse")
    Message = rsLanguage.Fields("Msg1") & vbCrLf & vbCrLf
    Message = Message & Label1(0).Caption & vbTab & vbTab & Trim(Text1(0).Text) & vbCrLf
    Message = Message & Label1(1).Caption & vbTab & Trim(Text1(1).Text) & vbCrLf
    Message = Message & "                                 " & vbTab & Trim(Text1(2).Text) & vbCrLf
    Message = Message & "                                 " & vbTab & Trim(Text1(3).Text) & vbCrLf
    Message = Message & Label1(2).Caption & vbTab & vbTab & Trim(Text1(4).Text) & vbCrLf
    Message = Message & Label1(3).Caption & vbTab & vbTab & vbTab & Trim(Text1(5).Text) & vbCrLf
    Message = Message & Label1(4).Caption & vbTab & vbTab & vbTab & Trim(Text1(6).Text) & vbCrLf
    Message = Message & Label1(5).Caption & vbTab & Trim(Text1(7).Text) & vbCrLf
    Message = Message & Label1(6).Caption & vbTab & vbTab & Trim(Text1(8).Text) & vbCrLf
    Message = Message & Label1(11).Caption & vbTab & vbTab & Trim(Text1(11).Text) & vbCrLf
    Message = Message & Label1(15).Caption & vbTab & vbTab & Trim(Text1(12).Text) & vbCrLf
    Message = Message & Label1(16).Caption & vbTab & vbTab & Trim(Text1(13).Text) & vbCrLf
    If Len(Text1(9).Text) <> 0 Then
        Message = Message & Label1(7).Caption & vbTab & vbTab & Trim(Text1(9).Text) & vbCrLf & vbCrLf
    Else
        Message = Message & Label1(7).Caption & vbTab & vbTab & Format(Now, "dd.mm.yyyy") & vbCrLf & vbCrLf
    End If
    If Check1.Value = 1 Then
        Message = Message & Label1(9).Caption & vbTab & rsLanguage.Fields("Yes") & vbCrLf
    Else
        Message = Message & Label1(9).Caption & vbTab & rsLanguage.Fields("No") & vbCrLf
    End If
    Call SendOutlookMail(Subject, Recipient, Message)
    Unload Me
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    Text1(9).Text = Format(Now, "dd.mm.yyyy")
    Text1(12).Text = App.Major & "." & App.Minor & "." & App.Revision
    InsertFields
    ReadText
    LoadBackground
    FindMyOS
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsSupplier = m_dbPrograming.OpenRecordset("Supplier")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmRegistration")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Close
    rsUser.Close
    rsLanguage.Close
    Set frmRegistration = Nothing
End Sub
