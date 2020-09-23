VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCountry 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Countries"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmCountry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton btnPaste 
      Caption         =   "&Paste"
      Height          =   375
      Left            =   2160
      TabIndex        =   13
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "&Open"
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton btnDeletePicture 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2160
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Country"
      DataSource      =   "rsCountry"
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "Currency"
      DataSource      =   "rsCountry"
      Height          =   285
      Left            =   4200
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "ExchangeRate"
      DataSource      =   "rsCountry"
      Height          =   285
      Left            =   4200
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CountryFix"
      DataSource      =   "rsCountry"
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CountryPrefix"
      DataSource      =   "rsCountry"
      Height          =   285
      Left            =   4200
      TabIndex        =   1
      Top             =   1920
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   4515
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Data rsCountry 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Country"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   960
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flag:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   14
      Top             =   2400
      Width           =   345
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   1920
      X2              =   1920
      Y1              =   2520
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   6240
      X2              =   6240
      Y1              =   2520
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   1920
      X2              =   6240
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   2760
      X2              =   6240
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Image Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      DataField       =   "CountryFlag"
      DataSource      =   "rsCountry"
      Height          =   1815
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   6240
      X2              =   6240
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1920
      X2              =   6240
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   1920
      X2              =   6240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Currency:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exchange Rate:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country Short:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   7
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone Prefix:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarksCountry() As Variant
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Dim bNewRecord As Boolean
Private Sub LoadBackground()
    Picture2.Visible = False
    Picture2.AutoRedraw = True
    Picture2.AutoSize = True
    Picture2.BorderStyle = 0
    Picture2.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture2
    For i = 0 To 5
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 7
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
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
                For i = 0 To 5
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("btnPaste")) Then
                    .Fields("btnPaste") = btnPaste.Caption
                Else
                    btnPaste.Caption = .Fields("btnPaste")
                End If
                If IsNull(.Fields("btnOpen")) Then
                    .Fields("btnOpen") = btnOpen.Caption
                Else
                    btnOpen.Caption = .Fields("btnOpen")
                End If
                If IsNull(.Fields("btnDeletePicture")) Then
                    .Fields("btnDeletePicture") = btnDeletePicture.Caption
                Else
                    btnDeletePicture.Caption = .Fields("btnDeletePicture")
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
        For i = 0 To 5
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnPaste") = btnPaste.Caption
        .Fields("btnOpen") = btnOpen.Caption
        .Fields("btnDeletePicture") = btnDeletePicture.Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsCountry.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarksCountry(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("Country")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarksCountry(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Public Sub DeleteRecord()
Dim DgDef, Msg, response, Title
    If bNewRecord Then Exit Sub
    On Error GoTo ErrbDelete_Click
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    Title = "Delete Record"
    Msg = "Do you really want to delete this Country ?"
    Beep
    response = MsgBox(Msg, DgDef, Title)
    If response = IdNo Then
        Exit Sub
    End If
    On Error Resume Next
    'delete this Country
    rsCountry.Recordset.Delete
    Beep
    MsgBox "Country is deleted !!"
    Exit Sub
    
ErrbDelete_Click:
    Beep
    MsgBox Err.Description, vbCritical, "Delete Country"
    Resume ErrbDelete_Click2
ErrbDelete_Click2:
End Sub
Public Sub NewRecord()
    On Error Resume Next
    If bNewRecord Then Exit Sub
    rsCountry.Recordset.AddNew
    bNewRecord = True
    Text2.SetFocus
End Sub

Private Sub btnDeletePicture_Click()
    Picture1.Picture = LoadPicture()
End Sub

Private Sub btnOpen_Click()
    With CMD1
        .FileName = ""
        .DialogTitle = "Load Picture from disk"
        .Filter = "(*.bmp)|*.bmp|(*.pcx)|*.pcx|(*.jpg)|*.jpg"
        .FilterIndex = 1
        .ShowOpen
        Picture1.Picture = LoadPicture(.FileName)
    End With
End Sub

Private Sub btnPaste_Click()
    Picture1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsCountry.Refresh
    LoadList1
    List1.ListIndex = 0
    ReadText
    DisableButtons 2
    frmMDI.Toolbar1.Buttons(8).Enabled = False
    frmMDI.Toolbar1.Buttons(9).Enabled = False
    LoadBackground
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsCountry.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCountry")
    m_iFormNo = 4
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCountry.Recordset.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    DisableButtons 1
    Set frmCountry = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsCountry.Recordset.Bookmark = bookmarksCountry(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text2_LostFocus()
    If bNewRecord Then
        On Error GoTo errText2_Click
        With rsCountry.Recordset
            .Fields("Country") = Text2.Text
            .Update
            LoadList1
            .Bookmark = .LastModified
        End With
        bNewRecord = False
        Text3.SetFocus
    End If
    Exit Sub
    
errText2_Click:
    Beep
    MsgBox Err.Description, vbInformation, "New Record"
    bNewRecord = False
    Resume errText2_Click2
errText2_Click2:
End Sub

