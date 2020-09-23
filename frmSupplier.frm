VERSION 5.00
Begin VB.Form frmSupplier 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Program Supplier"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      Picture         =   "frmSupplier.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   23
      Top             =   4320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data rsSupplier 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Supplier"
      Top             =   360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierName"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   0
      Left            =   2400
      MaxLength       =   80
      TabIndex        =   12
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierAddr1"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   11
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierAddr2"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   2
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   10
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierAddr3"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   3
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierZip"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   4
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierTown"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   5
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierCountry"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   6
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierContact"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   7
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierPhone"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   8
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierFax"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   9
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierEmail"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   10
      Left            =   2400
      MaxLength       =   100
      TabIndex        =   2
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "SupplierEmailySysResponse"
      DataSource      =   "rsSupplier"
      Height          =   285
      Index           =   11
      Left            =   2400
      MaxLength       =   100
      TabIndex        =   1
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton btnExit 
      Height          =   375
      Left            =   4440
      Picture         =   "frmSupplier.frx":01CE
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   3
      X1              =   5880
      X2              =   5880
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   2
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   1
      X1              =   240
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      Index           =   0
      X1              =   240
      X2              =   5880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Company:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   21
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   20
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Town:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   19
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   18
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   480
      TabIndex        =   17
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   480
      TabIndex        =   16
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No.:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   480
      TabIndex        =   15
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   480
      TabIndex        =   14
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 9
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
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
                For i = 0 To 9
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(1 + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
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
        For i = 0 To 9
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("btnExit") = btnExit.ToolTipText
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsSupplier.Refresh
    ReadText
    LoadBackground
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsSupplier.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmSupplier")
    Dither Me
    m_iFormNo = 24
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsSupplier.Recordset.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    Set frmSupplier = Nothing
End Sub
