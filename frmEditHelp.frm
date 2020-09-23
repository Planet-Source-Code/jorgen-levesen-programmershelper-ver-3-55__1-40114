VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditHelp 
   BackColor       =   &H00000000&
   Caption         =   "Edit Help Text"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   Icon            =   "frmEditHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmEditHelp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Exit"
      Top             =   7680
      Width           =   8895
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":058C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":06E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":0840
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":099A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":0AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":0C4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":0DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":105C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":11B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEditHelp.frx":1310
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar2 
      Height          =   330
      Left            =   5040
      TabIndex        =   11
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Uline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Left justify"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mid"
            Object.ToolTipText     =   "Mid justify"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Right justify"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spell"
            Object.ToolTipText     =   "Spellcheck"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Picture"
            Object.ToolTipText     =   "Picture"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "Bold"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Uline"
            Object.ToolTipText     =   "Underline"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "Italic"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Left"
            Object.ToolTipText     =   "Left justify"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mid"
            Object.ToolTipText     =   "Mid justify"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Right"
            Object.ToolTipText     =   "Right justify"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Spell"
            Object.ToolTipText     =   "Spell check"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Picture"
            Object.ToolTipText     =   "Picture"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cmbLanguage 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   1
      Left            =   6840
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   2175
   End
   Begin VB.ComboBox cmbLanguage 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   2175
   End
   Begin VB.Data rsLanguage2 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data rsLanguage 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.ComboBox cmbFonts 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   1
      Left            =   5040
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   3195
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   8400
      TabIndex        =   6
      Text            =   "12"
      Top             =   120
      Width           =   540
   End
   Begin VB.ComboBox cmbFonts 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Index           =   0
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   3210
   End
   Begin VB.TextBox Text17 
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Text            =   "12"
      Top             =   120
      Width           =   540
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage"
      Height          =   6255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEditHelp.frx":146A
   End
   Begin RichTextLib.RichTextBox Text1 
      DataField       =   "Help"
      DataSource      =   "rsLanguage2"
      Height          =   6255
      Index           =   1
      Left            =   4680
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11033
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEditHelp.frx":14E4
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   4560
      X2              =   4560
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Your Language Help Text"
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
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "English Help Text"
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
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmEditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Control As Control, vBookmarkLang0() As Variant, vBookmarkLang1() As Variant
Dim m_retText As String
Dim m_wdApp As Word.Application
Dim rsLanguageWord0 As Recordset
Dim rsLanguageWord1 As Recordset
Dim rsFormLanguage As Recordset
Private Sub ReadText()
    On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsFormLanguage
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
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Update
    End With
End Sub

Private Sub LoadcmbLanguage()
    cmbLanguage(0).Clear
    cmbLanguage(1).Clear
    With rsLanguageWord0
        .MoveLast
        .MoveFirst
        ReDim vBookmarkLang0(.RecordCount)
        ReDim vBookmarkLang1(.RecordCount)
        Do While Not .EOF
            cmbLanguage(0).AddItem .Fields("SpellLanguage")
                cmbLanguage(0).ItemData(cmbLanguage(0).NewIndex) = cmbLanguage(0).ListCount - 1
                vBookmarkLang0(cmbLanguage(0).ListCount - 1) = .Bookmark
            cmbLanguage(1).AddItem .Fields("SpellLanguage")
                cmbLanguage(1).ItemData(cmbLanguage(1).NewIndex) = cmbLanguage(1).ListCount - 1
                vBookmarkLang1(cmbLanguage(1).ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    cmbLanguage(0).Text = cmbLanguage(0).List(cmbLanguage(0).ListIndex = 0)
    cmbLanguage(1).Text = cmbLanguage(1).List(cmbLanguage(1).ListIndex = 0)
End Sub
Private Sub btnExit_Click()
    Unload Me
End Sub
Private Sub cmbFonts_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        cmbFonts(0).FontName = cmbFonts(0).Text
        With Text1(0)
            .SelFontName = cmbFonts(0).Text
            .SelFontSize = CInt(Text17(0).Text)
        End With
    Case 1
        cmbFonts(1).FontName = cmbFonts(1).Text
        With Text1(1)
            .SelFontName = cmbFonts(1).Text
            .SelFontSize = CInt(Text17(1).Text)
        End With
    Case Else
    End Select
End Sub

Private Sub cmbLanguage_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        rsLanguageWord0.Bookmark = vBookmarkLang0(cmbLanguage(0).ItemData(cmbLanguage(0).ListIndex))
    Case 1
        rsLanguageWord1.Bookmark = vBookmarkLang1(cmbLanguage(1).ItemData(cmbLanguage(1).ListIndex))
    Case Else
    End Select
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    rsLanguage.Refresh
    rsLanguage2.Refresh
    
    With rsLanguage.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = "ENG" Then Exit Do
        .MoveNext
        Loop
    End With
    
    With rsLanguage2.Recordset
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then Exit Do
        .MoveNext
        Loop
    End With
    
    cmbFonts(0).Clear
    cmbFonts(1).Clear
    For n = 0 To Screen.FontCount - 1
        cmbFonts(0).AddItem Screen.Fonts(n)
        cmbFonts(1).AddItem Screen.Fonts(n)
    Next
    LoadcmbLanguage
    ReadText
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguageWord0 = m_dbLanguage.OpenRecordset("SpellLanguage")
    Set rsLanguageWord1 = m_dbLanguage.OpenRecordset("SpellLanguage")
    rsLanguage.DatabaseName = m_strProgramLng
    rsLanguage2.DatabaseName = m_strProgramLng
    rsLanguage.RecordSource = Trim(CStr(frmHelp.Label1.Caption))
    rsLanguage2.RecordSource = Trim(CStr(frmHelp.Label1.Caption))
    Set rsFormLanguage = m_dbLanguage.OpenRecordset("frmEditHelp")
    Unload frmHelp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguageWord0.Close
    rsLanguageWord1.Close
    rsLanguage.Recordset.Move 0
    rsLanguage.Recordset.Close
    rsLanguage2.Recordset.Move 0
    rsLanguage2.Recordset.Close
    rsFormLanguage.Close
    Erase vBookmarkLang0
    Erase vBookmarkLang1
    Set frmEditHelp = Nothing
End Sub


Private Sub Text1_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        Text17(0).Text = Text1(0).SelFontSize
    Case 1
        Text17(1).Text = Text1(1).SelFontSize
    Case Else
    End Select
End Sub


Private Sub Text1_GotFocus(Index As Integer)
' Ignore errors for controls without the TabStop property.
    On Error Resume Next
    ' Switch off the change of focus when pressing TAB.
    For Each Control In Controls
        Control.TabStop = False
    Next Control
End Sub
Private Sub Text1_LostFocus(Index As Integer)
' Ignore errors for controls without the TabStop property.
    On Error Resume Next
    ' Turn on the change of focus when pressing TAB.
    For Each Control In Controls
        Control.TabStop = True
    Next Control
End Sub


Private Sub Text1_SelChange(Index As Integer)
        On Error Resume Next
        Select Case Index
        Case 0
            cmbFonts(0).FontName = Text1(0).Font.Name
            Text17(0).Text = Text1(0).SelFontSize
        Case 1
            cmbFonts(1).FontName = Text1(1).Font.Name
            Text17(1).Text = Text1(1).SelFontSize
        Case Else
        End Select
End Sub
Private Sub Text17_Change(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        With Text1(0)
            .SelFontName = cmbFonts(0).Text
            .SelFontSize = CInt(Text17(0).Text)
        End With
    Case 1
        With Text1(1)
            .SelFontName = cmbFonts(1).Text
            .SelFontSize = CInt(Text17(1).Text)
        End With
    Case Else
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Bold"
        If Text1(0).SelBold = False Then
            Text1(0).SelBold = True
        Else
            Text1(0).SelBold = False
        End If
    Case "Uline"
        If Text1(0).SelUnderline = False Then
            Text1(0).SelUnderline = True
        Else
            Text1(0).SelUnderline = False
        End If
    Case "Italic"
        If Text1(0).SelItalic = False Then
            Text1(0).SelItalic = True
        Else
            Text1(0).SelItalic = False
        End If
    Case "Left"
        If Text1(0).SelAlignment = 0 Then
        Else
            Text1(0).SelAlignment = 0
        End If
    Case "Mid"
        If Text1(0).SelAlignment = 2 Then
            Text1(0).SelAlignment = 0
        Else
            Text1(0).SelAlignment = 2
        End If
    Case "Right"
        If Text1(0).SelAlignment = 1 Then
            Text1(0).SelAlignment = 0
        Else
            Text1(0).SelAlignment = 1
        End If
    Case "Copy"
        Clipboard.Clear
        Clipboard.SetText Text1(0).SelText
    Case "Paste"
        If Clipboard.GetFormat(1) Then
            Text1(0).SelText = Clipboard.GetText()
        ElseIf Clipboard.GetFormat(2) Then
            Text1(0).SelText = Clipboard.GetData()
        End If
    Case "Delete"
        Text1(0).SelText = ""
    Case "Spell"
        Call SpellCheck(Text1(0), CLng(rsLanguageWord0.Fields("SpellID")))
        Me.Show
    Case "Picture"
        SendMessage Text1(0).hWnd, WM_PASTE, 0, 0
    Case Else
    End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Bold"
        If Text1(1).SelBold = False Then
            Text1(1).SelBold = True
        Else
            Text1(1).SelBold = False
        End If
    Case "Uline"
        If Text1(1).SelUnderline = False Then
            Text1(1).SelUnderline = True
        Else
            Text1(1).SelUnderline = False
        End If
    Case "Italic"
        If Text1(1).SelItalic = False Then
            Text1(1).SelItalic = True
        Else
            Text1(1).SelItalic = False
        End If
    Case "Left"
        If Text1(1).SelAlignment = 0 Then
            Text1(1).SelAlignment = 0
        Else
            Text1(1).SelAlignment = 0
        End If
    Case "Mid"
        If Text1(1).SelAlignment = 2 Then
            Text1(1).SelAlignment = 0
        Else
            Text1(1).SelAlignment = 2
        End If
    Case "Right"
        If Text1(1).SelAlignment = 1 Then
            Text1(1).SelAlignment = 0
        Else
            Text1(1).SelAlignment = 1
        End If
    Case "Copy"
        Clipboard.Clear
        Clipboard.SetText Text1(1).SelText
    Case "Paste"
        If Clipboard.GetFormat(1) Then
            Text1(1).SelText = Clipboard.GetText()
        ElseIf Clipboard.GetFormat(2) Then
            Text1(1).SelText = Clipboard.GetData()
        End If
    Case "Delete"
        Text1(1).SelText = ""
    Case "Spell"
        Call SpellCheck(Text1(1), CLng(rsLanguageWord1.Fields("SpellID")))
        Me.Show
    Case "Picture"
        SendMessage Text1(1).hWnd, WM_PASTE, 0, 0
    Case Else
    End Select
End Sub
