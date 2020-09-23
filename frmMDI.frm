VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Programing..."
   ClientHeight    =   7125
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9090
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data rsUser 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\CodeMaster.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "User"
      Top             =   360
      Visible         =   0   'False
      Width           =   9090
   End
   Begin Project1.CtlVerticalMenu Menu1 
      Align           =   3  'Align Left
      Height          =   6075
      Left            =   0
      TabIndex        =   2
      Top             =   675
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   10716
      MenuCaption1    =   "Menu1"
      MenuItemIcon11  =   "frmMDI.frx":08CA
      BackColor       =   14737632
      MenuItemForeColor=   0
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   2760
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   24
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1600
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit this system"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "New"
            Object.ToolTipText     =   "New Record"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Save Record"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete Record"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy to clipboard"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Print ...."
            ImageIndex      =   14
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Write"
            Object.ToolTipText     =   "Write Letter"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Fax"
            Object.ToolTipText     =   "Send Fax"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Mail"
            Object.ToolTipText     =   "Email"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Web"
            Object.ToolTipText     =   "Open Internet"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Open"
            Object.ToolTipText     =   "Open Project"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Spell"
            Object.ToolTipText     =   "Spell Check"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "View"
            Object.ToolTipText     =   "View Report"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "About ..."
            ImageIndex      =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1200
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         DataField       =   "BackgroundPicture"
         DataSource      =   "rsUser"
         Height          =   375
         Left            =   7560
         ScaleHeight     =   315
         ScaleWidth      =   555
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "frmMDI.frx":0BE4
            TextSave        =   "09:47"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "23.10.2002"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":10AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1360
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":14BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":433E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":4A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":4D2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":5044
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":5716
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":5DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":64BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":6B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":933E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9498
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":9672
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFiles 
      Caption         =   "&Files"
      Begin VB.Menu mnuUser 
         Caption         =   "User"
      End
      Begin VB.Menu mnuSpace1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuMich 
      Caption         =   "&Mich."
      Begin VB.Menu mnuKeyAscii 
         Caption         =   "Key ASCII"
      End
      Begin VB.Menu mnuPasword 
         Caption         =   "Computer Pasword"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Print"
      Begin VB.Menu mnuPrintSetUp 
         Caption         =   "Print Set-Up"
      End
      Begin VB.Menu mnuSpace2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintWord 
         Caption         =   "Use Word"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPrintDefault 
         Caption         =   "Use Default Printer"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuDatabaseInternet 
         Caption         =   "Internet Links"
      End
      Begin VB.Menu mnuSpace3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDatabaseLinkType 
         Caption         =   "Link Types"
      End
      Begin VB.Menu mnuDatabaseCodeLanguage 
         Caption         =   "Code Language"
      End
      Begin VB.Menu mnuDatabaseCodeType 
         Caption         =   "Code Type"
      End
      Begin VB.Menu mnuAPIType 
         Caption         =   "API Type"
      End
      Begin VB.Menu mnuIconType 
         Caption         =   "Icon Type"
      End
      Begin VB.Menu mnuSpace6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "Colors"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      Visible         =   0   'False
      Begin VB.Menu mnuFormatCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuFormatPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu munSpace5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatBold 
         Caption         =   "Bold"
      End
      Begin VB.Menu mnuFormatItalic 
         Caption         =   "Italic"
      End
      Begin VB.Menu mnuFormatUline 
         Caption         =   "Underline"
      End
      Begin VB.Menu mnuSpace4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatFont 
         Caption         =   "Font"
         Begin VB.Menu mnuFormatFontName 
            Caption         =   "Name"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatFontSize 
         Caption         =   "Font Size"
         Begin VB.Menu mnuFormatFontSizeNo 
            Caption         =   "Size"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFormatColor 
         Caption         =   "Color"
         Begin VB.Menu mnuFormatColorNo 
            Caption         =   "Color"
            Index           =   0
         End
      End
      Begin VB.Menu mnuSpace7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatPastePic 
         Caption         =   "Paste Picture"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolFirst As Boolean, bookColor() As Variant
Dim hSysMenu As Long      ' This is the system menu's handle.
Dim zMENU As MENUITEMINFO ' This is a structure that is used for the modification of the
Dim dbTemp As Database
Dim rsColor As Recordset
Dim rsLanguage As Recordset
Dim iCount As Integer
Dim boolOwnCode As Boolean
Private Sub IsCodeSnippet()
Dim sName As String, strOldName As String
    If IsNull(rsUser.Recordset.Fields("OwnSnippetName")) Then 'users own snippet database
        sName = InputBox(rsLanguage.Fields("Msg1") & vbCrLf & _
                rsLanguage.Fields("Msg2"), "My Own Code Snippets", "MyOwnSnippets")
        m_strMyCodeSnippet = App.Path & "\" & sName & ".mdb"
        strOldName = "MyOwnSnippets.mdb"
        Name strOldName As m_strMyCodeSnippet
        With rsUser.Recordset
            .Edit
            .Fields("OwnSnippetName") = m_strMyCodeSnippet
            .Update
        End With
    Else
        m_strMyCodeSnippet = rsUser.Recordset.Fields("OwnSnippetName")
    End If
    Set m_dbMyCodeSnippet = OpenDatabase(m_strMyCodeSnippet)
    boolOwnCode = True
End Sub
Private Sub LoadColor()
   'add color , font and font size to pop-up-menu
    On Error Resume Next
    iCount = 0
    With rsColor
        .MoveLast
        .MoveFirst
        ReDim bookColor(.RecordCount)
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                Me.mnuFormatColorNo(0).Caption = .Fields("ColorText")
                bookColor(iCount) = .Bookmark
                Exit Do
            End If
        .MoveNext
        Loop
        
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                iCount = iCount + 1
                Load Me.mnuFormatColorNo(iCount)
                Me.mnuFormatColorNo(iCount).Caption = .Fields("ColorText")
                bookColor(iCount) = .Bookmark
            End If
        .MoveNext
        Loop
    End With
    
    'load fonts
    On Error Resume Next
    Me.mnuFormatFontName(0).Caption = Screen.Fonts(0)
    For i = 1 To Screen.FontCount - 1
        Load Me.mnuFormatFontName(i)
        Me.mnuFormatFontName(i).Caption = Screen.Fonts(i)
    Next
    
    'load font size
    Me.mnuFormatFontSizeNo(0).Caption = 8
    For i = 9 To 48
        Load Me.mnuFormatFontSizeNo(i)
        Me.mnuFormatFontSizeNo(i).Caption = i
    Next
End Sub

Private Sub PrintLetter()
    On Error Resume Next
    Select Case m_iFormNo
    Case 5  'customers
        frmCustomer.PrintLetter
    Case Else
    End Select
End Sub


Public Sub ReadText()
Dim sHelp As String
    On Error Resume Next    'this is only text
    'find YOUR Language text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("mnuFiles")) Then
                    .Fields("mnuFiles") = mnuFiles.Caption
                Else
                    mnuFiles.Caption = .Fields("mnuFiles")
                End If
                If IsNull(.Fields("mnuUser")) Then
                    .Fields("mnuUser") = mnuUser.Caption
                Else
                    mnuUser.Caption = .Fields("mnuUser")
                End If
                If IsNull(.Fields("mnuExit")) Then
                    .Fields("mnuExit") = mnuExit.Caption
                Else
                    mnuExit.Caption = .Fields("mnuExit")
                End If
                If IsNull(.Fields("mnuMich")) Then
                    .Fields("mnuMich") = mnuMich.Caption
                Else
                    mnuMich.Caption = .Fields("mnuMich")
                End If
                If IsNull(.Fields("mnuKeyAscii")) Then
                    .Fields("mnuKeyAscii") = mnuKeyAscii.Caption
                Else
                    mnuKeyAscii.Caption = .Fields("mnuKeyAscii")
                End If
                If IsNull(.Fields("mnuPasword")) Then
                    .Fields("mnuPasword") = mnuPasword.Caption
                Else
                    mnuPasword.Caption = .Fields("mnuPasword")
                End If
                If IsNull(.Fields("mnuAbout")) Then
                    .Fields("mnuAbout") = mnuAbout.Caption
                Else
                    mnuAbout.Caption = .Fields("mnuAbout")
                End If
                If IsNull(.Fields("mnuPrint")) Then
                    .Fields("mnuPrint") = mnuPrint.Caption
                Else
                    mnuPrint.Caption = .Fields("mnuPrint")
                End If
                If IsNull(.Fields("mnuPrintSetUp")) Then
                    .Fields("mnuPrintSetUp") = mnuPrintSetUp.Caption
                Else
                    mnuPrintSetUp.Caption = .Fields("mnuPrintSetUp")
                End If
                If IsNull(.Fields("mnuPrintWord")) Then
                    .Fields("mnuPrintWord") = mnuPrintWord.Caption
                Else
                    mnuPrintWord.Caption = .Fields("mnuPrintWord")
                End If
                If IsNull(.Fields("mnuPrintDefault")) Then
                    .Fields("mnuPrintDefault") = mnuPrintDefault.Caption
                Else
                    mnuPrintDefault.Caption = .Fields("mnuPrintDefault")
                End If
                If IsNull(.Fields("mnuDatabaseInternet")) Then
                    .Fields("mnuDatabaseInternet") = mnuDatabaseInternet.Caption
                Else
                    mnuDatabaseInternet.Caption = .Fields("mnuDatabaseInternet")
                End If
                If IsNull(.Fields("mnuDatabaseLinkType")) Then
                    .Fields("mnuDatabaseLinkType") = mnuDatabaseLinkType.Caption
                Else
                    mnuDatabaseLinkType.Caption = .Fields("mnuDatabaseLinkType")
                End If
                If IsNull(.Fields("mnuDatabaseCodeLanguage")) Then
                    .Fields("mnuDatabaseCodeLanguage") = mnuDatabaseCodeLanguage.Caption
                Else
                    mnuDatabaseCodeLanguage.Caption = .Fields("mnuDatabaseCodeLanguage")
                End If
                If IsNull(.Fields("mnuDatabaseCodeType")) Then
                    .Fields("mnuDatabaseCodeType") = mnuDatabaseCodeType.Caption
                Else
                    mnuDatabaseCodeType.Caption = .Fields("mnuDatabaseCodeType")
                End If
                If IsNull(.Fields("mnuIconType")) Then
                    .Fields("mnuIconType") = mnuIconType.Caption
                Else
                    mnuIconType.Caption = .Fields("mnuIconType")
                End If
                If IsNull(.Fields("mnuAPIType")) Then
                    .Fields("mnuAPIType") = mnuAPIType.Caption
                Else
                    mnuAPIType.Caption = .Fields("mnuAPIType")
                End If
                If IsNull(.Fields("mnuColors")) Then
                    .Fields("mnuColors") = mnuColors.Caption
                Else
                    mnuColors.Caption = .Fields("mnuColors")
                End If
                If IsNull(.Fields("mnuFormatCopy")) Then
                    .Fields("mnuFormatCopy") = mnuFormatCopy.Caption
                Else
                    mnuFormatCopy.Caption = .Fields("mnuFormatCopy")
                End If
                If IsNull(.Fields("mnuFormatPaste")) Then
                    .Fields("mnuFormatPaste") = mnuFormatPaste.Caption
                Else
                    mnuFormatPaste.Caption = .Fields("mnuFormatPaste")
                End If
                If IsNull(.Fields("mnuFormatBold")) Then
                    .Fields("mnuFormatBold") = mnuFormatBold.Caption
                Else
                    mnuFormatBold.Caption = .Fields("mnuFormatBold")
                End If
                If IsNull(.Fields("mnuFormatItalic")) Then
                    .Fields("mnuFormatItalic") = mnuFormatItalic.Caption
                Else
                    mnuFormatItalic.Caption = .Fields("mnuFormatItalic")
                End If
                If IsNull(.Fields("mnuFormatUline")) Then
                    .Fields("mnuFormatUline") = mnuFormatUline.Caption
                Else
                    mnuFormatUline.Caption = .Fields("mnuFormatUline")
                End If
                If IsNull(.Fields("mnuFormatFont")) Then
                    .Fields("mnuFormatFont") = mnuFormatFont.Caption
                Else
                    mnuFormatFont.Caption = .Fields("mnuFormatFont")
                End If
                If IsNull(.Fields("mnuFormatFontSize")) Then
                    .Fields("mnuFormatFontSize") = mnuFormatFontSize.Caption
                Else
                    mnuFormatFontSize.Caption = .Fields("mnuFormatFontSize")
                End If
                If IsNull(.Fields("mnuFormatColor")) Then
                    .Fields("mnuFormatColor") = mnuFormatColor.Caption
                Else
                    mnuFormatColor.Caption = .Fields("mnuFormatColor")
                End If
                If IsNull(.Fields("mnuFormatPastePic")) Then
                    .Fields("mnuFormatPastePic") = mnuFormatPastePic.Caption
                Else
                    mnuFormatPastePic.Caption = .Fields("mnuFormatPastePic")
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
        .Fields("mnuFiles") = mnuFiles.Caption
        .Fields("mnuUser") = mnuUser.Caption
        .Fields("mnuExit") = mnuExit.Caption
        .Fields("mnuMich") = mnuMich.Caption
        .Fields("mnuKeyAscii") = mnuKeyAscii.Caption
        .Fields("mnuPasword") = mnuPasword.Caption
        .Fields("mnuAbout") = mnuAbout.Caption
        .Fields("mnuPrint") = mnuPrint.Caption
        .Fields("mnuDatabaseLinkType") = mnuDatabaseLinkType.Caption
        .Fields("mnuDatabaseCodeLanguage") = mnuDatabaseCodeLanguage.Caption
        .Fields("mnuDatabaseCodeType") = mnuDatabaseCodeType.Caption
        .Fields("mnuPrintSetUp") = mnuPrintSetUp.Caption
        .Fields("mnuPrintWord") = mnuPrintWord.Caption
        .Fields("mnuPrintDefault") = mnuPrintDefault.Caption
        .Fields("mnuDatabaseInternet") = mnuDatabaseInternet.Caption
        .Fields("mnuIconType") = mnuIconType.Caption
        .Fields("mnuAPIType") = mnuAPIType.Caption
        .Fields("mnuColors") = mnuColors.Caption
        .Fields("mnuFormatCopy") = mnuFormatCopy.Caption
        .Fields("mnuFormatPaste") = mnuFormatPaste.Caption
        .Fields("mnuFormatBold") = mnuFormatBold.Caption
        .Fields("mnuFormatItalic") = mnuFormatItalic.Caption
        .Fields("mnuFormatUline") = mnuFormatUline.Caption
        .Fields("mnuFormatFont") = mnuFormatFont.Caption
        .Fields("mnuFormatFontSize") = mnuFormatFontSize.Caption
        .Fields("mnuFormatColor") = mnuFormatColor.Caption
        .Fields("mnuFormatPastePic") = mnuFormatPastePic.Caption
        .Fields("nCode") = "Code"
        .Fields("nCodeSnip") = "Code Snippets"
        .Fields("nCodeZip") = "Code Zip-files"
        .Fields("nCodeTyp") = "Code Type"
        .Fields("nProgramming") = "Programming"
        .Fields("nProgramHours") = "Program Hours"
        .Fields("nProjects") = "Projects"
        .Fields("nLicense") = "Licence"
        .Fields("nMail") = "Mail"
        .Fields("nCustomers") = "Customers"
        .Fields("nInvoice") = "Invoice"
        .Fields("nPayments") = "Payments"
        .Fields("nDatabase") = "Database"
        .Fields("nPasswords") = "Passwords"
        .Fields("nDatabaseFields") = "Database Fields"
        .Fields("nRepairDatabase") = "Repair Database"
        .Fields("nUserRecord") = "User Record"
        .Fields("nScreenText") = "Screen Text"
        .Fields("nCountry") = "Country"
        .Fields("nAbout") = "About"
        .Fields("nAbout2") = "About..."
        .Fields("nWriteToMe") = "Write To Me"
        .Fields("nProgramSupplier") = "Program Supplier"
        .Fields("nRegisterProgram") = "Register Program"
        .Fields("nUpdateProgram") = "Update Program"
        .Fields("nVB") = "Visual Basic"
        .Fields("nVBCodeStatistic") = "Code Statistic"
        .Fields("nPic") = "Pictures"
        .Fields("nPicIcon") = "Viewer"
        .Fields("nIcon") = "Icons"
        .Fields("nSpell") = "Code Spell"
        .Fields("nSendMail") = "Send Mail"
        .Fields("Help") = sHelp
        .Update
        .Bookmark = .LastModified
    End With
End Sub

Private Sub CopyToClip()
    On Error Resume Next
    Select Case m_iFormNo
        Case 2  'code snippets
            frmCodeSnippets.CopySnippToClip
        Case 33 'code zil-files
            frmCodeZip.CopyZipToClip
        Case Else
    End Select
End Sub


Private Sub DeleteRecord()
    On Error Resume Next
    Select Case m_iFormNo
        Case 2  'code snippets
            frmCodeSnippets.DeleteRecord
        Case 4  'country
            frmCountry.DeleteRecord
        Case 5  'customer
            frmCustomer.DeleteRecord
        Case 8  'invoice
            frmInvoice.DeleteInvoice
        Case 10 'license
            frmLicence.DeleteRecord
        Case 14 'passwords
            frmPasswords.DeleteRecord
        Case 18 'programming
            If frmProgramming.Tab1.Tab = 0 Then
                frmProgramming.DeletePrograming
            Else
                frmProgramming.DeleteWish
            End If
        Case 33 'zip files
            frmCodeZip.DeleteRecord
        Case 34 'VB Internetlinks
            frmLinks.DeleteLink
        Case 35 'link types
            frmLinkType.DeleteLinkType
        Case 36 'icons
            frmIcons.DeleteIcon
        Case 37 'API
            frmAPI.DeleteRecord
        Case Else
    End Select
End Sub
Public Sub LoadMenu()
    With Menu1
        .MenusMax = 6
        .MenuCur = 1    'code
        .MenuItemsMax = 3
        .MenuCaption = rsLanguage.Fields("nCode")
        .MenuItemCur = 1    'code snippets
        Set .MenuItemIcon = LoadResPicture(101, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nCodeSnip")
        .MenuItemCur = 2    'code snippets
        Set .MenuItemIcon = LoadResPicture(125, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nCodeZip")
        .MenuItemCur = 3    'API code
        Set .MenuItemIcon = LoadResPicture(129, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nCodeAPI")
        
        .MenuCur = 2    'programming
        .MenuItemsMax = 7
        .MenuCaption = rsLanguage.Fields("nProgramming")
        .MenuItemCur = 1    'program hours
        Set .MenuItemIcon = LoadResPicture(103, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nProgramHours")
        .MenuItemCur = 2    'projects
        Set .MenuItemIcon = LoadResPicture(104, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nProjects")
        .MenuItemCur = 3    'License
        Set .MenuItemIcon = LoadResPicture(105, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nLicense")
        .MenuItemCur = 4    'mail
        Set .MenuItemIcon = LoadResPicture(106, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nMail")
        .MenuItemCur = 5    'customers
        Set .MenuItemIcon = LoadResPicture(107, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nCustomers")
        .MenuItemCur = 6    'invoice
        Set .MenuItemIcon = LoadResPicture(108, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nInvoice")
        .MenuItemCur = 7    'payments
        Set .MenuItemIcon = LoadResPicture(110, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nPayments")
        
        .MenuCur = 3    'database
        .MenuItemsMax = 7
        .MenuCaption = rsLanguage.Fields("nDatabase")
        .MenuItemCur = 1    'passwords
        Set .MenuItemIcon = LoadResPicture(113, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nPasswords")
        .MenuItemCur = 2    'database fields
        Set .MenuItemIcon = LoadResPicture(114, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nDatabaseFields")
        .MenuItemCur = 3    'repair database
        Set .MenuItemIcon = LoadResPicture(115, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nRepairDatabase")
        .MenuItemCur = 4    'repair database
        Set .MenuItemIcon = LoadResPicture(121, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nUserRecord")
        .MenuItemCur = 5    'invoice text
        Set .MenuItemIcon = LoadResPicture(109, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nScreenText")
        .MenuItemCur = 6    'country
        Set .MenuItemIcon = LoadResPicture(112, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nCountry")
        .MenuItemCur = 7    'send mail
        Set .MenuItemIcon = LoadResPicture(128, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nSendMail")
        
        .MenuCur = 4    'VisualBasic
        .MenuItemsMax = 2
        .MenuCaption = rsLanguage.Fields("nVB")
        .MenuItemCur = 1    'code statistic
        Set .MenuItemIcon = LoadResPicture(122, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nVBCodeStatistic")
        .MenuItemCur = 2    'code spell check
        Set .MenuItemIcon = LoadResPicture(127, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nSpell")
        
        .MenuCur = 5    'Pictures
        .MenuItemsMax = 2
        .MenuCaption = rsLanguage.Fields("nPic")
        .MenuItemCur = 1    'Pictures
        Set .MenuItemIcon = LoadResPicture(123, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nPicIcon")
        .MenuItemCur = 2    'icons
        Set .MenuItemIcon = LoadResPicture(124, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nIcon")
        
        .MenuCur = 6    'About
        .MenuItemsMax = 5
        .MenuCaption = rsLanguage.Fields("nAbout")
        .MenuItemCur = 1    'about
        Set .MenuItemIcon = LoadResPicture(117, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nAbout2")
        .MenuItemCur = 2    'write to me
        Set .MenuItemIcon = LoadResPicture(116, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nWriteToMe")
        .MenuItemCur = 3    'program supplier
        Set .MenuItemIcon = LoadResPicture(118, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nProgramSupplier")
        .MenuItemCur = 4    'program registration
        Set .MenuItemIcon = LoadResPicture(119, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nRegisterProgram")
        .MenuItemCur = 5    'program registration
        Set .MenuItemIcon = LoadResPicture(120, vbResIcon)
        .MenuItemCaption = rsLanguage.Fields("nUpdateProgram")
        
        .MenuCur = 1
    End With
End Sub


Private Sub CloseActiveForm()
Dim i As Integer
    On Error Resume Next
    For i = 0 To Forms.Count - 1
        ' Find first form besides "me" to unload
        If Forms(i).Name <> Me.Name Then
            Unload Forms(i)
        End If
    Next
End Sub
Private Sub MakeNewRecord()
    On Error Resume Next
    Select Case m_iFormNo
        Case 2  'code snippets
            frmCodeSnippets.NewRecord
        Case 4  'country
            frmCountry.NewRecord
        Case 5  'customer
            frmCustomer.NewRecord
        Case 8  'invoice
            frmInvoice.NewInvoice
        Case 10 'license
            frmLicence.NewRecord
        Case 14 'passwords
            frmPasswords.NewRecord
        Case 18 'programming
            If frmProgramming.Tab1.Tab = 0 Then
                frmProgramming.NewPrograming
            Else
                frmProgramming.NewWish
            End If
        Case 33 'zipped code
            frmCodeZip.NewRecord
        Case 34 'VB Internetlinks
            frmLinks.NewLink
        Case 35 'link types
            frmLinkType.NewLinkType
        Case 36 'icons
            frmIcons.NewIcon
        Case 37 'API
            frmAPI.NewRecord
        Case Else
    End Select
End Sub

Private Sub PrintRecord()
    On Error Resume Next
    Select Case m_iFormNo
        Case 2  'code snippets
            frmCodeSnippets.PrintRecord
        Case 5  'customers
            frmPrintCustomer.Show vbModal
        Case 8  'print invoice
            frmInvoice.PrintInvoice
        Case 10 'licence
            frmPrintLicence.Show vbModal
        Case 14 'passwords
            frmPasswords.WritePasswords
        Case 18 'programming
            If frmProgramming.Tab1.Tab = 0 Then
                frmWriteProgRep.Show vbModal
            Else
                frmProgramming.WriteWishList
            End If
        Case 37
            frmAPI.PrintRecord
        Case Else
    End Select
End Sub

Private Sub SaveRecords()
    On Error Resume Next
    Select Case m_iFormNo
        Case 8  'invoice
            frmInvoice.UpdateInvoice
        Case Else
    End Select
End Sub

Private Sub SearchForString()
    On Error Resume Next
    Select Case m_iFormNo
        Case 2  'code snippets
            frmSearchCode.Show 1
        Case Else
    End Select
End Sub

Private Sub SendMail()
Dim IsValid As Boolean
Dim InvalidReason As String
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        frmSnippetMail.Show 1
    Case 5  'customers
        IsValid = IsEMailAddress(Trim(frmCustomer.Text1(9).Text), InvalidReason)
        If Not IsValid Then
            MsgBox "Invalid mail address, the reason given is: " & InvalidReason
            Exit Sub
        End If
        With frmEmail
            .List1.AddItem Trim(frmCustomer.Text1(9).Text)
            .cboAdr.Visible = False
            .btnMailTo.Visible = False
            .Show vbModal
        End With
    Case 33 'zip codes
        If Len(frmCodeZip.Text1(3).Text) = 0 Then Exit Sub
        With frmEmail
            .List1.AddItem Trim(frmCodeZip.Text1(3).Text)
            .cboAdr.Visible = False
            .btnMailTo.Visible = False
            .Show vbModal
        End With
    Case 37 'api
        If Len(frmAPI.Text1(4).Text) = 0 Then Exit Sub
        With frmEmail
            .List1.AddItem Trim(frmAPI.Text1(4).Text)
            .cboAdr.Visible = False
            .btnMailTo.Visible = False
            .Show vbModal
        End With
    Case Else
    End Select
End Sub

Private Sub ShowHelp()
    Select Case m_iFormNo
        Case 0  'frmMDI
            frmHelp.Label1.Caption = "frmMDI"
            frmHelp.Show 1
        Case 2  'frmCodeSnippets
            frmHelp.Label1.Caption = "frmCodeSnippets"
            frmHelp.Show 1
        Case 3  'frmCodeType
            frmHelp.Label1.Caption = "frmCodeType"
            frmHelp.Show 1
        Case 4  'frmCountry
            frmHelp.Label1.Caption = "frmCountry"
            frmHelp.Show 1
        Case 4  'frmCountry
            frmHelp.Label1.Caption = "frmCountry"
            frmHelp.Show 1
        Case 8  'frmInvoice
            frmHelp.Label1.Caption = "frmInvoice"
            frmHelp.Show 1
        Case 10  'frmLicence
            frmHelp.Label1.Caption = "frmLicence"
            frmHelp.Show 1
        Case 12  'frmMassMail
            frmHelp.Label1.Caption = "frmMassMail"
            frmHelp.Show 1
        Case 14  'frmPasswords
            frmHelp.Label1.Caption = "frmPasswords"
            frmHelp.Show 1
        Case 16  'frmPayments
            frmHelp.Label1.Caption = "frmPayments"
            frmHelp.Show 1
        Case 17  'frmPrintDB
            frmHelp.Label1.Caption = "frmPrintDB"
            frmHelp.Show 1
        Case 18  'frmProgramming
            frmHelp.Label1.Caption = "frmProgramming"
            frmHelp.Show 1
        Case 19  'frmProjects
            frmHelp.Label1.Caption = "frmProjects"
            frmHelp.Show 1
        Case 20 'frmScreenLanguage
            frmHelp.Label1.Caption = "frmScreenLanguage"
            frmHelp.Show 1
        Case 22  'frmUser
            frmHelp.Label1.Caption = "frmUser"
            frmHelp.Show 1
        Case 30  'frmMaint
            frmHelp.Label1.Caption = "frmMaint"
            frmHelp.Show 1
        Case 31  'frmStats
            frmHelp.Label1.Caption = "frmStats"
            frmHelp.Show 1
        Case 33 'code zip
            frmHelp.Label1.Caption = "frmCodeZip"
            frmHelp.Show 1
        Case 34 'VB Internet links
            frmHelp.Label1.Caption = "frmLinks"
            frmHelp.Show 1
        Case 35 'colors
        Case 36 'icons
            frmHelp.Label1.Caption = "frmIcons"
            frmHelp.Show 1
        Case 37 'API
            frmHelp.Label1.Caption = "frmAPI"
            frmHelp.Show 1
        Case Else
    End Select
End Sub


Private Sub ShowInternet()
Dim iRet As Long
    On Error Resume Next
    Select Case m_iFormNo
    Case 2
        If Len(frmCodeSnippets.Text5.Text) = 0 Then Exit Sub
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            "http://" & Trim(frmCodeSnippets.Text5.Text), vbNullString, "c:\", _
            SW_SHOWNORMAL)
    Case 5  'customers
        If Len(frmCustomer.Text1(10).Text) = 0 Then Exit Sub
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            "http://" & Trim(frmCustomer.Text1(10).Text), vbNullString, "c:\", _
            SW_SHOWNORMAL)
    Case 33 'zip codes
        If Len(frmCodeZip.Text1(4).Text) = 0 Then Exit Sub
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            Trim(frmCodeZip.Text1(4).Text), vbNullString, "c:\", SW_SHOWNORMAL)
    Case 34 'VB Internetlinks
        If Len(frmLinks.Text1(1).Text) = 0 Then Exit Sub
        frmLinks.UpdateLinkDate
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            Trim(frmLinks.Text1(1).Text), vbNullString, "c:\", SW_SHOWNORMAL)
    Case 37 'api
        If Len(frmAPI.Text1(5).Text) = 0 Then Exit Sub
        iRet = ShellExceCute(Me.hWnd, _
            vbNullString, _
            Trim(frmAPI.Text1(5).Text), vbNullString, "c:\", SW_SHOWNORMAL)
    End Select
End Sub

Private Sub MDIForm_Activate()
    On Error Resume Next
    rsUser.Refresh
    If boolFirst = False Then Exit Sub
    If CBool(rsUser.Recordset.Fields("PrintUseWord")) Then
        mnuPrintWord.Checked = True
        mnuPrintDefault.Checked = False
    Else
        mnuPrintWord.Checked = False
        mnuPrintDefault.Checked = True
    End If
    m_FileExt = rsUser.Recordset.Fields("LanguageOnScreen")
    ReadText
    LoadMenu
    LoadColor
    Me.StatusBar1.Panels(1).Text = App.Path & "\Programming.exe"
    Me.Caption = Me.Caption & " -  Version: " & App.Major & "." & App.Minor & "." & App.Revision
    DisableButtons 1
    IsCodeSnippet  'do we have the user code snippet database ?
    boolFirst = False
End Sub
Private Sub MDIForm_Load()
Dim lngOldId As Long, retval As Long

    On Error GoTo errForm_Load
    m_strPrograming = App.Path & "\CodeMaster.mdb"
    Set m_dbPrograming = OpenDatabase(m_strPrograming)
    rsUser.DatabaseName = m_strPrograming
    
    m_strProgramLng = App.Path & "\CodeLang.mdb"
    Set m_dbLanguage = OpenDatabase(m_strProgramLng)
    
    m_strCodeSnippet = App.Path & "\CodeSnippets.mdb"
    Set m_dbCodeSnippet = OpenDatabase(m_strCodeSnippet)
    
    m_strCodeZip = App.Path & "\CodeZip.mdb"
    Set m_dbCodeZip = OpenDatabase(m_strCodeZip)
    
    Set rsColor = m_dbPrograming.OpenRecordset("Color")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmMDI")
    
    hSysMenu = GetSystemMenu(Me.hWnd, False)
    With zMENU
        .cbSize = Len(zMENU)
        .dwTypeData = String(80, 0)
        .cch = Len(.dwTypeData)
        .fMask = MENU_STATE
        .wid = SC_CLOSE
    End With
    
    retval = GetMenuItemInfo(hSysMenu, zMENU.wid, False, zMENU)
    
    With zMENU
        lngOldId = .wid         'You need the old wID.
        .wid = xSC_CLOSE        'Change the wID to "no close"
        .fState = MFS_GRAYED    'Make the close methods gray
        .fMask = MENU_ID        'Specifys that the value in wID is a id and not a state
    End With
    
    retval = SetMenuItemInfo(hSysMenu, lngOldId, False, zMENU)
    zMENU.fMask = MENU_STATE
    retval = SetMenuItemInfo(hSysMenu, zMENU.wid, False, zMENU)
    m_iFormNo = 0
    boolFirst = True
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    rsUser.Recordset.Close
    rsLanguage.Close
    rsColor.Close
    m_dbPrograming.Close
    m_dbMyCodeSnippet.Close
    Set frmMDI = Nothing
End Sub

Private Sub Menu1_MenuItemClick(MenuNumber As Long, MenuItem As Long)
    CloseActiveForm
    Select Case MenuNumber
    Case 1  'files
        Select Case MenuItem
            Case 1  'code snippets
                frmCodeSnippets.Show
            Case 2  'zipped code
                frmCodeZip.Show
            Case 3  'api
                frmAPI.Show
            Case Else
        End Select
    Case 2  'programing
        Select Case MenuItem
            Case 1
                frmProgramming.Show
            Case 2
                frmProjects.Show
            Case 3
                frmLicence.Show
            Case 4
                frmMassMail.Show
            Case 5
                frmCustomer.Show
            Case 6
                frmInvoice.Show
            Case 7
                frmPayments.Show
            Case Else
        End Select
    Case 3  'database
        Select Case MenuItem
            Case 1
                frmPasswords.Show
            Case 2
                frmPrintDB.Show
            Case 3
                frmMaint.Show
            Case 4  'user's record
                frmUser.Show
            Case 5  'screen text
                frmScreenLanguage.Show
            Case 6
                frmCountry.Show
            Case 7  'send mail
                With frmEmail
                    .btnMailTo.Visible = True
                    .List1.Visible = False
                    .Text2.Visible = True
                    .Label1(1).Visible = False
                    .Show
                End With
            Case Else
        End Select
    Case 4  'VB
        Select Case MenuItem
            Case 1
                frmStats.Show
            Case 2
                frmSpellChecker.Show
            Case Else
        End Select
    Case 5  'pictures
        Select Case MenuItem
            Case 1
                frmPicViewer.Show
            Case 2  'icons
                frmIcons.Show
            Case Else
        End Select
    Case 6  'about
        Select Case MenuItem
            Case 1
                frmAbout.Show vbModal
            Case 2
                frmWriteToMe.Show 1
            Case 3
                frmSupplier.Show 1
            Case 4
                frmRegistration.Show 1
            Case 5  'live update
                MsgBox "Not implemented yet - sorry !"
                'frmLiveUpdate.Show 1
            Case Else
        End Select
    Case Else
    End Select
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuAPIType_Click()
    CloseActiveForm
    frmAPIType.Show vbModal
End Sub

Private Sub mnuColors_Click()
    CloseActiveForm
    frmColor.Show
End Sub

Private Sub mnuDatabaseCodeLanguage_Click()
    CloseActiveForm
    frmCodeLanguage.Show vbModal
End Sub

Private Sub mnuDatabaseCodeType_Click()
    CloseActiveForm
    frmCodeType.Show vbModal
End Sub

Private Sub mnuDatabaseInternet_Click()
    CloseActiveForm
    frmLinks.Show
End Sub

Private Sub mnuDatabaseLinkType_Click()
    CloseActiveForm
    frmLinkType.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFormatBold_Click()
    Select Case m_iFormNo
    Case 2  'code snippets
        If frmCodeSnippets.Text1.SelBold = False Then
             frmCodeSnippets.Text1.SelBold = True
        Else
             frmCodeSnippets.Text1.SelBold = False
        End If
    Case 33 'code zip
        If frmCodeZip.Text2.SelBold = False Then
             frmCodeZip.Text2.SelBold = True
        Else
             frmCodeZip.Text2.SelBold = False
        End If
    Case 37 'api
        If frmAPI.RichText1.SelBold = False Then
            frmAPI.RichText1.SelBold = True
        Else
            frmAPI.RichText1.SelBold = False
        End If
    Case Else
    End Select
End Sub

Private Sub mnuFormatColorNo_Click(Index As Integer)
Dim lRed As Long, lGreen As Long, lBlue As Long
    On Error Resume Next
    With rsColor
        .Bookmark = bookColor(Index)
        lRed = CLng(.Fields("RedValue"))
        lGreen = CLng(.Fields("GreenValue"))
        lBlue = CLng(.Fields("BlueValue"))
    End With
    Select Case m_iFormNo
    Case 2  'code snippets
        frmCodeSnippets.Text1.SelColor = RGB(lRed, lGreen, lBlue)
    Case 33 'code zip
        frmCodeZip.Text2.SelColor = RGB(lRed, lGreen, lBlue)
    Case 37 'api
        frmAPI.RichText1.SelColor = RGB(lRed, lGreen, lBlue)
    Case Else
    End Select
End Sub

Private Sub mnuFormatCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Select Case m_iFormNo
    Case 2  'code snippets
        Clipboard.SetText frmCodeSnippets.Text1.SelText
    Case 33 'code zip
        Clipboard.SetText frmCodeZip.Text2.SelText
    Case 37 'api
        Clipboard.SetText frmAPI.RichText1.SelText
    Case Else
    End Select
End Sub

Private Sub mnuFormatFontName_Click(Index As Integer)
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        frmCodeSnippets.Text1.SelFontName = mnuFormatFontName(Index).Caption
    Case 33 'code zip
        frmCodeZip.Text2.SelFontName = mnuFormatFontName(Index).Caption
    Case 37 'api
        frmAPI.RichText1.SelFontName = mnuFormatFontName(Index).Caption
    Case Else
    End Select
End Sub

Private Sub mnuFormatFontSizeNo_Click(Index As Integer)
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        frmCodeSnippets.Text1.SelFontSize = CInt(mnuFormatFontSizeNo(Index).Caption)
    Case 33 'code zip-files
        frmCodeZip.Text2.SelFontSize = CInt(mnuFormatFontSizeNo(Index).Caption)
    Case 37 'api
        frmAPI.RichText1.SelFontSize = CInt(mnuFormatFontSizeNo(Index).Caption)
    Case Else
    End Select
End Sub

Private Sub mnuFormatItalic_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        If frmCodeSnippets.Text1.SelItalic = False Then
            frmCodeSnippets.Text1.SelItalic = True
        Else
            frmCodeSnippets.Text1.SelItalic = False
        End If
    Case 33 'code zip
        If frmCodeZip.Text2.SelItalic = False Then
            frmCodeZip.Text2.SelItalic = True
        Else
            frmCodeZip.Text2.SelItalic = False
        End If
    Case 37 'api
        If frmAPI.RichText1.SelItalic = False Then
            frmAPI.RichText1.SelItalic = True
        Else
            frmAPI.RichText1.SelItalic = False
        End If
    Case Else
    End Select
End Sub

Private Sub mnuFormatPaste_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        frmCodeSnippets.Text1.SelText = Clipboard.GetText
    Case 33 'code zip
        frmCodeZip.Text2.SelText = Clipboard.GetText
    Case 37 'api
        frmAPI.RichText1.SelText = Clipboard.GetText
    Case Else
    End Select
End Sub

Private Sub mnuFormatPastePic_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        SendMessage frmCodeSnippets.Text1.hWnd, WM_PASTE, 0, 0
    Case 33 'code zip
        SendMessage frmCodeZip.Text2.hWnd, WM_PASTE, 0, 0
    Case 37 'api
        SendMessage frmAPI.RichText1.hWnd, WM_PASTE, 0, 0
    Case Else
    End Select
End Sub

Private Sub mnuFormatUline_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippets
        If frmCodeSnippets.Text1.SelUnderline = False Then
             frmCodeSnippets.Text1.SelUnderline = True
        Else
             frmCodeSnippets.Text1.SelUnderline = False
        End If
    Case 33 'code zip
        If frmCodeZip.Text2.SelUnderline = False Then
             frmCodeZip.Text2.SelUnderline = True
        Else
             frmCodeZip.Text2.SelUnderline = False
        End If
    Case 37 'api
        If frmAPI.RichText1.SelUnderline = False Then
             frmAPI.RichText1.SelUnderline = True
        Else
             frmAPI.RichText1.SelUnderline = False
        End If
    Case Else
    End Select
End Sub

Private Sub mnuIconType_Click()
    frmIconTypes.Show
End Sub

Private Sub mnuKeyAscii_Click()
    CloseActiveForm
    frmKeyAscii.Show 1
End Sub

Private Sub mnuPasword_Click()
   CloseActiveForm
   frmPasword.Show 1
End Sub

Private Sub mnuPrintDefault_Click()
    mnuPrintWord.Checked = False
    mnuPrintDefault.Checked = True
    With rsUser.Recordset
        .Edit
        .Fields("PrintUseWord") = False
        .Update
    End With
End Sub

Private Sub mnuPrintSetUp_Click()
    With CMD1
        .DialogTitle = "Printer Set-Up"
        .PrinterDefault = True
        .flags = PD_PRINTSETUP
        .Action = 5
    End With
End Sub

Private Sub mnuPrintWord_Click()
    mnuPrintWord.Checked = True
    mnuPrintDefault.Checked = False
    With rsUser.Recordset
        .Edit
        .Fields("PrintUseWord") = True
        .Update
    End With
End Sub

Private Sub mnuUser_Click()
    CloseActiveForm
    frmUser.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Exit"
            Unload Me
        Case "New"
            MakeNewRecord
        Case "Save"
            SaveRecords
        Case "Delete"
            DeleteRecord
        Case "Copy"
            CopyToClip
        Case "Search"
            SearchForString
        Case "Print"
            PrintRecord
        Case "Write"
            PrintLetter
        Case "Fax"
            frmFax.Show vbModal
        Case "Mail"
            SendMail
        Case "Web"
            ShowInternet
        Case "Open"
            frmSpellChecker.OpenProject
        Case "Spell"
            frmSpellChecker.ControlSpell
        Case "View"
            frmSpellChecker.ShowReport
        Case "Help"
            ShowHelp
        Case Else
    End Select
End Sub
