VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrintDB 
   BackColor       =   &H00404040&
   ClientHeight    =   8145
   ClientLeft      =   510
   ClientTop       =   1635
   ClientWidth     =   12735
   ControlBox      =   0   'False
   Icon            =   "frmPrintDB.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   12735
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "Show/ Print Fields"
      TabPicture(0)   =   "frmPrintDB.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "New Recordset / Fields"
      TabPicture(1)   =   "frmPrintDB.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame4 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   6255
         Begin VB.CommandButton btnEndRecordset 
            Caption         =   "Finish Recordset"
            Height          =   615
            Left            =   120
            TabIndex        =   46
            Top             =   6120
            Width           =   1455
         End
         Begin VB.CommandButton btnWriteChangesToFile 
            Caption         =   "Save Changes"
            Height          =   615
            Left            =   1680
            Picture         =   "frmPrintDB.frx":047A
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   6120
            Width           =   1455
         End
         Begin VB.CommandButton btnNewRecordset 
            Caption         =   "New Recordset"
            Height          =   615
            Left            =   3240
            Picture         =   "frmPrintDB.frx":0B3C
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   6120
            Width           =   1455
         End
         Begin VB.CommandButton btnNewField 
            Caption         =   "New Field"
            Height          =   615
            Left            =   4800
            Picture         =   "frmPrintDB.frx":11FE
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   6120
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   31
            Top             =   3960
            Width           =   2655
         End
         Begin VB.ComboBox cmbFieldType 
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3480
            TabIndex        =   30
            Top             =   4320
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   1
            Left            =   3480
            TabIndex        =   29
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Index           =   2
            Left            =   3480
            TabIndex        =   28
            Top             =   5280
            Width           =   1215
         End
         Begin VB.Frame Frame5 
            Caption         =   "Fields added"
            Height          =   3615
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   6015
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   2760
               Index           =   0
               Left            =   240
               TabIndex        =   23
               Top             =   600
               Width           =   735
            End
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   2760
               Index           =   1
               Left            =   960
               TabIndex        =   22
               Top             =   600
               Width           =   3135
            End
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   2760
               Index           =   2
               Left            =   4080
               TabIndex        =   21
               Top             =   600
               Width           =   1095
            End
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               Height          =   2760
               Index           =   3
               Left            =   5160
               TabIndex        =   20
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Field Name"
               Height          =   255
               Index           =   5
               Left            =   960
               TabIndex        =   27
               Top             =   360
               Width           =   3135
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Type"
               Height          =   255
               Index           =   6
               Left            =   4080
               TabIndex        =   26
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Size"
               Height          =   255
               Index           =   7
               Left            =   5160
               TabIndex        =   25
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Pos."
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   24
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Field Name:"
            Height          =   255
            Index           =   8
            Left            =   1440
            TabIndex        =   35
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Field Type:"
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   34
            Top             =   4320
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Field Size:"
            Height          =   255
            Index           =   10
            Left            =   1440
            TabIndex        =   33
            Top             =   4800
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Number oif decimals:"
            Height          =   255
            Index           =   11
            Left            =   1440
            TabIndex        =   32
            Top             =   5280
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000004&
         ForeColor       =   &H00000000&
         Height          =   6735
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   6135
         Begin VB.CommandButton btnCopyText 
            Caption         =   "Copy Field text"
            Height          =   615
            Left            =   600
            Picture         =   "frmPrintDB.frx":18C0
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   6000
            Width           =   2055
         End
         Begin VB.CommandButton btnPrint 
            Caption         =   "Print Recordset"
            Height          =   615
            Left            =   3360
            Picture         =   "frmPrintDB.frx":1F82
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   6000
            Width           =   1695
         End
         Begin VB.ListBox lstFields 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5295
            Left            =   1200
            TabIndex        =   13
            ToolTipText     =   "Click to select a Field"
            Top             =   600
            Width           =   1935
         End
         Begin VB.ListBox typeFields 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5295
            Left            =   3120
            TabIndex        =   12
            ToolTipText     =   "Click to select a Field"
            Top             =   600
            Width           =   1095
         End
         Begin VB.ListBox sizeFields 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5295
            Left            =   4200
            TabIndex        =   11
            ToolTipText     =   "Click to select a Field"
            Top             =   600
            Width           =   855
         End
         Begin VB.ListBox posFields 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            Height          =   5295
            Left            =   600
            TabIndex        =   10
            ToolTipText     =   "Click to select a Field"
            Top             =   600
            Width           =   615
         End
         Begin VB.Timer Timer1 
            Interval        =   50
            Left            =   0
            Top             =   4800
         End
         Begin MSComDlg.CommonDialog CMD1 
            Left            =   360
            Top             =   4920
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Field Name"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   17
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Size"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   15
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Pos."
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Height          =   975
      Left            =   9360
      TabIndex        =   5
      Top             =   6960
      Width           =   3255
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "HTML print"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "Printer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         Caption         =   "MS Word"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Height          =   7815
      Left            =   6720
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton btnOpen 
         Caption         =   "Open Recordset"
         Height          =   735
         Left            =   120
         Picture         =   "frmPrintDB.frx":20CC
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CommandButton btnPrintAll 
         Caption         =   "Print All Recordsets"
         Height          =   735
         Left            =   1320
         Picture         =   "frmPrintDB.frx":278E
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6960
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   6660
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Click to select a recordset"
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.DriveListBox DrvList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
   Begin VB.DirListBox DirList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   3240
      Left            =   9360
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.FileListBox FilList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   2955
      Left            =   9360
      Pattern         =   "*.mdb"
      TabIndex        =   0
      ToolTipText     =   "Click to select a database"
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label Label2 
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
      Left            =   1920
      TabIndex        =   37
      Top             =   120
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Database Files in:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmPrintDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolNewField As Boolean, fileLog As String, strFieldType As String
Dim FileNo As Integer
Dim boolNewRecordset As Boolean
Dim boolRecordWrite As Boolean
Dim sDatabaseName As Variant
Dim dbOpenDatabase As Database, rsOpenRecordset As Recordset
Dim rsLanguage As Recordset
Dim rsUser As Recordset
Dim wrkJet As Workspace
Dim wdApp As Word.Application
Dim iNumberOfFields As Integer, iIndex As Integer
Dim sTmp As String, i As Integer
Dim sTemp1 As String, sTemp2 As String
Dim tbl As TableDef
Dim tdfNew As TableDef
Dim iDirIndex As Integer
Dim iFilIndex As Integer
Dim iListIndex As Integer
Dim sField As Field
Dim FieldType As String
Dim LeftMargin As Integer
Dim cdlPDPrintSetup As Integer
Private Sub PrintHeadPreview()
    cPrint.FontBold = True
    cPrint.FontSize = "14"
    cPrint.pBox , , , 1.9, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint
    cPrint.pPrint "Database:", 0.5, True
    cPrint.FontBold = False
    cPrint.pPrint dbOpenDatabase.Name, 3, False
    cPrint.FontBold = True
    cPrint.pPrint "Table Definition for:", 0.5, True
    cPrint.FontBold = False
    cPrint.pPrint List1.List(List1.ListIndex), 3, False
    cPrint.pPrint
    cPrint.pDoubleLine
    cPrint.pPrint
    cPrint.FontSize = "10"
    cPrint.FontBold = True
    cPrint.pPrint "Pos nr", 0.5, True
    cPrint.pPrint "Description", 1, True
    cPrint.pPrint "Type", 3.5, True
    cPrint.pPrint "Size", 4.5
    cPrint.FontBold = False
    cPrint.pLine
    cPrint.BackColor = vbWhite
    cPrint.pPrint
    cPrint.pPrint
End Sub


Private Sub PrintWithPreviewAll()
Dim boolFirst As Boolean
    On Error GoTo errbPrintAllPreview
    Set cPrint = New clsMultiPgPreview
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    boolFirst = True
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    For n = 0 To List1.ListCount - 1
        List1.ListIndex = n
        If boolFirst Then
            Call PrintHeadPreview
            boolFirst = False
        Else
            cPrint.pFooter
            cPrint.pNewPage
            Call PrintHeadPreview
        End If
        DoEvents
        
        Set rsOpenRecordset = dbOpenDatabase.OpenRecordset(List1.List(List1.ListIndex))
        
        For i = 0 To rsOpenRecordset.Fields.Count - 1
            If cPrint.pEndOfPage Then
                cPrint.pFooter
                cPrint.pNewPage
                PrintHeadPreview
            End If
            Select Case rsOpenRecordset.Fields(i).Type
            Case dbBoolean
                FieldType = "Boolsk"
            Case dbByte
                FieldType = "Byte"
            Case dbInteger
                FieldType = "Integer"
            Case dbLong
                FieldType = "Long"
            Case dbCurrency
                FieldType = "Currency"
            Case dbSingle
                FieldType = "Single"
            Case dbDouble
                FieldType = "Double"
            Case dbDate
                FieldType = "Date"
            Case dbText
                FieldType = "Text"
            Case dbLongBinary
                FieldType = "LongBinary"
            Case dbMemo
                FieldType = "Memo"
            Case dbGUID
                FieldType = "GUID"
            End Select
            
            cPrint.pPrint i, 0.5, True
            cPrint.pPrint rsOpenRecordset.Fields(i).Name, 1, True
            cPrint.pPrint FieldType, 3.5, True
            cPrint.pPrint rsOpenRecordset.Fields(i).Size, 4.5, False
            DoEvents
        Next
    Next
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Exit Sub
    
errbPrintAllPreview:
    Beep
    MsgBox Err.Description, vbExclamation, "Print with preview"
    Err.Clear
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
                If IsNull(.Fields("label1")) Then
                    .Fields("label1") = Label1.Caption
                Else
                    Label1.Caption = .Fields("label1")
                End If
                For i = 0 To 11
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(i + 2) = Label7(i).Caption
                    Else
                        Label7(i).Caption = .Fields(i + 2)
                    End If
                Next
                SSTab1.Tab = 0
                If IsNull(.Fields("SSTab11")) Then
                    .Fields("SSTab11") = SSTab1.Caption
                Else
                    SSTab1.Caption = .Fields("SSTab11")
                End If
                SSTab1.Tab = 1
                If IsNull(.Fields("SSTab12")) Then
                    .Fields("SSTab12") = SSTab1.Caption
                Else
                    SSTab1.Caption = .Fields("SSTab12")
                End If
                If IsNull(.Fields("btnPrintAll")) Then
                    .Fields("btnPrintAll") = btnPrintAll.Caption
                Else
                    btnPrintAll.Caption = .Fields("btnPrintAll")
                End If
                If IsNull(.Fields("btnOpen")) Then
                    .Fields("btnOpen") = btnOpen.Caption
                Else
                    btnOpen.Caption = .Fields("btnOpen")
                End If
                If IsNull(.Fields("btnPrint")) Then
                    .Fields("btnPrint") = btnPrint.Caption
                Else
                    btnPrint.Caption = .Fields("btnPrint")
                End If
                If IsNull(.Fields("btnCopyText")) Then
                    .Fields("btnCopyText") = btnCopyText.Caption
                Else
                    btnCopyText.Caption = .Fields("btnCopyText")
                End If
                If IsNull(.Fields("btnNewField")) Then
                    .Fields("btnNewField") = btnNewField.Caption
                Else
                    btnNewField.Caption = .Fields("btnNewField")
                End If
                If IsNull(.Fields("btnNewRecordset")) Then
                    .Fields("btnNewRecordset") = btnNewRecordset.Caption
                Else
                    btnNewRecordset.Caption = .Fields("btnNewRecordset")
                End If
                If IsNull(.Fields("btnWriteChangesToFile")) Then
                    .Fields("btnWriteChangesToFile") = btnWriteChangesToFile.Caption
                Else
                    btnWriteChangesToFile.Caption = .Fields("btnWriteChangesToFile")
                End If
                If IsNull(.Fields("btnEndRecordset")) Then
                    .Fields("btnEndRecordset") = btnEndRecordset.Caption
                Else
                    btnEndRecordset.Caption = .Fields("btnEndRecordset")
                End If
                If IsNull(.Fields("Check1")) Then
                    .Fields("Check1") = Check1.Caption
                Else
                    Check1.Caption = .Fields("Check1")
                End If
                If IsNull(.Fields("Check2")) Then
                    .Fields("Check2") = Check2.Caption
                Else
                    Check2.Caption = .Fields("Check2")
                End If
                If IsNull(.Fields("Check3")) Then
                    .Fields("Check3") = Check3.Caption
                Else
                    Check3.Caption = .Fields("Check3")
                End If
                SSTab1.Tab = 0
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
        .Fields("label1") = Label1.Caption
        For i = 0 To 11
            .Fields(i + 2) = Label7(i).Caption
        Next
        SSTab1.Tab = 0
        .Fields("SSTab11") = SSTab1.Caption
        SSTab1.Tab = 1
        .Fields("SSTab12") = SSTab1.Caption
        .Fields("btnPrintAll") = btnPrintAll.Caption
        .Fields("btnOpen") = btnOpen.Caption
        .Fields("btnPrint") = btnPrint.Caption
        .Fields("btnCopyText") = btnCopyText.Caption
        .Fields("btnNewField") = btnNewField.Caption
        .Fields("btnNewRecordset") = btnNewRecordset.Caption
        .Fields("btnWriteChangesToFile") = btnWriteChangesToFile.Caption
        .Fields("btnEndRecordset") = btnEndRecordset.Caption
        .Fields("Check1") = Check1.Caption
        .Fields("Check2") = Check2.Caption
        .Fields("Check3") = Check3.Caption
        .Fields("Help") = sHelp
        .Update
        SSTab1.Tab = 0
    End With
End Sub

Private Sub FindDatabaseName()
Dim strName As String, initLength As Long
Dim n As Long, strTmp As String

    strName = sTmp
    initLength = Len(strName)
    n = initLength
    
startSeek:  'find the first occurance of the backslash (\)
    strTmp = Left(strName, n)
    If Right(strTmp, 1) = "\" Then
        sTmp = Right(sTmp, (initLength - n))
        Exit Sub
    ElseIf n = 0 Then   'we did not find any backslash (just in case)
        Exit Sub
    Else
        n = n - 1
        GoTo startSeek
    End If
End Sub

Private Sub AppendField()
    With tdfNew
        Select Case strFieldType
        Case "Boolean"
            .Fields.Append .CreateField(sTemp1, dbBoolean)
        Case "Byte"
            .Fields.Append .CreateField(sTemp1, dbByte, sTemp2)
        Case "Integer"
            .Fields.Append .CreateField(sTemp1, dbInteger, sTemp2)
        Case "Long"
            .Fields.Append .CreateField(sTemp1, dbLong, sTemp2)
        Case "Currency"
            .Fields.Append .CreateField(sTemp1, dbCurrency, sTemp2)
        Case "Single"
            .Fields.Append .CreateField(sTemp1, dbSingle, sTemp2)
        Case "Double"
            .Fields.Append .CreateField(sTemp1, dbDouble, sTemp2)
        Case "Date"
            .Fields.Append .CreateField(sTemp1, dbDate)
        Case "Text"
            .Fields.Append .CreateField(sTemp1, dbText, sTemp2)
        Case "LongBinary"
            .Fields.Append .CreateField(sTemp1, dbLongBinary)
        Case "Memo"
            .Fields.Append .CreateField(sTemp1, dbMemo)
        Case "GUID"
            .Fields.Append .CreateField(sTemp1, dbGUID)
        Case Else
        End Select
    End With
End Sub

Sub GetTableList()
  On Error GoTo errGetTableList
  List1.Clear
  posFields.Clear
  lstFields.Clear
  typeFields.Clear
  sizeFields.Clear
  
  'add the tabledefs
  For Each tbl In dbOpenDatabase.TableDefs
    sTmp = tbl.Name
    If (dbOpenDatabase.TableDefs(sTmp).Attributes And dbSystemObject) = 0 Then
        List1.AddItem sTmp
        List1.ItemData(List1.NewIndex) = 0
    End If
  Next
  Exit Sub
  
errGetTableList:
  MsgBox Err.Description, vbExclamation
  Err.Clear
End Sub
Sub OpenLocalDB()
    On Error GoTo OpenError
    ' Create Microsoft Jet Workspace object.
    Set wrkJet = CreateWorkspace("", "admin", "")
    sDatabaseName = DirList.List(iDirIndex) & "\" & FilList.List(iFilIndex)
    Label2.Caption = sDatabaseName
    Set dbOpenDatabase = wrkJet.OpenDatabase(sDatabaseName, False, True, vbNullString)
    Exit Sub
    
OpenError:
    Beep
    MsgBox Err.Description, vbCritical, "Open database"
    Resume OpenError2
OpenError2:
End Sub

Private Sub PrintWithPreview()
    
    Set cPrint = New clsMultiPgPreview
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If

    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    
    PrintHeadPreview
    
    For i = 0 To posFields.ListCount - 1
        If cPrint.pEndOfPage Then
            cPrint.pFooter
            cPrint.pNewPage
            PrintHeadPreview
        End If
        cPrint.pPrint posFields.List(i), 0.5, True
        cPrint.pPrint lstFields.List(i), 1, True
        cPrint.pPrint typeFields.List(i), 3.5, True
        cPrint.pPrint sizeFields.List(i), 4.5, False
    Next
    
    Screen.MousePointer = vbDefault
    cPrint.pFooter
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
End Sub

Private Sub btnCopyText_Click()
Dim sText As String
    On Error Resume Next
    If CBool(rsUser.Fields("CopyWithAppro")) Then
        sText = """" & CStr(lstFields.List(iIndex)) & """"
    Else
        sText = CStr(lstFields.List(iIndex))
    End If
    Clipboard.Clear
    Clipboard.SetText sText
End Sub

Private Sub btnEndRecordset_Click()
    On Error GoTo errbtnEndRecordset_Click
    dbOpenDatabase.TableDefs.Append tdfNew
    boolRecordWrite = False
    btnEndRecordset.Enabled = False
    Exit Sub
    
errbtnEndRecordset_Click:
    Beep
    MsgBox Err.Description, vbCritical, "End Recordset"
    Err.Clear
End Sub
Private Sub btnNewField_Click()
    On Error GoTo errbtnNewField_Click
    If Not boolNewField Then    'the first field
        boolNewField = True
        FileNo = FileNo + 1
        fileLog = "DatabaseUpdate" & FileNo & ".txt"
        sTmp = dbOpenDatabase.Name
        sTemp1 = rsOpenRecordset.Name
        Set tdfNew = dbOpenDatabase.TableDefs(sTemp1)

        'Call FindDatabaseName
        
        Open fileLog For Output As #1
        Write #1, sTmp, sTemp1, "Append"
        
        List2(0).Clear
        List2(1).Clear
        List2(2).Clear
        List2(3).Clear
        iNumberOfFields = iNumberOfFields - 1
    End If
    
    Text1(0).Text = " "
    Text1(1).Text = 0
    Text1(2).Text = 0
    cmbFieldType.Text = " "
    btnWriteChangesToFile.Enabled = True
    btnNewField.Enabled = False
    btnNewRecordset.Enabled = False
    Text1(0).SetFocus
    Exit Sub
    
errbtnNewField_Click:
    Beep
    MsgBox Err.Description, vbCritical, "New Field"
    Err.Clear
End Sub

Private Sub btnNewRecordset_Click()
    On Error GoTo errbtnNewRecordset_Click
    Text1(0).Text = " "
    Label7(8).Caption = "Recordset Name:"
    cmbFieldType.Enabled = False
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    btnWriteChangesToFile.Enabled = True
    btnNewField.Enabled = False
    btnNewRecordset.Enabled = False
    btnEndRecordset.Enabled = True
    boolNewRecordset = True
    Text1(0).SetFocus
    Exit Sub
    
errbtnNewRecordset_Click:
    Beep
    MsgBox Err.Description, vbCritical, "New Recordset"
    Err.Clear
End Sub


Private Sub btnOpen_Click()
    On Error GoTo bOpen_Click
    Set rsOpenRecordset = dbOpenDatabase.OpenRecordset(List1.List(iListIndex))
    Frame1.Caption = rsOpenRecordset.Name
    Frame4.Caption = rsOpenRecordset.Name
    
    posFields.Clear
    lstFields.Clear
    typeFields.Clear
    sizeFields.Clear
    
    For i = 0 To rsOpenRecordset.Fields.Count - 1
        posFields.AddItem i
        lstFields.AddItem rsOpenRecordset.Fields(i).Name
        Select Case rsOpenRecordset.Fields(i).Type
        Case dbBoolean
            FieldType = "Boolsk"
        Case dbByte
            FieldType = "Byte"
        Case dbInteger
            FieldType = "Integer"
        Case dbLong
            FieldType = "Long"
        Case dbCurrency
            FieldType = "Currency"
        Case dbSingle
            FieldType = "Single"
        Case dbDouble
            FieldType = "Double"
        Case dbDate
            FieldType = "Date"
        Case dbText
            FieldType = "Text"
        Case dbLongBinary
            FieldType = "LongBinary"
        Case dbMemo
            FieldType = "Memo"
        Case dbGUID
            FieldType = "GUID"
        End Select
        typeFields.AddItem FieldType
        sizeFields.AddItem rsOpenRecordset.Fields(i).Size
    Next
    Exit Sub
    
bOpen_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Open recordset"
    Err.Clear
End Sub
Private Sub btnPrint_Click()
    If Check2.Value = 1 Then
        PrintWithPreview
        Exit Sub
    End If
    
    If Check3.Value = 1 Then
        frmDatabasePrint.txtDBPath.Text = sDatabaseName
        frmDatabasePrint.Show 1
        Exit Sub
    End If
    
    
    On Error GoTo errbtnPrint_Click
    Set wdApp = New Word.Application
    With wdApp
        .Application.Visible = True
        .Documents.Add
        .Caption = "Database Print"
        .Documents.Application.WindowState = wdWindowStateMaximize
        .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3) _
                , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
         .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(7.5) _
                 , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
         .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10.5) _
                , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With
        
    With wdApp.ActiveWindow.Selection
        .Font.Bold = wdToggle
        .Font.Size = 16
        .TypeText Text:="Database: " & dbOpenDatabase.Name
        .TypeParagraph
        .TypeText Text:="Table Definition for:  " & rsOpenRecordset.Name
        .TypeParagraph
        .TypeParagraph
        .Font.Size = 12
        .TypeText Text:="Pos Nr." & _
              vbTab & "Description" & _
              vbTab & "Type" & _
              vbTab & "Size"
        .TypeParagraph
        .Font.Bold = wdToggle
    End With
        
    For i = 0 To posFields.ListCount - 1
        With wdApp.ActiveWindow.Selection
            .TypeParagraph
            .TypeText Text:=posFields.List(i) & _
                  vbTab & lstFields.List(i) & _
                  vbTab & typeFields.List(i) & _
                  vbTab & sizeFields.List(i)
        End With
    Next
    
    Set wdApp = Nothing
    Beep
    Exit Sub
    
errbtnPrint_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Printing"
    Err.Clear
End Sub

Private Sub btnPrintAll_Click()
Dim n As Integer, iCounter As Integer, FloodValue, iAnt As Integer
    On Error GoTo errbPrintAll_Click
    
    If Check2.Value = 1 Then
        PrintWithPreviewAll
        Exit Sub
    End If
    
    If Check3.Value = 1 Then
        frmDatabasePrint.txtDBPath.Text = sDatabaseName
        frmDatabasePrint.Show 1
        Exit Sub
    End If
    
    Set wdApp = New Word.Application
    With wdApp
        .Application.Visible = True
        .Documents.Add
        .Caption = "Database Print"
        .Documents.Application.WindowState = wdWindowStateMaximize
        .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(3) _
                , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
         .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(7.5) _
                 , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
         .ActiveWindow.Selection.ParagraphFormat.TabStops.Add Position:=CentimetersToPoints(10.5) _
                , Alignment:=wdAlignTabLeft, Leader:=wdTabLeaderSpaces
    End With
    
    Frame3.Visible = False
    iAnt = List1.ListCount
    
    For n = 0 To List1.ListCount - 1
        With wdApp.ActiveWindow.Selection
            .InsertBreak Type:=wdPageBreak
            .Font.Bold = wdToggle
            .Font.Size = 16
            .TypeText Text:="Database: " & dbOpenDatabase.Name
            .TypeParagraph
            .TypeText Text:="Table Definition for:  " & _
                rsOpenRecordset.Name
            .TypeParagraph
            .TypeParagraph
            .Font.Size = 12
            .TypeText Text:="Pos Nr." & _
                vbTab & "Description" & _
                vbTab & "Type" & _
                vbTab & "Size"
            .TypeParagraph
            .Font.Bold = wdToggle
        End With
        
        Set rsOpenRecordset = dbOpenDatabase.OpenRecordset(List1.List(List1.ListIndex))
        
        For i = 0 To rsOpenRecordset.Fields.Count - 1
            Select Case rsOpenRecordset.Fields(i).Type
            Case dbBoolean
                FieldType = "Boolsk"
            Case dbByte
                FieldType = "Byte"
            Case dbInteger
                FieldType = "Integer"
            Case dbLong
                FieldType = "Long"
            Case dbCurrency
                FieldType = "Currency"
            Case dbSingle
                FieldType = "Single"
            Case dbDouble
                FieldType = "Double"
            Case dbDate
                FieldType = "Date"
            Case dbText
                FieldType = "Text"
            Case dbLongBinary
                FieldType = "LongBinary"
            Case dbMemo
                FieldType = "Memo"
            Case dbGUID
                FieldType = "GUID"
            End Select
            
            wdApp.ActiveWindow.Selection.TypeParagraph
            wdApp.ActiveWindow.Selection.TypeText Text:=i & _
                  vbTab & rsOpenRecordset.Fields(i).Name & _
                  vbTab & FieldType & _
                  vbTab & rsOpenRecordset.Fields(i).Size
            
            DoEvents
        Next
    Next
    
    Set wdApp = Nothing
    
    Beep
    MsgBox "Printing Finished !!"
    Exit Sub
    
errbPrintAll_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Printing"
    Err.Clear
    Frame3.Visible = True
End Sub

Private Sub btnWriteChangesToFile_Click()
    'On Error GoTo errbtnWriteChangesToFile_Click
    FileNo = FileNo + 1
    fileLog = "DatabaseUpdate" & FileNo & ".txt"
    'Open fileLog For Output As #1
    
    If boolNewRecordset Then    'user wants to create a new recordset within the given database
        'Call FindDatabaseName
        sTmp = dbOpenDatabase.Name
        sTemp2 = Trim(Text1(0).Text)
        Write #1, sTmp, sTemp2, "New"
        
        'make this new recordset
        Set tdfNew = dbOpenDatabase.CreateTableDef(sTemp2)
        
        List2(0).Clear
        List2(1).Clear
        List2(2).Clear
        List2(3).Clear
        
        iNumberOfFields = -1
        boolNewRecordset = False
        btnNewField.Enabled = True
        boolNewField = True
        boolRecordWrite = True
        btnWriteChangesToFile.Enabled = False
        
        Label7(8).Caption = "Field Name:"
        cmbFieldType.Enabled = True
        Text1(0).Text = " "
        Text1(1).Enabled = True
        Text1(2).Enabled = True
        Exit Sub
    End If
    
    'user wants to append new field(s) to the recordset
    strFieldType = cmbFieldType.Text
    sTemp1 = Text1(0).Text
    sTemp2 = Text1(1).Text
    Write #1, sTemp1, strFieldType, sTemp2
    
    Call AppendField
    
    iNumberOfFields = iNumberOfFields + 1
    List2(0).AddItem iNumberOfFields
    List2(1).AddItem sTemp1
    List2(2).AddItem strFieldType
    List2(3).AddItem sTemp2
    Text1(0).SetFocus
    btnNewField.Enabled = True
    btnWriteChangesToFile.Enabled = False
    Exit Sub
    
errbtnWriteChangesToFile_Click:
    Beep
    MsgBox Err.Description, vbCritical, "Write changes"
    Err.Clear
End Sub

Private Sub Check1_Click()
    
    If Check1.Value = 1 Then
        Check2.Value = 0
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Check1.Value = 0
    End If
End Sub

Private Sub cmbFieldType_Click()
    'check which field type is chosen
        Select Case cmbFieldType.Text
        Case "Boolean"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        Case "Byte"
            Text1(1).Enabled = False
        Case "Integer"
            Text1(1).SetFocus
        Case "Long"
            Text1(1).SetFocus
        Case "Currency"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        Case "Single"
            Text1(1).SetFocus
        Case "Double"
            Text1(1).SetFocus
        Case "Date"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        Case "Text"
            Text1(1).Text = 50
            Text1(1).SetFocus
        Case "LongBinary"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        Case "Memo"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        Case "GUID"
            Text1(1).Enabled = False
            Text1(2).Enabled = False
        End Select
End Sub

Private Sub DirList_Change()
    On Error GoTo errDirList_Change
    ' Update the file list box to synchronize with the directory list box.
    FilList.Path = DirList.Path
    iDirIndex = DirList.ListIndex
    Exit Sub
errDirList_Change:
    Beep
    MsgBox Error$, 48, "Change Dictory List"
    Resume errDirList_Change2
errDirList_Change2:
End Sub
Private Sub DrvList_Change()
    On Error GoTo DriveHandler
    DirList.Path = DrvList.Drive
    Exit Sub
DriveHandler:
    Beep
    MsgBox Error$, 48, "Change Drive List"
    Resume Drivehandler2
Drivehandler2:
    DrvList.Drive = DirList.Path
    Exit Sub
End Sub
Private Sub FilList_Click()
    On Error GoTo errFilList_Click
    iFilIndex = FilList.ListIndex
    List1.Clear
    Call OpenLocalDB
    Call GetTableList
    Exit Sub
errFilList_Click:
    Beep
    MsgBox Err.Description, vbExclamation, "Change File List"
    Err.Clear
End Sub

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
    With cmbFieldType
        .Clear
        .AddItem "Boolean"
        .AddItem "Byte"
        .AddItem "Integer"
        .AddItem "Long"
        .AddItem "Currency"
        .AddItem "Single"
        .AddItem "Double"
        .AddItem "Date"
        .AddItem "Text"
        .AddItem "LongBinary"
        .AddItem "Memo"
        .AddItem "GUID"
    End With
    ReadText
    fileLog = "DatabaseUpdate.txt"
    FileNo = 0
    DisableButtons 1
End Sub

Private Sub Form_Load()
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmPrintDB")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    m_iFormNo = 17
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If boolNewField Then    'we have an open field-file
        Close #1
    End If
    If boolRecordWrite Then 'we have to append this new recordset before closing
        dbOpenDatabase.TableDefs.Append tdfNew
    End If
    dbOpenDatabase.Close
    rsLanguage.Close
    rsUser.Close
    m_iFormNo = 0
    DisableButtons 2
    frmMDI.Toolbar1.Buttons(11).Enabled = True
    Set frmPrintDB = Nothing
End Sub
Private Sub List1_Click()
    iListIndex = List1.ListIndex
End Sub
Private Sub lstFields_Click()
    iIndex = lstFields.ListIndex
    posFields.ListIndex = iIndex
    typeFields.ListIndex = iIndex
    sizeFields.ListIndex = iIndex
End Sub
Private Sub posFields_Click()
    iIndex = posFields.ListIndex
    lstFields.ListIndex = iIndex
    typeFields.ListIndex = iIndex
    sizeFields.ListIndex = iIndex
End Sub
Private Sub Timer1_Timer()
    If lstFields.TopIndex <> posFields.TopIndex Then
        lstFields.TopIndex = posFields.TopIndex
        typeFields.TopIndex = posFields.TopIndex
        sizeFields.TopIndex = posFields.TopIndex
    End If
End Sub

