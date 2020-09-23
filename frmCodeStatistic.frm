VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCodeStatistic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code Statistic"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmCodeStatistic.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   7080
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Data rsUserSnippet 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\MyOwnSnippets.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CodeSnippet"
         Top             =   2520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox cboLanguage 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   1
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   6015
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   10610
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code language:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ComboBox cboLanguage 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Index           =   0
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
      Begin VB.Data rsCodeSnippet 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\CodeSnippets.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CodeSnippet"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsCodeZip 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\CodeZip.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "CodeZip"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   6015
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   10610
         _Version        =   393216
         AllowUserResizing=   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Code language:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCodeStatistic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim boolFirst As Boolean
Dim rsUser As Recordset
Dim rsCodeLanguage As Recordset
Dim rsLanguage As Recordset
Dim sName As String
Dim sType As String, lSum As Long
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
End Sub

Private Sub AutosizeGridColumns(msFG As MSFlexGrid, ByVal MaxRowsToParse As Integer, _
  ByVal MaxColWidth As Integer)
    Dim i As Integer, J As Integer
    Dim txtString As String
    Dim intTempWidth, intBiggestWidth As Integer
    Dim intRows As Integer
    Const intPadding = 150
    
    On Error Resume Next
    With msFG

         .Col = 0
         ' Set the active colunm
         intRows = .Rows
         ' Set the number of rows
         If intRows > MaxRowsToParse Then intRows = MaxRowsToParse
         ' If there are more rows of data, reset intRows to the MaxRowsToParse constant
        
         intBiggestWidth = 0
         ' Reset some values to 0

         For J = 0 To intRows - 1
             ' check up to MaxRowsToParse # of rows and obtain the greatest width of the cell contents
             .Row = J
             txtString = .Text
             intTempWidth = TextWidth(txtString) + intPadding
             ' The intPadding constant compensates for text insets You can adjust this value above as desired.
             
             If intTempWidth > intBiggestWidth Then intBiggestWidth = intTempWidth
             ' Reset intBiggestWidth to the intMaxColWidth value if necessary
         Next J
         
         .ColWidth(0) = intBiggestWidth
        ' Now check to see if the columns aren't  as wide as the grid itself.
        ' If not, determine the difference and expand each column proportionately
        ' to fill the grid
        intTempWidth = 0
        intTempWidth = intTempWidth + .ColWidth(0)
            ' Add up the width of all the columns
       
        If intTempWidth < msFG.Width Then
            ' Compate the width of the columns to the width of the grid control
            ' and if necessary expand the columns.
            intTempWidth = Fix((msFG.Width - intTempWidth) / .Cols)
            ' Determine the amount od width expansion needed by each column
            .ColWidth(0) = .ColWidth(0) + intTempWidth
            ' add the necessary width to each column
        End If
    End With
End Sub

Private Sub LoadCodeLanguage()
    On Error Resume Next
    With rsCodeLanguage
        .MoveFirst
        Do While Not .EOF
            cboLanguage(0).AddItem .Fields("Language")
            cboLanguage(1).AddItem .Fields("Language")
        .MoveNext
        Loop
    End With
    With rsUser
        If Not IsNull(.Fields("PrefferedLanguage")) Then
            If m_boolSnippet Then
                cboLanguage(0).Text = .Fields("PrefferedLanguage")
                cboLanguage(1).Text = .Fields("PrefferedLanguage")
            Else
                cboLanguage(0).Text = .Fields("PrefferedLanguage")
            End If
        End If
    End With
End Sub
Private Sub LoadGrid1()
    On Error Resume Next
    ' MyOwnSnippets.mdb:
    i = 0
    lSum = 0
    sType = " "
    boolFirst = True
    With rsUserSnippet.Recordset
        .MoveFirst
        Do While Not .EOF
            If Not boolFirst Then
                If .Fields("CodeType") <> sType Then
                    Grid1(1).AddItem sType & vbTab & i
                    sType = .Fields("CodeType")
                    i = 1
                    lSum = lSum + 1
                Else
                    i = i + 1
                    lSum = lSum + 1
                End If
            Else
                sType = .Fields("CodeType")
                i = 1
                lSum = lSum + 1
                boolFirst = False
            End If
        .MoveNext
        Loop
        Grid1(1).AddItem sType & vbTab & i
        Grid1(1).AddItem " " & vbTab & " "
        Grid1(1).AddItem "Sum: " & vbTab & lSum
    End With
    Call AutosizeGridColumns(Grid1(1), 1000, 1000)
End Sub

Private Sub ReadText()
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
                    .Fields("label1") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("label1")
                    Label1(1).Caption = .Fields("label1")
                End If
                Grid1(0).Row = 0
                Grid1(0).Col = 0
                If IsNull(.Fields("Grid1Coln0")) Then
                    .Fields("Grid1Coln0") = "Code Type"
                Else
                    Grid1(0).Text = .Fields("Grid1Coln0")
                End If
                Grid1(0).Row = 0
                Grid1(0).Col = 1
                If IsNull(.Fields("Grid1Coln1")) Then
                    .Fields("Grid1Coln1") = "Quantity"
                Else
                    Grid1(0).Text = .Fields("Grid1Coln1")
                End If
                Grid1(1).Row = 0
                Grid1(1).Col = 0
                If IsNull(.Fields("Grid1Coln0")) Then
                    .Fields("Grid1Coln0") = "Code Type"
                Else
                    Grid1(1).Text = .Fields("Grid1Coln0")
                End If
                Grid1(1).Row = 0
                Grid1(1).Col = 1
                If IsNull(.Fields("Grid1Coln1")) Then
                    .Fields("Grid1Coln1") = "Quantity"
                Else
                    Grid1(1).Text = .Fields("Grid1Coln1")
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
        .Fields("label1") = Label1(0).Caption
        .Fields("Grid1Coln0") = Grid1(0).Text
        .Fields("Grid1Coln1") = Grid1(1).Text
        .Fields("btnExit") = btnExit.ToolTipText
        .Update
    End With
End Sub


Private Sub LoadGrid0()
    On Error Resume Next
    If m_boolSnippet Then
        ' CodeSnippets.mdb
        i = 0
        lSum = 0
        sType = " "
        boolFirst = True
        With rsCodeSnippet.Recordset
            .MoveFirst
            Do While Not .EOF
                If Not boolFirst Then
                    If .Fields("CodeType") <> sType Then
                        Grid1(0).AddItem sType & vbTab & i
                        sType = .Fields("CodeType")
                        i = 1
                        lSum = lSum + 1
                    Else
                        i = i + 1
                        lSum = lSum + 1
                    End If
                Else
                    sType = .Fields("CodeType")
                    i = 1
                    lSum = lSum + 1
                    boolFirst = False
                End If
            .MoveNext
            Loop
            Grid1(0).AddItem sType & vbTab & i
            Grid1(0).AddItem " " & vbTab & " "
            Grid1(0).AddItem "Sum: " & vbTab & lSum
        End With
    Else
        ' CodeZip.mdb:
        i = 0
        lSum = 0
        sType = " "
        boolFirst = True
        With rsCodeZip.Recordset
            .MoveFirst
            Do While Not .EOF
                If Not boolFirst Then
                    If .Fields("CodeType") <> sType Then
                        Grid1(0).AddItem sType & vbTab & i
                        sType = .Fields("CodeType")
                        i = 1
                        lSum = lSum + 1
                    Else
                        i = i + 1
                        lSum = lSum + 1
                    End If
                Else
                    sType = .Fields("CodeType")
                    i = 1
                    lSum = lSum + 1
                    boolFirst = False
                End If
            .MoveNext
            Loop
            Grid1(0).AddItem sType & vbTab & i
            Grid1(0).AddItem " " & vbTab & " "
            Grid1(0).AddItem "Sum: " & vbTab & lSum
        End With
    End If
    Call AutosizeGridColumns(Grid1(0), 1000, 1000)
End Sub

Private Sub SelectGridCell(GridRow As Integer, Index As Integer)
    Dim i As Integer
    Dim J As Integer
    Dim NumRows As Integer
    Dim NumCols As Integer
    
    On Error Resume Next
    NumRows = Grid1(Index).Rows - 1 '.rows returns num of rows
    NumCols = Grid1(Index).Cols - 1 '.cols reutrns num of columns
    Grid1(Index).Highlight = flexHighlightNever 'since this Sub takes
    'care of highlighting we tell it to never highlight so only 1 row
    'is selected at a time

    For i = 1 To NumRows
        If i <> GridRow Then
            Grid1(Index).Row = i

            For J = 1 To NumCols
                Grid1(Index).Col = J
                If Grid1(Index).CellBackColor = vbHighlight Then
                    Grid1(Index).CellBackColor = vbWindowBackground
                    Grid1(Index).CellForeColor = vbWindowText
                Else
                    Exit For
                End If
            Next J
        End If
    Next i
    Grid1(Index).Row = GridRow 'set the row To the clicked row

    For i = 1 To NumCols 'setting the clicked row to highlighted
        Grid1(Index).Col = i
        Grid1(Index).CellBackColor = vbHighlight
        Grid1(Index).CellForeColor = vbHighlightText
    Next i
    'note when leaving this sub the msflexgrid.row will be the gridrow
    'and the column will be the last column.   i.e. if there are 4 columns
    'then the Grid1.col will be 4.
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub cboLanguage_Click(Index As Integer)
Dim Sql As String
    On Error Resume Next
    If m_boolSnippet Then
        Sql = "SELECT * FROM CodeSnippet"
        rsCodeSnippet.RecordSource = Sql
        rsCodeSnippet.Refresh
        
        Sql = "SELECT * FROM CodeSnippet"
        rsUserSnippet.RecordSource = Sql
        rsUserSnippet.Refresh
    Else
        Sql = "SELECT * FROM CodeZip"
        rsCodeZip.RecordSource = Sql
        rsCodeZip.Refresh
    End If
    
    Select Case Index
    Case 0
        Grid1(0).Clear
        Grid1(0).Rows = 2
        If m_boolSnippet Then
            Sql = "SELECT * FROM CodeSnippet WHERE Trim(CodeLanguage) ="
            Sql = Sql & Chr(34) & Trim(cboLanguage(0).Text) & Chr(34)
            Sql = Sql & " ORDER BY CodeType"
            With rsCodeSnippet
                .RecordSource = Sql
                .Refresh
                If Not .Recordset.EOF And Not .Recordset.BOF Then
                    .Recordset.MoveFirst
                End If
            End With
            LoadGrid0
        Else
            Sql = "SELECT * FROM CodeZip WHERE Trim(CodeLanguage) ="
            Sql = Sql & Chr(34) & Trim(cboLanguage(0).Text) & Chr(34)
            Sql = Sql & " ORDER BY CodeType"
            With rsCodeZip
                .RecordSource = Sql
                .Refresh
                If Not .Recordset.EOF And Not .Recordset.BOF Then
                    .Recordset.MoveFirst
                End If
            End With
            LoadGrid0
        End If
    Case 1
        Grid1(1).Clear
        Grid1(1).Rows = 2
        Sql = "SELECT * FROM CodeSnippet WHERE Trim(CodeLanguage) ="
        Sql = Sql & Chr(34) & Trim(cboLanguage(1).Text) & Chr(34)
        Sql = Sql & " ORDER BY CodeType"
        With rsUserSnippet
            .RecordSource = Sql
            .Refresh
            If Not .Recordset.EOF And Not .Recordset.BOF Then
                .Recordset.MoveFirst
            End If
        End With
        LoadGrid1
    Case Else
    End Select
    boolFirst = True
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsCodeSnippet.Refresh
    rsUserSnippet.Refresh
    rsCodeZip.Refresh
    LoadCodeLanguage
    ReadText
    LoadBackground
End Sub
Private Sub Form_Load()
Dim dbTemp As Database
    On Error GoTo errForm_Load
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsCodeLanguage = m_dbCodeSnippet.OpenRecordset("Language")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCodeStatistic")
    rsCodeSnippet.DatabaseName = m_strCodeSnippet
    rsUserSnippet.DatabaseName = rsUser.Fields("OwnSnippetName")
    rsCodeZip.DatabaseName = m_strCodeZip
    
    If m_boolSnippet Then
        Frame1(0).Caption = ExtractFileName(m_dbCodeSnippet.Name)
        Frame1(1).Caption = ExtractFileName(rsUser.Fields("OwnSnippetName"))
        Me.Width = Me.Width + 4200
        btnExit.Width = btnExit.Width + 275
    Else
        Frame1(0).Caption = ExtractFileName(m_strCodeZip)
    End If
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    TileForm Me, Picture1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCodeSnippet.Recordset.Close
    rsUserSnippet.Recordset.Close
    rsCodeZip.Recordset.Close
    rsCodeLanguage.Close
    rsUser.Close
    rsLanguage.Close
    Set frmCodeStatistic = Nothing
End Sub


Private Sub Grid1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SelectGridCell(Grid1(Index).Row, Index)
End Sub


