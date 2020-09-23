VERSION 5.00
Begin VB.Form frmMaint 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   ClientHeight    =   6870
   ClientLeft      =   2100
   ClientTop       =   1260
   ClientWidth     =   11325
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmMaint.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6870
   ScaleWidth      =   11325
   Begin VB.CommandButton bStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8040
      TabIndex        =   11
      Top             =   5520
      Width           =   3135
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   3960
      TabIndex        =   9
      Top             =   240
      Width           =   3975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4515
      Left            =   8040
      TabIndex        =   8
      Top             =   240
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   5
      Top             =   4920
      Width           =   3135
   End
   Begin VB.DirListBox DirList1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   3960
      TabIndex        =   4
      Top             =   4920
      Width           =   3975
   End
   Begin VB.FileListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   240
      Pattern         =   "*.mdb"
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   5400
      Width           =   3615
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Result / Errors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Compacted Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Databases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const Default = 0        ' 0 - Default
Private Const HOURGLASS = 11     ' 11 - Hourglass
Dim rsLanguage As Recordset
Dim strDatabase As String, tmpName As String
Dim strPath As String, strLogFile As String
Dim iErrNumber As Integer
Dim intList1Index As Integer
Dim n As Long
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
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("label2")
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3.Caption
                Else
                    Label3.Caption = .Fields("label3")
                End If
                If IsNull(.Fields("label4")) Then
                    .Fields("label4") = Label4.Caption
                Else
                    Label4.Caption = .Fields("label4")
                End If
                If IsNull(.Fields("bStart")) Then
                    .Fields("bStart") = bStart.Caption
                Else
                    bStart.Caption = .Fields("bStart")
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
        .Fields("label1") = Label1.Caption
        .Fields("label2") = Label2.Caption
        .Fields("label3") = Label3.Caption
        .Fields("label4") = Label4.Caption
        .Fields("bStart") = bStart.Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub bStart_Click()
    Me.MousePointer = HOURGLASS
    On Error GoTo errbStart_Click
    iErrNumber = 0
    tmpName = List1.Path & "\temp.mdb"
    strLogFile = List1.Path & "\RepairLog.tmp"
    Open strLogFile For Output As #1
    
    For intList1Index = 0 To List1.ListCount - 1
        List1.ListIndex = intList1Index
        strDatabase = strPath & "\" & List1.List(intList1Index)
        Label5.Caption = strDatabase
        Label5.Refresh
        Call CompactDb
    Next
    
    Label4.Caption = " "
    Label4.Refresh
    
    If iErrNumber = 0 Then
        Label5.Caption = "Finished sucessful operation"
        On Error Resume Next
        Close #1
        Kill strLogFile
    Else
        On Error Resume Next
        Label5.Caption = "Finished operation - found " & iErrNumber & " errors !!" & _
            " See logfile: " & strLogFile
        Close #1
    End If
    
    Label5.Refresh
    Me.MousePointer = Default
    Exit Sub
    
errbStart_Click:
    Beep
    Me.MousePointer = Default
    iErrNumber = iErrNumber + 1
    Resume Next
End Sub
Private Sub CompactDb()
Dim strOldName As String, strNewName As String
Dim sSizeOld As String, sSizeNew As String
        Label4.Caption = "Compacting: "
        Label4.Refresh
        n = Len(strDatabase)
        strOldName = Left(strDatabase, (n - 3))
        strOldName = strOldName & "bck"
        strNewName = strDatabase
        
        'do we have a leftover from last compact ?
        On Error Resume Next
        Kill tmpName
        
        On Error GoTo 0
        On Error GoTo errCompactDb
        sSizeOld = GetFileSize(strDatabase)
        DBEngine.CompactDatabase strDatabase, tmpName
        DoEvents
        
        On Error Resume Next
        Kill strOldName
        
        On Error GoTo errCompactDb
        Name strDatabase As strOldName
        Name tmpName As strNewName
        sSizeNew = GetFileSize(strNewName)
        DoEvents
        
        List3.AddItem "Compact OK: " & List1.List(intList1Index)
        List2.AddItem "Before: " & sSizeOld & "  -  After: " & sSizeNew
        DoEvents
        Exit Sub
        
errCompactDb:
        Beep
        Me.MousePointer = Default
        iErrNumber = iErrNumber + 1
        List3.AddItem "Error Compacting: " & List1.List(intList1Index)
        Write #1, List1.List(intList1Index) & ":  " & Err.Description
        List2.AddItem Err.Description
        Resume errCompactDb2
errCompactDb2:
End Sub

Private Sub DirList1_Change()
    strPath = DirList1.Path
    List1.Path = DirList1.Path
End Sub

Private Sub Drive1_Change()
    DirList1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
    Me.WindowState = vbMaximized
    DisableButtons 1
    ReadText
End Sub

Private Sub Form_Initialize()
    List3.Clear
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmMaint")
    m_iFormNo = 30
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    m_iFormNo = 0
    DisableButtons 2
    rsLanguage.Close
    Set frmMaint = Nothing
End Sub
