VERSION 5.00
Begin VB.Form frmShowAuthors 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Show Author's snippets"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btnShow 
      Caption         =   "&Show Snippet"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Result"
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4695
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   3735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Author"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton btnSearch 
         Caption         =   "&Find"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmShowAuthors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodeBook() As Variant
Dim rsCodeSnippet As Recordset
Dim rsCodeZip As Recordset
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
                If IsNull(.Fields("Frame1")) Then
                    .Fields("Frame1") = Frame1.Caption
                Else
                    Frame1.Caption = .Fields("Frame1")
                End If
                If IsNull(.Fields("Frame2")) Then
                    .Fields("Frame2") = Frame2.Caption
                Else
                    Frame2.Caption = .Fields("Frame2")
                End If
                If IsNull(.Fields("btnSearch")) Then
                    .Fields("btnSearch") = btnSearch.Caption
                Else
                    btnSearch.Caption = .Fields("btnSearch")
                End If
                If IsNull(.Fields("btnShow")) Then
                    .Fields("btnShow") = btnShow.Caption
                Else
                    btnShow.Caption = .Fields("btnShow")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.Caption
                Else
                    btnExit.Caption = .Fields("btnExit")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("Frame1") = Frame1.Caption
        .Fields("Frame2") = Frame2.Caption
        .Fields("btnSearch") = btnSearch.Caption
        .Fields("btnShow") = btnShow.Caption
        .Fields("btnExit") = btnExit.Caption
        .Update
    End With
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnSearch_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippet
        With rsCodeSnippet
            .MoveLast
            .MoveFirst
            ReDim vCodeBook(.RecordCount)
            Do While Not .EOF
                If Trim(.Fields("Author")) = Trim(Text1.Text) Then
                    List1.AddItem .Fields("CodeText")
                    List1.ItemData(List1.NewIndex) = List1.ListCount - 1
                    vCodeBook(List1.ListCount - 1) = .Bookmark
                End If
            .MoveNext
            Loop
        End With
    Case 33
        With rsCodeZip
            .MoveLast
            .MoveFirst
            ReDim vCodeBook(.RecordCount)
            Do While Not .EOF
                If Trim(.Fields("Author")) = Trim(Text1.Text) Then
                    List1.AddItem .Fields("CodeText")
                    List1.ItemData(List1.NewIndex) = List1.ListCount - 1
                    vCodeBook(List1.ListCount - 1) = .Bookmark
                End If
            .MoveNext
            Loop
        End With
    Case Else
    End Select
End Sub

Private Sub btnShow_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippet
        m_lSnippet = CLng(rsCodeSnippet.Fields("CodeNo"))
        frmCodeSnippets.ShowAuthor
    Case 33 'code zip
        m_lSnippet = CLng(rsCodeZip.Fields("CodeAuto"))
        frmCodeZip.ShowAuthor
    Case Else
    End Select
    Unload Me
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsCodeSnippet = m_dbCodeSnippet.OpenRecordset("CodeSnippet")
    Set rsCodeZip = m_dbCodeZip.OpenRecordset("CodeZip")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmShowAuthors")
    ReadText
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbInformation, "Load Form"
    Err.Clear
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCodeSnippet.Close
    rsCodeZip.Close
    rsLanguage.Close
    Erase vCodeBook
    Set frmShowAuthors = Nothing
End Sub
Private Sub List1_Click()
    On Error Resume Next
    Select Case m_iFormNo
    Case 2  'code snippet
        rsCodeSnippet.Bookmark = vCodeBook(List1.ItemData(List1.ListIndex))
    Case 33 'code zip
        rsCodeZip.Bookmark = vCodeBook(List1.ItemData(List1.ListIndex))
    Case Else
    End Select
End Sub
