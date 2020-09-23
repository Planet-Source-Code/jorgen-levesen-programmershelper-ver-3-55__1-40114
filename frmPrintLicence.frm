VERSION 5.00
Begin VB.Form frmPrintLicence 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Licence"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   120
      Picture         =   "frmPrintLicence.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton btnPrint 
      Height          =   495
      Left            =   2520
      Picture         =   "frmPrintLicence.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Only shown project"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All projects"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmPrintLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLicence As Recordset
Private Sub PrintHead()
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 18
    cPrint.pBox 0.25, cPrint.CurrentY, cPrint.GetPaperWidth - 0.25, 0.3, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pCenter "Licence"
    cPrint.BackColor = -1
    cPrint.FontSize = 12
    cPrint.FontBold = False
    cPrint.pPrint
End Sub

Private Sub PrintAllProjects()
Dim sProject As String, boolFirst As Boolean
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter2:
    Screen.MousePointer = vbHourglass
    On Error GoTo errPrintAllProjects
    boolFirst = True
    sProject = ""
    cPrint.pStartDoc
    
    PrintHead
    
    With rsLicence
        .Index = "ProjectID"
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("ProjectID")) = Trim(sProject) Then
                cPrint.pPrint
                cPrint.pPrint frmLicence.Label1(1).Caption, 0.7, True
                If Not IsNull(.Fields("CustomerName")) Then
                    cPrint.pPrint .Fields("CustomerName"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(3).Caption, 0.7, True
                cPrint.pPrint .Fields("NoOfProgrammes"), 3
                cPrint.pPrint frmLicence.Label1(5).Caption, 0.7, True
                If IsDate(.Fields("PurchaseDate")) Then
                    cPrint.pPrint .Fields("PurchaseDate"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(4).Caption, 0.7, True
                If Not IsNull(.Fields("LiberationKey")) Then
                    cPrint.pPrint .Fields("LiberationKey"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(8).Caption, 0.7, True
                If Not IsNull(.Fields("ProgrammeVersion")) Then
                    cPrint.pPrint .Fields("ProgrammeVersion"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    PrintHead
                End If
            Else
                If boolFirst Then
                    sProject = Trim(.Fields("ProjectID"))
                    cPrint.pPrint
                    cPrint.FontBold = True
                    cPrint.pPrint frmLicence.Label1(2).Caption, 0.7, True
                    cPrint.pPrint .Fields("ProjectID"), 3
                    cPrint.pPrint " "
                    cPrint.FontBold = False
                    cPrint.pPrint frmLicence.Label1(1).Caption, 0.7, True
                    If Not IsNull(.Fields("CustomerName")) Then
                        cPrint.pPrint .Fields("CustomerName"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(3).Caption, 0.7, True
                    cPrint.pPrint .Fields("NoOfProgrammes"), 3
                    cPrint.pPrint frmLicence.Label1(5).Caption, 0.7, True
                    If IsDate(.Fields("PurchaseDate")) Then
                        cPrint.pPrint .Fields("PurchaseDate"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(4).Caption, 0.7, True
                    If Not IsNull(.Fields("LiberationKey")) Then
                        cPrint.pPrint .Fields("LiberationKey"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(8).Caption, 0.7, True
                    If Not IsNull(.Fields("ProgrammeVersion")) Then
                        cPrint.pPrint .Fields("ProgrammeVersion"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    boolFirst = False
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        PrintHead
                    End If
                Else
                    'print first project
                    cPrint.pNewPage
                    PrintHead
                    sProject = Trim(.Fields("ProjectID"))
                    cPrint.pPrint
                    cPrint.FontBold = True
                    cPrint.pPrint frmLicence.Label1(2).Caption, 0.7, True
                    cPrint.pPrint .Fields("ProjectID"), 3
                    cPrint.FontBold = False
                    cPrint.pPrint " "
                    cPrint.pPrint frmLicence.Label1(1).Caption, 0.7, True
                    If Not IsNull(.Fields("CustomerName")) Then
                        cPrint.pPrint .Fields("CustomerName"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(3).Caption, 0.7, True
                    cPrint.pPrint .Fields("NoOfProgrammes"), 3
                    cPrint.pPrint frmLicence.Label1(5).Caption, 0.7, True
                    If IsDate(.Fields("PurchaseDate")) Then
                        cPrint.pPrint .Fields("PurchaseDate"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(4).Caption, 0.7, True
                    If Not IsNull(.Fields("LiberationKey")) Then
                        cPrint.pPrint .Fields("LiberationKey"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    cPrint.pPrint frmLicence.Label1(8).Caption, 0.7, True
                    If Not IsNull(.Fields("ProgrammeVersion")) Then
                        cPrint.pPrint .Fields("ProgrammeVersion"), 3
                    Else
                        cPrint.pPrint " ", 3
                    End If
                    If cPrint.pEndOfPage Then
                        cPrint.pFooter
                        cPrint.pNewPage
                        PrintHead
                    End If
                End If
            End If
        .MoveNext
        Loop
    End With
    
    cPrint.pFooter
    cPrint.pEndDoc
    Screen.MousePointer = vbHourglass
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter2
    Set cPrint = Nothing
    Exit Sub
    
errPrintAllProjects:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Print All Customers"
    Err.Clear
End Sub
Private Sub PrintOneProject()
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter2:
    Screen.MousePointer = vbHourglass
    On Error GoTo errPrintOneProjects
    cPrint.pStartDoc
    PrintHead
    
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.pPrint frmLicence.Label1(2).Caption, 0.7, True
    cPrint.pPrint frmLicence.Text1(0).Text, 3
    cPrint.FontBold = False
    cPrint.pPrint
    
    With rsLicence
        .Index = "ProjectID"
        .MoveFirst
        Do While Not .EOF
            If .Fields("ProjectID") = frmLicence.Text1(0).Text Then
                cPrint.pPrint
                cPrint.pPrint frmLicence.Label1(1).Caption, 0.7, True
                If Not IsNull(.Fields("CustomerName")) Then
                    cPrint.pPrint .Fields("CustomerName"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(3).Caption, 0.7, True
                cPrint.pPrint .Fields("NoOfProgrammes"), 3
                cPrint.pPrint frmLicence.Label1(5).Caption, 0.7, True
                If IsDate(.Fields("PurchaseDate")) Then
                    cPrint.pPrint .Fields("PurchaseDate"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(4).Caption, 0.7, True
                If Not IsNull(.Fields("LiberationKey")) Then
                    cPrint.pPrint .Fields("LiberationKey"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                cPrint.pPrint frmLicence.Label1(8).Caption, 0.7, True
                If Not IsNull(.Fields("ProgrammeVersion")) Then
                    cPrint.pPrint .Fields("ProgrammeVersion"), 3
                Else
                    cPrint.pPrint " ", 3
                End If
                If cPrint.pEndOfPage Then
                    cPrint.pFooter
                    cPrint.pNewPage
                    PrintHead
                End If
            End If
        .MoveNext
        Loop
    End With
    cPrint.pFooter
    cPrint.pEndDoc
    Screen.MousePointer = vbHourglass
    
    If cPrint.SendToPrinter Then GoTo SendToPrinter2
    Set cPrint = Nothing
    Exit Sub
    
errPrintOneProjects:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Print one project"
    Err.Clear
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    If Option1(0).Value = True Then
        PrintAllProjects
    Else
        PrintOneProject
    End If
    MsgBox "Printing finished !"
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsLicence = m_dbPrograming.OpenRecordset("Licence")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLicence.Close
    Set frmPrintLicence = Nothing
End Sub
