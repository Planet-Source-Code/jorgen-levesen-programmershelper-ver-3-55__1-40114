VERSION 5.00
Begin VB.Form frmPrintCustomer 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Customers"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Picture         =   "frmPrintCustomer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton btnPrint 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Picture         =   "frmPrintCustomer.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Print"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Just Print Shown Customer"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Print All Customers"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPrintCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLanguage As Recordset
Private Sub PrintHead()
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 18
    cPrint.pBox 0.25, cPrint.CurrentY, cPrint.GetPaperWidth - 0.25, 0.3, &HC0E0FF, , vbFSSolid
    cPrint.BackColor = &HC0E0FF
    cPrint.pCenter "Customers"
    cPrint.BackColor = -1
    cPrint.FontSize = 12
    cPrint.FontBold = False
    cPrint.pPrint
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
                If IsNull(.Fields("Option1(0)")) Then
                    .Fields("Option1(0)") = Option1(0).Caption
                Else
                    Option1(0).Caption = .Fields("Option1(0)")
                End If
                If IsNull(.Fields("Option1(1)")) Then
                    .Fields("Option1(1)") = Option1(1).Caption
                Else
                    Option1(1).Caption = .Fields("Option1(1)")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.ToolTipText
                Else
                    btnExit.ToolTipText = .Fields("btnExit")
                End If
                If IsNull(.Fields("btnPrint")) Then
                    .Fields("btnPrint") = btnPrint.ToolTipText
                Else
                    btnPrint.ToolTipText = .Fields("btnPrint")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("Option1(0)") = Option1(0).Caption
        .Fields("Option1(1)") = Option1(1).Caption
        .Fields("btnExit") = btnExit.ToolTipText
        .Fields("btnPrint") = btnPrint.ToolTipText
        .Update
    End With
End Sub

Private Sub PrintAll()
    On Error GoTo errPrintAll
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter2:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    With frmCustomer.rsCustomer.Recordset
        .MoveFirst
        Do While Not .EOF
            PrintHead
            cPrint.FontBold = True
            cPrint.pPrint frmCustomer.Label1(0).Caption, 0.3, True 'name
            cPrint.pPrint .Fields("CustomerName"), 2.5
            cPrint.FontBold = False
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(1).Caption, 0.3, True 'address 1
            If Not IsNull(.Fields("CustomerAdress1")) Then
                cPrint.pPrint .Fields("CustomerAdress1"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            If Not IsNull(.Fields("CustomerAdress2")) Then
                cPrint.pPrint .Fields("CustomerAdress2"), 2.5   'address 2
            End If
            cPrint.pPrint frmCustomer.Label1(2).Caption, 0.3, True 'zip code
            If Not IsNull(.Fields("CustomerZip")) Then
                cPrint.pPrint .Fields("CustomerZip"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(3).Caption, 0.3, True 'town
            If Not IsNull(.Fields("CustomerTown")) Then
                cPrint.pPrint .Fields("CustomerTown"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(4).Caption, 0.3, True 'country
            If Not IsNull(.Fields("CustomerCountry")) Then
                cPrint.pPrint .Fields("CustomerCountry"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(5).Caption, 0.3, True 'phone prefix
            If Not IsNull(.Fields("CustomerPrefix")) Then
                cPrint.pPrint .Fields("CustomerPrefix"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(6).Caption, 0.3, True 'phone number
            If Not IsNull(.Fields("CustomerPhone")) Then
                cPrint.pPrint .Fields("CustomerPhone"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(7).Caption, 0.3, True 'Fax number
            If Not IsNull(.Fields("CustomerFax")) Then
                cPrint.pPrint .Fields("CustomerFax"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(8).Caption, 0.3, True 'email
            If Not IsNull(.Fields("CustomerEMail")) Then
                cPrint.pPrint .Fields("CustomerEMail"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(9).Caption, 0.3, True 'internet
            If Not IsNull(.Fields("CustomerURL")) Then
                cPrint.pPrint .Fields("CustomerURL"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(10).Caption, 0.3, True 'customer contact
            If Not IsNull(.Fields("CustomerContact")) Then
                cPrint.pPrint .Fields("CustomerContact"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(12).Caption, 0.3, True 'on mail list?
            If CBool(.Fields("CustomerOnMailList")) Then
                cPrint.pPrint "X", 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(11).Caption, 0.3, True 'wat calculation?
            If CBool(.Fields("CustomerVAT")) Then
                cPrint.pPrint "X", 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint frmCustomer.Label1(13).Caption, 0.3, True 'language on print
            If Not IsNull(.Fields("Language")) Then
                cPrint.pPrint .Fields("Language"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pPrint
            cPrint.pPrint frmCustomer.Label1(14).Caption, 0.3, True 'payment
            If Not IsNull(.Fields("Payment")) <> 0 Then
                cPrint.pPrint .Fields("Payment"), 2.5
            Else
                cPrint.pPrint " ", 2.5
            End If
            cPrint.pFooter
            cPrint.pNewPage
        .MoveNext
        Loop
    End With
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter2
    Set cPrint = Nothing
    Exit Sub
    
errPrintAll:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Print All Customers"
    Err.Clear
End Sub
Private Sub PrintShown()
    On Error GoTo errPrintShown
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter1:
    Screen.MousePointer = vbHourglass
    cPrint.pStartDoc
    PrintHead
    With frmCustomer
        cPrint.FontBold = True
        cPrint.pPrint .Label1(0).Caption, 0.3, True
        cPrint.pPrint .Text1(0).Text, 2.5
        cPrint.FontBold = False
        cPrint.pPrint
        cPrint.pPrint .Label1(1).Caption, 0.3, True
        If Len(.Text1(1).Text) <> 0 Then
            cPrint.pPrint .Text1(1).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        If Len(.Text1(2).Text) <> 0 Then
            cPrint.pPrint .Text1(2).Text, 2.5
        End If
        cPrint.pPrint .Label1(2).Caption, 0.3, True 'zip code
        If Len(.Text1(3).Text) <> 0 Then
            cPrint.pPrint .Text1(3).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(3).Caption, 0.3, True 'town
        If Len(.Text1(4).Text) <> 0 Then
            cPrint.pPrint .Text1(4).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(4).Caption, 0.3, True 'country
        If Len(.cmbCountry.Text) <> 0 Then
            cPrint.pPrint .cmbCountry.Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint
        cPrint.pPrint .Label1(5).Caption, 0.3, True 'phone prefix
        If Len(.Text1(6).Text) <> 0 Then
            cPrint.pPrint .Text1(6).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(6).Caption, 0.3, True 'phone number
        If Len(.Text1(7).Text) <> 0 Then
            cPrint.pPrint .Text1(7).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(7).Caption, 0.3, True 'Fax number
        If Len(.Text1(8).Text) <> 0 Then
            cPrint.pPrint .Text1(8).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint
        cPrint.pPrint .Label1(8).Caption, 0.3, True 'email
        If Len(.Text1(9).Text) <> 0 Then
            cPrint.pPrint .Text1(9).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(9).Caption, 0.3, True 'internet
        If Len(.Text1(10).Text) <> 0 Then
            cPrint.pPrint .Text1(10).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint
        cPrint.pPrint .Label1(10).Caption, 0.3, True 'customer contact
        If Len(.Text1(11).Text) <> 0 Then
            cPrint.pPrint .Text1(11).Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint
        cPrint.pPrint .Label1(12).Caption, 0.3, True 'on mail list?
        If .Check1(1).Value = 1 Then
            cPrint.pPrint "X", 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(11).Caption, 0.3, True 'wat calculation?
        If .Check1(0).Value = 1 Then
            cPrint.pPrint "X", 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint .Label1(13).Caption, 0.3, True 'language on print
        If Len(.cmbLanguage.Text) <> 0 Then
            cPrint.pPrint .cmbLanguage.Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
        cPrint.pPrint
        cPrint.pPrint .Label1(14).Caption, 0.3, True 'payment
        If Len(.cmbPayment.Text) <> 0 Then
            cPrint.pPrint .cmbPayment.Text, 2.5
        Else
            cPrint.pPrint " ", 2.5
        End If
    End With
    
    cPrint.pFooter
    cPrint.pEndDoc
    
    Screen.MousePointer = vbDefault
    If cPrint.SendToPrinter Then GoTo SendToPrinter1
    Set cPrint = Nothing
    Exit Sub
    
errPrintShown:
    Beep
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Print Shown Customer"
    Err.Clear
End Sub


Private Sub btnExit_Click()
    Unload Me
End Sub


Private Sub btnPrint_Click()
    If Option1(0).Value = True Then
        PrintAll
    Else
        PrintShown
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmPrintCustomer")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmPrintCustomer = Nothing
End Sub
