VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00404040&
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   9465
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      ForeColor       =   &H00FFFFFF&
      Height          =   7455
      Left            =   1680
      TabIndex        =   10
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.Data rsMyRecord 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "User"
         Top             =   240
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "Language"
         DataSource      =   "rsInvoice"
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Caption         =   "Invoice Lines"
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
         Height          =   3615
         Left            =   240
         TabIndex        =   19
         Top             =   3720
         Width           =   7095
         Begin MSDBGrid.DBGrid Grid1 
            Bindings        =   "frmInvoice.frx":0000
            Height          =   3135
            Left            =   120
            OleObjectBlob   =   "frmInvoice.frx":001C
            TabIndex        =   20
            Top             =   360
            Width           =   6855
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "InvoiceNo"
         DataSource      =   "rsInvoice"
         Height          =   285
         Index           =   6
         Left            =   3720
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00404040&
         Caption         =   "Delivery Address"
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
         Height          =   2175
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   7095
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "PaymentDays"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   1335
         End
         Begin VB.ComboBox cmbPayment 
            BackColor       =   &H00FFFFC0&
            DataField       =   "Payment"
            DataSource      =   "rsInvoice"
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryCountry"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   5
            Left            =   3480
            TabIndex        =   8
            Top             =   1680
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryTown"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   4
            Left            =   3480
            TabIndex        =   7
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryZip"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   3
            Left            =   3480
            TabIndex        =   6
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryAddress3"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   2
            Left            =   3480
            TabIndex        =   5
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryAddress2"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   1
            Left            =   3480
            TabIndex        =   4
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CustmerDeliveryAddress1"
            DataSource      =   "rsInvoice"
            Height          =   285
            Index           =   0
            Left            =   3480
            TabIndex        =   3
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Days:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Terms:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   16
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   3255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   3135
         End
      End
      Begin VB.ComboBox cmbCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         DataField       =   "CustomerName"
         DataSource      =   "rsInvoice"
         Height          =   315
         Left            =   3720
         TabIndex        =   2
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Data rsInvoice 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Invoice"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data rsInvoiceLine 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\Programming.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "InvoiceLine"
         Top             =   360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSMask.MaskEdBox Mask1 
         DataField       =   "InvoiceDate"
         DataSource      =   "rsInvoice"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   10
         Mask            =   "##.##.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Giro ?"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   12
         Left            =   600
         TabIndex        =   31
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   ": Language"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   1080
         TabIndex        =   24
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "VAT ?"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   7245
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   29
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCustomBook() As Variant, vInvoiceBook() As Variant
Dim iNoLines As Integer, dblLineSum As Long, dblInvoiceSum As Long, dblSumVAT As Long
Dim PaymentDate As Date
Dim bNewRecord As Boolean
Dim WClone As Recordset
Dim rsCustomer As Recordset
Dim rsPayment As Recordset
Dim rsDueDate As Recordset
Dim rsLanguage As Recordset
Dim rsLanguage2 As Recordset
Private Sub PrintHeading()
    cPrint.FontSize = 10
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyName")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyName")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyAddress1")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyAddress1")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyAddress2")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyAddress2")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyZip")) And Not IsNull(rsMyRecord.Recordset.Fields("CompanyTown")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyZip") & "  " & rsMyRecord.Recordset.Fields("CompanyTown")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCountry")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyCountry")
    End If
    cPrint.pPrint
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 14
    If Not IsNull(rsLanguage2.Fields("Invoice")) Then
        cPrint.pRightTab rsLanguage2.Fields("Invoice")
    End If
    cPrint.FontBold = False
    cPrint.FontSize = 10
    cPrint.pPrint cmbCustomer.Text, 1, True
    If Not IsNull(rsLanguage2.Fields("InvoiceNo")) And Len(Text1(6).Text) <> 0 Then
        cPrint.pRightTab rsLanguage2.Fields("InvoiceNo") & "  " & Text1(6).Text
    End If
    cPrint.FontSize = 12
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pPrint Text1(0).Text, 1
    End If
    If Len(Text1(3).Text) <> 0 And Len(Text1(4).Text) <> 0 Then
        cPrint.pPrint Text1(3).Text & " " & Text1(4).Text, 1
    End If
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Text1(5).Text, 1, True
    End If
    cPrint.FontSize = 10
    If Not IsNull(rsLanguage2.Fields("Date")) Then
        cPrint.pRightTab rsLanguage2.Fields("Date") & "  " & Mask1.FormattedText
    End If
    cPrint.pPrint
    cPrint.pBox 0.25, cPrint.CurrentY, , 0.2, , &HFFFF&, vbFSSolid
    cPrint.FontBold = True
    cPrint.FontTransparent = True
    cPrint.pPrint rsLanguage2.Fields("ItemNo"), 0.3, True
    cPrint.pPrint rsLanguage2.Fields("ItemText"), 1, True
    cPrint.pPrint rsLanguage2.Fields("APrice"), 4.5, True
    cPrint.pPrint rsLanguage2.Fields("Quantity"), 5.5, True
    cPrint.pPrint rsLanguage2.Fields("Currency"), 6.2, True
    cPrint.pPrint rsLanguage2.Fields("SumLine"), 7
    cPrint.FontBold = False
    cPrint.BackColor = -1
    cPrint.pPrint
    cPrint.pPrint
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
                For i = 0 To 12
                    If IsNull(i + 1) Then
                        .Fields(i + 1) = Label2(i).Caption
                    Else
                        Label2(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("Frame2")) Then
                    .Fields("Frame2") = Frame2.Caption
                Else
                    Frame2.Caption = .Fields("Frame2")
                End If
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
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
        For i = 0 To 12
            .Fields(i + 1) = Label2(i).Caption
        Next
        .Fields("Frame2") = Frame2.Caption
        .Fields("Frame3") = Frame3.Caption
        .Fields("Help") = sHelp
        .Fields("Invoice") = "INVOICE"
        .Fields("To") = "To:"
        .Fields("InvoiceNo") = "Invoice No.:"
        .Fields("Date") = "Date:"
        .Fields("Page") = "Page:"
        .Fields("Of") = "Of:"
        .Fields("ItemNo") = "Item No."
        .Fields("ItemText") = "Item Text"
        .Fields("APrice") = "A Price"
        .Fields("Quantity") = "Quantity"
        .Fields("Currency") = "Currency"
        .Fields("SumLine") = "Sum Line"
        .Fields("InvoiceDate") = "Invoice Date:"
        .Fields("InvoiceDueDate") = "Invoice Due Date:"
        .Fields("SumThisInvoice") = "Sum This Invoice:"
        .Fields("SumTotalInvoice") = "Invoice Total:"
        .Fields("GiroReciept") = "Reciept"
        .Fields("GiroToAccount") = "Paid to account"
        .Fields("GiroAmount") = "Amount"
        .Fields("GiroPayAccountNo") = "Paid to account no."
        .Fields("GiroFormularNo") = "Formular No."
        .Fields("GiroPayInfo") = "Pay infomation"
        .Fields("GiroPay") = "Due-"
        .Fields("GiroPayTerm") = "date:"
        .Fields("GiroSignature") = "Signature"
        .Fields("GiroPaidBy") = "Paid by"
        .Fields("GiroPaidTo") = "Paid to"
        .Fields("GiroCharge") = "Debit"
        .Fields("GiroReceipt") = "Receipt"
        .Fields("GiroAccount") = "account"
        .Fields("GiroBack") = "back"
        .Fields("GiroCustomerId") = "Customer identification (KID)"
        .Fields("GiroCurrency") = "Dollar"
        .Fields("GiroCurrencySmall") = "Cent"
        .Fields("GiroPhone") = "Phone:"
        .Fields("GiroFax") = "Fax:"
        .Fields("GiroEmail") = "Email:"
        .Update
    End With
End Sub

Private Sub LoadDueDate()
    With rsDueDate
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("Language")) = Trim(rsInvoice.Recordset.Fields("Language")) Then Exit Sub
        .MoveNext
        Loop
        
        'use English text
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("Language")) = "ENG" Then Exit Do
        .MoveNext
        Loop
    End With
End Sub

Private Sub SelectRecords()
Dim Sql As String
    On Error Resume Next
    Sql = "SELECT * FROM InvoiceLine WHERE Clng(InvoiceNo) ="
    Sql = Sql & Chr(34) & CLng(rsInvoice.Recordset.Fields("InvoiceNo")) & Chr(34)
    rsInvoiceLine.RecordSource = Sql
    rsInvoiceLine.Refresh
    iNoLines = rsInvoiceLine.Recordset.RecordCount
End Sub
Private Sub LoadcmbCustomer()
    On Error Resume Next
    cmbCustomer.Clear
    With rsCustomer
        .MoveLast
        .MoveFirst
        ReDim vCustomBook(.RecordCount)
        Do While Not .EOF
            cmbCustomer.AddItem .Fields("CustomerName")
            cmbCustomer.ItemData(cmbCustomer.NewIndex) = cmbCustomer.ListCount - 1
            vCustomBook(cmbCustomer.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    Set WClone = rsInvoice.Recordset.Clone()
    With WClone
        .MoveLast
        .MoveFirst
        ReDim vInvoiceBook(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("InvoiceNo")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            vInvoiceBook(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
    Set WClone = Nothing
End Sub
Private Sub SetGridText()
    With Grid1
        .Columns(0).Caption = "Invoice No."
        .Columns(1).Caption = "Line No."
        .Columns(2).Caption = "Item ID"
        .Columns(3).Caption = "Item Text"
        .Columns(4).Caption = "Item Price"
        .Columns(5).Caption = "Currency"
        .Columns(6).Caption = "Quantity"
    End With
End Sub
Private Sub WriteFootnote()
    'On Error Resume Next
    cPrint.CurrentY = 10.7
    cPrint.pLine 0.3
    cPrint.FontBold = True
    cPrint.pCenter rsMyRecord.Recordset.Fields("CompanyName") & " - " & _
                    rsMyRecord.Recordset.Fields("CompanyAddress1") & " - " & _
                    rsMyRecord.Recordset.Fields("CompanyZip") & " " & rsMyRecord.Recordset.Fields("CompanyTown")
    cPrint.pCenter rsLanguage2.Fields("GiroPhone") & " " & rsMyRecord.Recordset.Fields("CompanyPhoneNo") & " - " & _
                    rsLanguage2.Fields("GiroFax") & " " & rsMyRecord.Recordset.Fields("CompanyFaxNo") & " - " & _
                    rsLanguage2.Fields("GiroEmail") & " " & rsMyRecord.Recordset.Fields("CompanyEMail")
End Sub

Private Sub WriteGiroPartPreview()
    'On Error Resume Next
    PaymentDate = DateAdd("d", rsInvoice.Recordset.Fields("PaymentDays"), CDate(Mask1.FormattedText))
    cPrint.CurrentY = 6.7
    cPrint.pPrint rsInvoice.Recordset.Fields("Payment"), 0.3
    cPrint.pPrint rsDueDate.Fields("DueDateText"), 0.3
    cPrint.pBox 0.3, cPrint.CurrentY, , 0.81, , &HFFFF&, vbFSSolid
    cPrint.pBox 2.59, 7.3, 1.4, 0.4, &HFFFFFF, &HFFFFFF, vbFSSolid
    cPrint.pBox 4.4, 7.3, 1.61, 0.4, &HFFFFFF, &HFFFFFF, vbFSSolid
    cPrint.pBox 6.14, 7.3, 1.2, 0.4, &HFFFFFF, &HFFFFFF, vbFSSolid
    cPrint.BackColor = -1
    cPrint.FontBold = True
    cPrint.pPrint rsLanguage2.Fields("GiroReciept"), 0.32
    cPrint.FontSize = 8
    cPrint.pPrint rsLanguage2.Fields("GiroToAccount"), 0.32, True
    cPrint.pPrint rsLanguage2.Fields("GiroAmount"), 2.59, True
    cPrint.pPrint rsLanguage2.Fields("GiroPayAccountNo"), 4.52, True
    cPrint.pPrint rsLanguage2.Fields("GiroFormularNo"), 6.22, False
    cPrint.FontSize = 10
    cPrint.FontBold = False
    cPrint.pPrint
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyName")) Then
        cPrint.pPrint rsMyRecord.Recordset.Fields("CompanyName"), 0.3, True
    End If
    cPrint.pPrint Format(dblInvoiceSum, "###,###.00"), 2.8
    cPrint.pPrint
    cPrint.FontSize = 8
    cPrint.FontBold = True
    cPrint.pHalfSpace
    cPrint.pPrint rsLanguage2.Fields("GiroPayInfo"), 0.3, True
    cPrint.pPrint rsLanguage2.Fields("GiroPay"), 5.43
    cPrint.pPrint rsLanguage2.Fields("GiroPayTerm"), 5.43, True
    cPrint.FontBold = False
    cPrint.FontSize = 10
    cPrint.pPrint Format(PaymentDate, "dd.mm.yyyy"), 6.29
    cPrint.pPrint rsLanguage2.Fields("GiroSignature"), 3.9
    cPrint.pPrint
    cPrint.pPrint rsLanguage2.Fields("InvoiceNo") & "    " & Text1(6).Text, 0.3
    cPrint.pPrint rsLanguage2.Fields("InvoiceDate") & "  " & CDate(Mask1.FormattedText), 0.3
    cPrint.pPrint rsLanguage2.Fields("InvoiceDueDate") & "   " & Format(CDate(PaymentDate), "dd.mm.yyyy"), 0.3
    cPrint.pPrint
    cPrint.FontSize = 8
    cPrint.FontBold = True
    cPrint.pPrint rsLanguage2.Fields("GiroPaidBy"), 0.3, True
    cPrint.pPrint rsLanguage2.Fields("GiroPaidTo"), 3.9
    cPrint.FontBold = False
    cPrint.FontSize = 10
    cPrint.pPrint cmbCustomer.Text, 0.3, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyName")) Then
        cPrint.pPrint rsMyRecord.Recordset.Fields("CompanyName"), 3.9
    End If
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pPrint Text1(0).Text, 0.3, True
    Else
        cPrint.pPrint " ", , True
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyAddress1")) Then
        cPrint.pPrint rsMyRecord.Recordset.Fields("CompanyAddress1"), 3.9
    End If
    If Len(Text1(3).Text) <> 0 And Len(Text1(4).Text) <> 0 Then
        cPrint.pPrint Text1(3).Text & "  " & Text1(4).Text, 0.3, True
    Else
        cPrint.pPrint " ", , True
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyZip")) And Not IsNull(rsMyRecord.Recordset.Fields("CompanyTown")) Then
        cPrint.pPrint rsMyRecord.Recordset.Fields("CompanyZip") & " " & rsMyRecord.Recordset.Fields("CompanyTown"), 3.9
    End If
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Text1(5).Text, 0.3
    Else
        cPrint.pPrint
    End If
    cPrint.pBox 0.3, cPrint.CurrentY, , 0.3, -1, &HFFFF&, vbFSSolid
    cPrint.pBox 1.5, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 1.73, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 1.96, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 2.19, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 2.42, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 2.65, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 2.88, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 3.11, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 3.34, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 3.57, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 3.8, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.pBox 7.08, cPrint.CurrentY + 0.04, 0.19, 0.23, &H0&, &HFFFFFF, vbFSSolid
    cPrint.FontSize = 8
    cPrint.FontBold = True
    cPrint.pPrint rsLanguage2.Fields("GiroCharge"), 1, True
    cPrint.pPrint rsLanguage2.Fields("GiroReceipt"), 6.4
    cPrint.pPrint rsLanguage2.Fields("GiroAccount"), 1, True
    cPrint.pPrint rsLanguage2.Fields("GiroBack"), 6.4
    cPrint.pVerticalLine 0.3, 10.3, 10.69
    cPrint.pVerticalLine 2.4, 10.3, 10.69
    cPrint.pVerticalLine 3.5, 10.3, 10.69
    cPrint.pVerticalLine 4.78, 10.3, 10.69
    cPrint.pVerticalLine 6.35, 10.3, 10.69
    cPrint.pQuarterSpace
    cPrint.pPrint rsLanguage2.Fields("GiroCustomerId"), 0.3, True
    cPrint.pPrint rsLanguage2.Fields("GiroCurrency"), 2.5, True
    cPrint.pPrint rsLanguage2.Fields("GiroCurrencySmall"), 3.66, True
    cPrint.pPrint rsLanguage2.Fields("GiroAccount"), 4.88, True
    cPrint.pPrint rsLanguage2.Fields("GiroFormularNo"), 6.45
    cPrint.FontSize = 10
    cPrint.FontBold = False
    cPrint.pPrint
    cPrint.pPrint dblInvoiceSum, 2.5, True
    cPrint.pPrint "00", 3.66, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyBankAcount")) Then
        cPrint.pPrint rsMyRecord.Recordset.Fields("CompanyBankAcount"), 4.88
    End If
    cPrint.pBox 0.3, 10.8, , 0.1, -1, &HFFFF&, vbFSSolid
End Sub

Private Sub WriteHeadingPreView()
Dim bFound As Boolean
    'On Error Resume Next
    'find the customer language if exist
    bFound = True
    With rsLanguage2
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("Language")) = Trim(Text1(7).Text) Then
                bFound = False
                Exit Do
            End If
        .MoveNext
        Loop
        
        'we did not find this particular language, use english
        If bFound Then
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("Language")) = "ENG" Then Exit Do
        .MoveNext
        Loop
        End If
    End With
    
    cPrint.pStartDoc
    cPrint.FontSize = 10
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyName")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyName")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyAddress1")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyAddress1")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyAddress2")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyAddress2")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyZip")) And Not IsNull(rsMyRecord.Recordset.Fields("CompanyTown")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyZip") & "  " & rsMyRecord.Recordset.Fields("CompanyTown")
    End If
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCountry")) Then
        cPrint.pRightTab rsMyRecord.Recordset.Fields("CompanyCountry")
    End If
    cPrint.pPrint
    cPrint.pPrint
    cPrint.FontBold = True
    cPrint.FontSize = 14
    If Not IsNull(rsLanguage2.Fields("Invoice")) Then
        cPrint.pRightTab rsLanguage2.Fields("Invoice")
    End If
    cPrint.FontBold = False
    cPrint.FontSize = 10
    cPrint.pPrint cmbCustomer.Text, 1, True
    If Not IsNull(rsLanguage2.Fields("InvoiceNo")) And Len(Text1(6).Text) <> 0 Then
        cPrint.pRightTab rsLanguage2.Fields("InvoiceNo") & "  " & Text1(6).Text
    End If
    cPrint.FontSize = 12
    If Len(Text1(0).Text) <> 0 Then
        cPrint.pPrint Text1(0).Text, 1
    End If
    If Len(Text1(3).Text) <> 0 And Len(Text1(4).Text) <> 0 Then
        cPrint.pPrint Text1(3).Text & " " & Text1(4).Text, 1
    End If
    If Len(Text1(5).Text) <> 0 Then
        cPrint.pPrint Text1(5).Text, 1, True
    End If
    cPrint.FontSize = 10
    If Not IsNull(rsLanguage2.Fields("Date")) Then
        cPrint.pRightTab rsLanguage2.Fields("Date") & "  " & Mask1.FormattedText
    End If
    cPrint.pPrint
    cPrint.pBox 0.25, cPrint.CurrentY, , 0.2, , &HFFFF&, vbFSSolid
    cPrint.FontBold = True
    cPrint.FontTransparent = True
    cPrint.pPrint rsLanguage2.Fields("ItemNo"), 0.3, True
    cPrint.pPrint rsLanguage2.Fields("ItemText"), 1, True
    cPrint.pPrint rsLanguage2.Fields("APrice"), 4.5, True
    cPrint.pPrint rsLanguage2.Fields("Quantity"), 5.5, True
    cPrint.pPrint rsLanguage2.Fields("Currency"), 6.2, True
    cPrint.pPrint rsLanguage2.Fields("SumLine"), 7
    cPrint.FontBold = False
    cPrint.BackColor = -1
    cPrint.pPrint
    cPrint.pPrint
End Sub


Private Sub WriteLinesPreView()
    'On Error Resume Next
    cPrint.pRightJust rsInvoiceLine.Recordset.Fields("Item ID"), 0.8, True
    cPrint.pPrint Format(rsInvoiceLine.Recordset.Fields("Item Text")), 1.1, True
    cPrint.pRightJust Format(rsInvoiceLine.Recordset.Fields("Item Price"), "###,###.00"), 5, True
    cPrint.pRightJust rsInvoiceLine.Recordset.Fields("Quantity"), 5.8, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCurrency")) Then
        cPrint.pRightJust rsMyRecord.Recordset.Fields("CompanyCurrency"), 6.6, True
    End If
    cPrint.pRightJust Format(dblLineSum, "###,###.00"), 7.5, False
    If cPrint.CurrentY >= 5.7 Then
        WriteFootnote
        cPrint.pNewPage
        PrintHeading
    End If
End Sub


Private Sub WriteSumInvoicePreView()
    'On Error Resume Next
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("SumTotalInvoice"), 1.1, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCurrency")) Then
        cPrint.pRightJust rsMyRecord.Recordset.Fields("CompanyCurrency"), 6.6, True
    End If
    cPrint.pRightJust Format(dblInvoiceSum, "###,###.00"), 7.5, False
End Sub

Private Sub WriteSumLinePreView()
    'On Error Resume Next
    cPrint.pPrint
    cPrint.pPrint rsLanguage.Fields("SumThisInvoice"), 1.1, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCurrency")) Then
        cPrint.pRightJust rsMyRecord.Recordset.Fields("CompanyCurrency"), 6.6, True
    End If
    cPrint.pRightJust Format(dblInvoiceSum, "###,###.00"), 7.5, False
End Sub

Public Sub DeleteInvoice()
    On Error Resume Next
    With rsInvoiceLine.Recordset
        .MoveFirst
        Do While Not .EOF
            If CLng(.Fields("InvoiceNo")) = CLng(Text1(6).Text) Then
                .Delete
            End If
        .MoveNext
        Loop
    End With
    rsInvoice.Recordset.Delete
    LoadList1
    List1.ListIndex = 0
    SelectRecords
End Sub

Public Sub NewInvoice()
    On Error GoTo errNewInvoice
    rsInvoice.Recordset.AddNew
    bNewRecord = True
    'find the next free invoice no
    With rsMyRecord.Recordset
        Text1(6).Text = .Fields("NextInvoiceNo")
        .Edit
        .Fields("NextInvoiceNo") = CLng(.Fields("NextInvoiceNo")) + 1
        .Update
    End With
    Mask1.Text = Format(Now, "dd.mm.yyyy")
    Mask1.SetFocus
    Exit Sub
    
errNewInvoice:
    Beep
    MsgBox Err.Description, vbCritical, "New Invoice"
    Err.Clear
    'put the invoice counter back in place
    With rsMyRecord.Recordset
        .Edit
        .Fields("NextInvoiceNo") = CLng(.Fields("NextInvoiceNo")) - 1
        .Update
    End With
    On Error Resume Next    'just in case
    bNewRecord = False
    List1.ListIndex = 0
End Sub

Private Sub WriteWATLinePreView()
    'On Error Resume Next
    cPrint.pPrint
    cPrint.pPrint "VAT (" & rsMyRecord.Recordset.Fields("CompanyVAT") & " %)", 1.1, True
    If Not IsNull(rsMyRecord.Recordset.Fields("CompanyCurrency")) Then
        cPrint.pRightJust rsMyRecord.Recordset.Fields("CompanyCurrency"), 6.6, True
    End If
    cPrint.pRightJust Format(dblSumVAT, "###,###.00"), 7.5, False
End Sub

Public Sub PrintInvoice()
    dblLineSum = 0
    dblInvoiceSum = 0
    dblSumVAT = 0
    
    'On Error Resume Next
    LoadDueDate
    
    Set cPrint = New clsMultiPgPreview
    
    frmPrinterSetUp.Show vbModal
    If QuitCommand Then
        QuitCommand = False
        Set cPrint = Nothing
        Exit Sub
    End If
    
SendToPrinter:
    Screen.MousePointer = vbHourglass
    WriteHeadingPreView
    'write the Invoice lines
    With rsInvoiceLine.Recordset
        .MoveFirst
        Do While Not .EOF
            dblLineSum = CDbl(.Fields("Item Price")) * CLng(.Fields("Quantity"))
            dblInvoiceSum = dblInvoiceSum + (CDbl(.Fields("Item Price")) * CLng(.Fields("Quantity")))
            WriteLinesPreView
        .MoveNext
        Loop
    End With
    
    WriteSumLinePreView
    
    'do we have to calculate VAT ?
    If Check1.Value = 1 Then
        dblSumVAT = (dblInvoiceSum / 100) * rsMyRecord.Recordset.Fields("CompanyVAT")
        WriteWATLinePreView
        dblInvoiceSum = dblInvoiceSum + dblSumVAT
        WriteSumInvoicePreView
    End If
    
    If Check2.Value = 1 Then
        'now write the giro part of the Invoice
        WriteGiroPartPreview
    Else
        WriteFootnote
    End If
    
    cPrint.pEndDoc
    If cPrint.SendToPrinter Then GoTo SendToPrinter
    Set cPrint = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Sub UpdateInvoice()
    With rsInvoice.Recordset
        .Edit
        .Fields("InvoiceDate") = Mask1.FormattedText
        .Fields("CustomerName") = cmbCustomer.Text
        If Len(Text1(0).Text) <> 0 Then
            .Fields("CustmerDeliveryAddress1") = Text1(0).Text
        Else
            .Fields("CustmerDeliveryAddress1") = " "
        End If
        If Len(Text1(1).Text) <> 0 Then
            .Fields("CustmerDeliveryAddress2") = Text1(1).Text
        Else
            .Fields("CustmerDeliveryAddress2") = " "
        End If
        If Len(Text1(2).Text) <> 0 Then
            .Fields("CustmerDeliveryAddress3") = Text1(2).Text
        Else
            .Fields("CustmerDeliveryAddress3") = " "
        End If
        If Len(Text1(3).Text) <> 0 Then
            .Fields("CustmerDeliveryZip") = Text1(3).Text
        Else
            .Fields("CustmerDeliveryZip") = " "
        End If
        If Len(Text1(4).Text) <> 0 Then
            .Fields("CustmerDeliveryTown") = Text1(4).Text
        Else
            .Fields("CustmerDeliveryTown") = " "
        End If
        If Len(Text1(5).Text) <> 0 Then
            .Fields("CustmerDeliveryCountry") = Text1(5).Text
        Else
            .Fields("CustmerDeliveryCountry") = " "
        End If
        If Len(Text1(7).Text) <> 0 Then
            .Fields("Language") = Text1(7).Text
        Else
            .Fields("Language") = " "
        End If
        If Len(Text1(8).Text) <> 0 Then
            .Fields("PaymentDays") = CInt(Text1(8).Text)
        Else
            .Fields("PaymentDays") = 0
        End If
        If Len(cmbPayment.Text) <> 0 Then
            .Fields("Payment") = cmbPayment.Text
        Else
            .Fields("Payment") = " "
        End If
        If Check1.Value = 1 Then
            .Fields("CalcVAT") = True
        Else
            .Fields("CalcVAT") = False
        End If
        .Update
    End With
End Sub
Private Sub cmbCustomer_LostFocus()
    'On Error Resume Next
    If bNewRecord Then
        With rsInvoice.Recordset
            .Fields("InvoiceNo") = CLng(Text1(6).Text)
            .Fields("InvoiceDate") = CDate(Mask1.FormattedText)
            .Fields("CustomerAutoLineNo") = CLng(rsCustomer.Fields("AutoLine"))
            .Fields("CustomerName") = cmbCustomer.Text
            .Update
            .Bookmark = .LastModified
            LoadList1
            SelectRecords
        End With
        bNewRecord = False
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsInvoice.Refresh
    rsInvoiceLine.Refresh
    rsMyRecord.Refresh
    LoadcmbCustomer
    DoEvents
    LoadList1
    List1.ListIndex = 0
    DoEvents
    SetGridText
    DoEvents
    ReadText
    DisableButtons 2
    frmMDI.Toolbar1.Buttons(8).Enabled = False
    frmMDI.Toolbar1.Buttons(9).Enabled = False
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsInvoice.DatabaseName = m_strPrograming
    rsInvoiceLine.DatabaseName = m_strPrograming
    rsMyRecord.DatabaseName = m_strPrograming
    Set rsCustomer = m_dbPrograming.OpenRecordset("Customer")
    Set rsPayment = m_dbPrograming.OpenRecordset("Payment")
    Set rsDueDate = m_dbPrograming.OpenRecordset("DueDateText")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmInvoice")
    Set rsLanguage2 = m_dbLanguage.OpenRecordset("frmInvoice")
    m_iFormNo = 8
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub


Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsInvoice.Recordset.Close
    rsInvoiceLine.Recordset.Close
    rsMyRecord.Recordset.Close
    rsCustomer.Close
    rsPayment.Close
    rsDueDate.Close
    rsLanguage.Close
    rsLanguage2.Close
    m_iFormNo = 0
    DisableButtons 1
    Set frmInvoice = Nothing
End Sub

Private Sub cmbCustomer_Click()
        On Error Resume Next
        With rsCustomer
            .Bookmark = vCustomBook(cmbCustomer.ItemData(cmbCustomer.ListIndex))
            Text1(0).Text = Format(.Fields("CustomerAdress1"))
            Text1(1).Text = Format(.Fields("CustomerAdress2"))
            Text1(3).Text = Format(.Fields("CustomerZip"))
            Text1(4).Text = Format(.Fields("CustomerTown"))
            Text1(5).Text = Format(.Fields("CustomerCountry"))
            Text1(7).Text = Format(.Fields("Language"))
            If Not IsNull(.Fields("Payment")) Then
                cmbPayment.Text = .Fields("Payment")
                rsPayment.MoveFirst
                Do While Not rsPayment.EOF
                    If rsPayment.Fields("Payment") = .Fields("Payment") Then
                        Text1(8).Text = CInt(rsPayment.Fields("PaymentDays"))
                        Exit Do
                    End If
                rsPayment.MoveNext
                Loop
            End If
        End With
End Sub

Private Sub Grid1_AfterDelete()
    iNoLines = iNoLines - 1
End Sub

Private Sub Grid1_OnAddNew()
    Grid1.Columns(0).Text = CLng(Text1(6).Text)
    iNoLines = iNoLines + 1
    Grid1.Columns(1).Text = iNoLines
    Grid1.Columns(5).Text = rsMyRecord.Recordset.Fields("CompanyCurrency")
End Sub
Private Sub List1_Click()
        On Error Resume Next
        rsInvoice.Recordset.Bookmark = vInvoiceBook(List1.ItemData(List1.ListIndex))
        SelectRecords
        Frame1.Caption = "Invoice No.: " & rsInvoice.Recordset.Fields("InvoiceNo")
End Sub

Private Sub Mask1_LostFocus()
    If bNewRecord Then
        cmbCustomer.SetFocus
    End If
End Sub


