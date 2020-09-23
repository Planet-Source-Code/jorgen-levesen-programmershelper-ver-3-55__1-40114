VERSION 5.00
Begin VB.Form frmCustomer 
   BackColor       =   &H00404040&
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9480
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   8760
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   35
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerName"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   0
      Left            =   4320
      TabIndex        =   18
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerAdress1"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   17
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerAdress2"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerZip"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   3
      Left            =   4320
      TabIndex        =   15
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerTown"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   4
      Left            =   4320
      TabIndex        =   14
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerPrefix"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   6
      Left            =   4320
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerPhone"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   7
      Left            =   4320
      TabIndex        =   12
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerFax"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   8
      Left            =   4320
      TabIndex        =   11
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerEMail"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   9
      Left            =   4320
      TabIndex        =   10
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerURL"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   10
      Left            =   4320
      TabIndex        =   9
      Top             =   4800
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerContact"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   11
      Left            =   4320
      TabIndex        =   8
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      DataField       =   "CustomerVAT"
      DataSource      =   "rsCustomer"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   7
      Top             =   6240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      DataField       =   "CustomerOnMailList"
      DataSource      =   "rsCustomer"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   6240
      Width           =   255
   End
   Begin VB.ComboBox cmbLanguage 
      BackColor       =   &H00FFFFC0&
      DataField       =   "Language"
      DataSource      =   "rsCustomer"
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
   End
   Begin VB.ComboBox cmbCountry 
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerCountry"
      DataSource      =   "rsCustomer"
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.ComboBox cmbPayment 
      BackColor       =   &H00FFFFC0&
      DataField       =   "Payment"
      DataSource      =   "rsCustomer"
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Data rsCustomer 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeMaster.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton btnTransferToOutlook 
      Caption         =   "Transfer To Outlook"
      Height          =   975
      Left            =   7800
      Picture         =   "frmCustomer.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      DataField       =   "CustomerWindowsVersion"
      DataSource      =   "rsCustomer"
      Height          =   285
      Index           =   5
      Left            =   4320
      TabIndex        =   1
      Top             =   5640
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   7245
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   2760
      X2              =   2760
      Y1              =   7320
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   9360
      X2              =   9360
      Y1              =   7320
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2760
      X2              =   9360
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2760
      X2              =   9360
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   34
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Zip Code:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   32
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Town:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   31
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   30
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Prefix:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   29
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   28
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fax Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   27
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Adress:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   26
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Internet URL:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2760
      TabIndex        =   25
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Contact:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   24
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "VAT Calculation:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   5040
      TabIndex        =   23
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "On Mail List:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   22
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Language:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   2880
      TabIndex        =   21
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payment:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   6120
      TabIndex        =   20
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Version:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   2760
      TabIndex        =   19
      Top             =   5640
      Width           =   1455
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bookmarksPro() As Variant, vBookCountry() As Variant
Dim bNewRecord As Boolean
Dim rsCountry As Recordset
Dim rsUser As Recordset
Dim rsPayment As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 15
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 3
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Public Sub PrintLetter()
Dim wdApp As Word.Application, sDir As String
    'write the heading
    Set wdApp = New Word.Application
    sDir = AppDir
    wdApp.Application.Visible = True
    wdApp.Application.WindowState = wdWindowStateMaximize
    wdApp.Caption = "Write Letter"
    wdApp.Documents.Add sDir & "Letter.dot"
    With wdApp
        .ActiveWindow.ActivePane.view.SeekView = wdSeekCurrentPageHeader
        .Selection.MoveRight Unit:=wdCharacter, Count:=2
        .Selection.TypeText Text:=rsUser.Fields("CompanyName")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsUser.Fields("CompanyAddress1")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsUser.Fields("CompanyZip") & "  " & rsUser.Fields("CompanyTown")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsUser.Fields("CompanyCountry")
        If .Selection.HeaderFooter.IsHeader = True Then
            .ActiveWindow.ActivePane.view.SeekView = wdSeekCurrentPageFooter
        Else
            .ActiveWindow.ActivePane.view.SeekView = wdSeekCurrentPageHeader
        End If
        .Selection.TypeText Text:=rsUser.Fields("CompanyName") & _
            " -  " & rsUser.Fields("CompanyAddress1") & _
            " -  " & rsUser.Fields("CompanyZip") & "  " & rsUser.Fields("CompanyTown") & _
            " -  " & rsUser.Fields("CompanyCountry")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsLanguage.Fields("Phone") & "  " & "(" & rsUser.Fields("CompanyPrefixPhone") & ")" & rsUser.Fields("CompanyPhoneNo") & _
            "  -  " & rsLanguage.Fields("Fax") & "  " & "(" & rsUser.Fields("CompanyPrefixPhone") & ")" & rsUser.Fields("CompanyFaxNo")
        .ActiveWindow.ActivePane.view.SeekView = wdSeekMainDocument
        .Selection.GoTo What:=wdGoToBookmark, Name:="Name"
        .Selection.TypeText Text:=Text1(0).Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="Address1"
        .Selection.TypeText Text:=Text1(1).Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="Address2"
        .Selection.TypeText Text:=Text1(2).Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="ZipTown"
        .Selection.TypeText Text:=Text1(3).Text & "  " & Text1(4).Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="Country"
        .Selection.TypeText Text:=cmbCountry.Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="Contact"
        .Selection.TypeText Text:="Attn.:  " & Text1(11).Text
        .Selection.GoTo What:=wdGoToBookmark, Name:="Date"
        .Selection.TypeText Text:=rsLanguage.Fields("Date2") & "  " & Format(Now, "dd.mm.yyyy")
        .Selection.TypeText Text:=vbTab & rsLanguage.Fields("OurRef") & _
            vbTab & rsLanguage.Fields("YourRef")
        .Selection.MoveDown Unit:=wdLine, Count:=1
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsLanguage.Fields("Sincerely")
        .Selection.TypeParagraph
        .Selection.TypeText Text:=rsUser.Fields("CompanyName")
        .Selection.TypeParagraph
        .Selection.GoTo What:=wdGoToBookmark, Name:="Start"
        If Not IsNull(rsUser.Fields("LetterDirectory")) Then
            sDir = Trim(rsUser.Fields("LetterDirectory"))
        End If
        'save the document
        .ActiveDocument.SaveAs FileName:=sDir & "/" & Format(Now, "dd.mm.yyyy") & "-" & Text1(0).Text & ".doc", _
            FileFormat:=wdFormatDocument, _
            LockComments:=False, Password:="", AddToRecentFiles:=True, WritePassword _
            :="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
            SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
            False
    End With
    Set wdApp = Nothing
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
                For i = 0 To 15
                    If IsNull(.Fields(i + 1)) Then
                        .Fields(i + 1) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 1)
                    End If
                Next
                If IsNull(.Fields("btnTransferToOutlook")) Then
                    .Fields("btnTransferToOutlook") = btnTransferToOutlook.Caption
                Else
                    btnTransferToOutlook.Caption = .Fields("btnTransferToOutlook")
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
        For i = 0 To 15
            .Fields(i + 1) = Label1(i).Caption
        Next
        .Fields("btnTransferToOutlook") = btnTransferToOutlook.Caption
        .Fields("Phone") = "Phone:"
        .Fields("Fax") = "Fax:"
        .Fields("Date2") = "Date:"
        .Fields("OurRef") = "Our Ref.:"
        .Fields("YourRef") = "Your Ref.:"
        .Fields("Sincerely") = "Sincerely Yours"
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Private Sub LoadLanguage()
    With rsCountry
        .MoveLast
        .MoveFirst
        ReDim vBookCountry(.RecordCount)
        Do While Not .EOF
            cmbLanguage.AddItem .Fields("Country")
            cmbCountry.AddItem .Fields("Country")
            cmbLanguage.ItemData(cmbLanguage.NewIndex) = cmbLanguage.ListCount - 1
            vBookCountry(cmbLanguage.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub LoadList1()
    On Error Resume Next
    List1.Clear
    With rsCustomer.Recordset
        .MoveLast
        .MoveFirst
        ReDim bookmarksPro(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("CustomerName")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarksPro(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub LoadPayment()
    With rsPayment
        .MoveFirst
        Do While Not .EOF
            cmbPayment.AddItem .Fields("Payment")
        .MoveNext
        Loop
    End With
End Sub

Public Sub DeleteRecord()
Dim rsLicence As Recordset
    On Error Resume Next
    'delete any referance in the Licence recordset, if any that is
    Set rsLicence = m_dbPrograming.OpenRecordset("Licence")
    With rsLicence
        .MoveFirst
        Do While Not .EOF
            If Trim(.Fields("CustomerName")) = Trim(rsCustomer.Recordset.Fields("CustomerName")) Then
                .Delete
            End If
        .MoveNext
        Loop
    End With
    rsLicence.Close
    rsCustomer.Recordset.Delete
    LoadList1
End Sub


Public Sub NewRecord()
    rsCustomer.Recordset.AddNew
    bNewRecord = True
    Text1(0).SetFocus
End Sub

Private Sub btnTransferToOutlook_Click()
    Dim oOutlook As Outlook.Application
    Dim oContact As Outlook.ContactItem
    
    On Error GoTo errTrans
    Set oOutlook = New Outlook.Application
    Set oContact = oOutlook.CreateItem(olContactItem)
    With oContact
        .FullName = rsCustomer.Recordset.Fields("CustomerName")
        .BusinessAddress = Format(rsCustomer.Recordset.Fields("CustomerAdress1"))
        .BusinessAddressCity = Format(rsCustomer.Recordset.Fields("CustomerTown"))
        .BusinessAddressPostalCode = Format(rsCustomer.Recordset.Fields("CustomerZip"))
        .HomeTelephoneNumber = Format(rsCustomer.Recordset.Fields("CustomerPhone"))
        .Email1Address = Format(rsCustomer.Recordset.Fields("CustomerEMail"))
        .BusinessFaxNumber = Format(rsCustomer.Recordset.Fields("CustomerFax"))
        .BusinessHomePage = Format(rsCustomer.Recordset.Fields("CustomerURL"))
        .Save
    End With
    
    MsgBox "Contact has been Added", vbInformation
    Set oOutlook = Nothing
    Exit Sub
    
errTrans:
    Beep
    MsgBox Err.Description, vbExclamation, "Transfer to Outlook"
    Err.Clear
End Sub

Private Sub cmbLanguage_Click()
    On Error Resume Next
    rsCountry.Bookmark = vBookCountry(cmbLanguage.ItemData(cmbLanguage.ListIndex))
    cmbLanguage.Text = rsCountry.Fields("CountryFix")
End Sub


Private Sub Form_Activate()
    On Error Resume Next
    LoadPayment
    LoadLanguage
    rsCustomer.Refresh
    LoadList1
    List1.ListIndex = 0
    ReadText
    DisableButtons 2
    
    With frmMDI.Toolbar1
        .Buttons(8).Enabled = False
        .Buttons(9).Enabled = False
        .Buttons(13).Enabled = True
        .Buttons(14).Enabled = True
        .Buttons(15).Enabled = True
        .Buttons(16).Enabled = True
    End With
    Me.WindowState = vbMaximized
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    rsCustomer.DatabaseName = m_strPrograming
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsCountry = m_dbPrograming.OpenRecordset("Country")
    Set rsPayment = m_dbPrograming.OpenRecordset("Payment")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmCustomer")
    m_iFormNo = 5
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Error$, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Resize()
    ResizeForm Me
    LoadBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsCustomer.Recordset.Close
    rsCountry.Close
    rsPayment.Close
    rsLanguage.Close
    rsUser.Close
    m_iFormNo = 0
    DisableButtons 1
    With frmMDI.Toolbar1
        .Buttons(13).Enabled = False
        .Buttons(14).Enabled = False
        .Buttons(15).Enabled = False
        .Buttons(16).Enabled = False
    End With
    Set frmCustomer = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsCustomer.Recordset.Bookmark = bookmarksPro(List1.ItemData(List1.ListIndex))
End Sub


Private Sub Text1_LostFocus(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
        If bNewRecord Then
            With rsCustomer.Recordset
                .Fields("CustomerName") = Trim(Text1(0).Text)
                .Update
                LoadList1
                .Bookmark = .LastModified
                bNewRecord = False
            End With
        End If
    Case Else
    End Select
End Sub


