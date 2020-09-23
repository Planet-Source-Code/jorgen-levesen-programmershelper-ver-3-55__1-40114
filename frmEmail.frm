VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEmail 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "E-mail"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   345
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1800
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1440
      Width           =   6255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   810
      Left            =   2400
      TabIndex        =   12
      Top             =   480
      Width           =   5655
   End
   Begin VB.ComboBox cboAdr 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   2400
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton btnMailTo 
      Caption         =   "To (Outlook address).."
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.CommandButton btnExit 
      Height          =   495
      Left            =   6480
      Picture         =   "frmEmail.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton btnSend 
      Height          =   495
      Left            =   2160
      Picture         =   "frmEmail.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Send this Mail"
      Top             =   6360
      Width           =   4215
   End
   Begin VB.CommandButton btnAttachment 
      Height          =   495
      Left            =   240
      Picture         =   "frmEmail.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Attachments"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Attachments"
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   8400
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   1665
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   1980
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   2175
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Dbl Click to remove"
         Top             =   4440
         Width           =   2535
      End
   End
   Begin RichTextLib.RichTextBox RichText1 
      Height          =   3855
      Left            =   360
      TabIndex        =   15
      Top             =   2280
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6800
      _Version        =   393217
      BackColor       =   16777152
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmEmail.frx":09E0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   17
      Top             =   1920
      Width           =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   8160
      X2              =   8160
      Y1              =   240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   240
      X2              =   8160
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1440
      X2              =   8160
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   240
      X2              =   8160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "To :"
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
      Index           =   1
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ol As Object
Dim olns As Object
Dim objFolder As Object
Dim objAllContacts As Object
Dim Contact As Object
Dim rsMassMail As Recordset
Dim rsUser As Recordset
Dim rsLanguage As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
    For i = 0 To 2
        Label1(i).ForeColor = rsUser.Fields("LabelColor")
    Next
    For i = 0 To 7
        Line1(i).BorderColor = rsUser.Fields("FrameColor")
    Next
End Sub

Private Sub LoadContacts()
    On Error GoTo errLoadContact
    If IsOutlookPresent Then
        ' Set the application object
        Set ol = New Outlook.Application
        ' Set the namespace object
        Set olns = ol.GetNamespace("MAPI")
        ' Set the default Contacts folder
        Set objFolder = olns.GetDefaultFolder(olFolderContacts)
        ' Set objAllContacts = the collection of all contacts
        Set objAllContacts = objFolder.Items
        
        cboAdr.Clear
        ' Loop through each contact
        For Each Contact In objAllContacts
           'Display the Fullname field for the contact
           cboAdr.AddItem Contact.FullName
        Next
    End If
    Exit Sub
    
errLoadContact:
    Beep
    MsgBox Err.Description, vbCritical, "Load Outlook Contacts"
    Err.Clear
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
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("label1(2)")) Then
                    .Fields("Label1(2)") = Label1(2).Caption
                Else
                    Label1(2).Caption = .Fields("Label1(2)")
                End If
                If IsNull(.Fields("btnSend")) Then
                    .Fields("btnSend") = btnSend.ToolTipText
                Else
                    btnSend.ToolTipText = .Fields("btnSend")
                End If
                If IsNull(.Fields("btnAttachment")) Then
                    .Fields("btnAttachment") = btnAttachment.ToolTipText
                Else
                    btnAttachment.ToolTipText = .Fields("btnAttachment")
                End If
                If IsNull(.Fields("btnExit")) Then
                    .Fields("btnExit") = btnExit.ToolTipText
                Else
                    btnExit.ToolTipText = .Fields("btnExit")
                End If
                If IsNull(.Fields("btnMailTo")) Then
                    .Fields("btnMailTo") = btnMailTo.Caption
                Else
                    btnMailTo.Caption = .Fields("btnMailTo")
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
        .Fields("label1(2)") = Label1(2).Caption
        .Fields("btnSend") = btnSend.ToolTipText
        .Fields("btnAttachment") = btnAttachment.ToolTipText
        .Fields("btnExit") = btnExit.ToolTipText
        .Fields("btnMailTo") = btnMailTo.Caption
        .Update
    End With
End Sub

Private Sub btnAttachment_Click()
    List2.Clear
    Me.Width = Me.Width + 2855
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnMailTo_Click()
    If IsOutlookPresent Then cboAdr.Visible = True
End Sub
Private Sub btnSend_Click()
Dim clsOutlook As cOutlookSendMail, strMailTo As String
    
    If IsOutlookPresent Then
        On Error Resume Next
        strMailTo = ""
        Set clsOutlook = New cOutlookSendMail
        With clsOutlook
            .StartOutlook
            .CreateNewMail
            If List1.Visible Then
                For i = 0 To List1.ListCount - 1
                    .Recipient_TO = List1.List(i)
                    If strMailTo = "" Then
                        strMailTo = strMailTo & List1.List(i)
                    Else
                        strMailTo = strMailTo & ";" & List1.List(i)
                    End If
                Next
            Else    'wish to send a complete code snippet
                If Len(Text2.Text) <> 0 Then
                    .Recipient_TO = Text2.Text
                    strMailTo = strMailTo & Text2.Text
                Else    'forgot to write the email address ?
                    Text2.SetFocus
                    Exit Sub
                End If
            End If
            .Subject = Text1.Text
            .Body = RichText1.Text
            If List2.ListCount <> 0 Then
                For i = 0 To List2.ListCount - 1
                    .Attachment List2.List(i)
                Next
            End If
            .SendMail
            .CloseOutlook
        End With
        Set clsOutlook = Nothing
        
        'write logfile
        With rsMassMail
            .AddNew
            .Fields("MailDate") = Format(Now, "dd.mm.yyyy")
            .Fields("MailTime") = Format(Now, "hh:mm")
            .Fields("MailTo") = strMailTo
            .Fields("Mailregarding") = Text1.Text
            .Fields("MailContents") = RichText1.Text
            .Update
        End With
        MsgBox "Mail is send !"
    Else    'send this mail via WEB (vbSendMail)
        If IsWebConnected() Then
            If Not IsNull(rsUser.Fields("CompanyMailServerName")) Then
                On Error Resume Next
                Dim poSendMail As vbSendMail.clsSendMail
                Set poSendMail = New clsSendMail
                With poSendMail
                    .SMTPHost = Trim(rsUser.Fields("CompanyMailServerName"))
                    .from = rsUser.Fields("CompanyEMail")
                    .FromDisplayName = rsUser.Fields("CompanyName")
                    .Message = RichText1.Text
            If List1.Visible Then
                For i = 0 To List1.ListCount - 1
                    .Recipient = List1.List(i)
                Next
            Else    'wish to send a complete code snippet
                If Len(Text2.Text) <> 0 Then
                    .Recipient = Text2.Text
                Else    'forgot to write the email address ?
                    Text2.SetFocus
                    Exit Sub
                End If
            End If
                    .Subject = Text1.Text
                    .Send
                End With
            End If
            Set poSendMail = Nothing
            MsgBox "Mail is send !"
            Unload Me
        End If
    End If
    Unload Me
End Sub

Private Sub cboAdr_Click()
Dim sFilter As String
    sFilter = "[FullName] = """ & cboAdr.List(cboAdr.ListIndex) & """"
    Set Contact = objAllContacts.Find(sFilter)
    If Contact Is Nothing Then ' the Find failed
       MsgBox "Not Found"
    Else
        If Contact.Email1Address <> "" Then
            Text2.Text = Contact.Email1Address
        End If
    End If
    cboAdr.Visible = False
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    List2.AddItem Dir1.List(Dir1.ListIndex) & "\" & File1.List(File1.ListIndex)
End Sub

Private Sub File1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ItemHeight As Long
   Dim NewIndex As Long
   Static OldIndex As Long

   With File1
      ItemHeight = SendMessage(.hWnd, LB_GETITEMHEIGHT, 0, ByVal 0&)
      ItemHeight = .Parent.ScaleY(ItemHeight, vbPixels, vbTwips)
      NewIndex = .TopIndex + (Y \ ItemHeight)
      If NewIndex <> OldIndex Then
         If NewIndex < .ListCount Then
            .ToolTipText = .List(NewIndex)
         Else
            .ToolTipText = vbNullString
         End If
         OldIndex = NewIndex
      End If
   End With
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    DisableButtons 1
    ReadText
    LoadBackground
    LoadContacts
    If Not IsNull(rsUser.Fields("ElectronicSign")) Then
        RichText1.Text = RichText1.Text & vbCrLf & vbCrLf & _
                        vbCrLf & vbCrLf & _
                        Format(rsUser.Fields("ElectronicSign"))
    End If
    LoadBackground
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Set rsMassMail = m_dbPrograming.OpenRecordset("MassMail")
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmEmail")
    m_iFormNo = 7
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "LoadForm"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsMassMail.Close
    rsUser.Close
    rsLanguage.Close
    m_iFormNo = 0
    DisableButtons 2
    Set frmEmail = Nothing
End Sub
Private Sub List2_DblClick()
    List2.RemoveItem (List2.ListIndex)
End Sub
