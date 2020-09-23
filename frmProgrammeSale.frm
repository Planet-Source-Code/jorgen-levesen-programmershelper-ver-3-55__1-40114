VERSION 5.00
Begin VB.Form frmProgrammeSale 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programme Sale"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5895
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton btnExit 
         Height          =   615
         Left            =   2520
         Picture         =   "frmProgrammeSale.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton btnDelete 
         Height          =   615
         Left            =   1800
         Picture         =   "frmProgrammeSale.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5040
         Width           =   615
      End
      Begin VB.CommandButton btnNew 
         Height          =   615
         Left            =   1080
         Picture         =   "frmProgrammeSale.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5040
         Width           =   615
      End
      Begin VB.ComboBox cmbCustomer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "InstalationKey"
         DataSource      =   "rsProjectSale"
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   12
         Top             =   3000
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "InstalationDate"
         DataSource      =   "rsProjectSale"
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   8
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "PurchasePrice"
         DataSource      =   "rsProjectSale"
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         DataField       =   "NumberOfProgramme"
         DataSource      =   "rsProjectSale"
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   5
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
      Begin VB.Data rsProjectSale 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Amiprog\New Programmes\Programming\Programming.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "ProjectSale"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   5490
         Left            =   3360
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Installation Key:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Purchase Date:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Purchase price:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Quant. purchased:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   5490
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Projects"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmProgrammeSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bNewRecord As Boolean
Dim bookmarksPro() As Variant
Dim dbTemp As Database
Dim rsProjects As Recordset
Dim rsCustomer As Recordset
Private Sub LoadList1()
    On Error Resume Next
    With rsProjects
        .MoveLast
        .MoveFirst
        ReDim bookmarksPro(.RecordCount)
        Do While Not .EOF
            List1.AddItem .Fields("ProjectID")
            List1.ItemData(List1.NewIndex) = List1.ListCount - 1
            bookmarksPro(List1.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub


Private Sub btnDelete_Click()
    On Error Resume Next
    rsProjectSale.Recordset.Delete
    LoadList2
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnNew_Click()
    rsCustomer.Recordset.AddNew
    bNewRecord = True
    cmbCustomer.SetFocus
End Sub

Private Sub cmbCustomer_Change()

End Sub

Private Sub cmbCustomer_LostFocus()
        If bNewRecord Then
            With rsProjectSale.Recordset
                .Fields("ProjectID") = rsProjects.Fields("ProjectID")
                .Fields("CustomerID") = rsCustomer.Fields("AutoLine")
                .Update
                LoadList1
                .Bookmark = .LastModified
                bNewRecord = False
            End With
        End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rsProjectSale.Refresh
End Sub

Private Sub Form_Load()
    On Error GoTo errForm_Load
    sPath = App.Path & "\Programming.mdb"
    rsProjectSale.DatabaseName = sPath
    Set dbTemp = OpenDatabase(sPath)
    Set rsProjects = dbTemp.OpenRecordset("Projects")
    Set rsCustomer = dbTemp.OpenRecordset("Customer")
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Resume errForm_Load2
errForm_Load2:
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsProjectSale.Recordset.Close
    rsProjects.Close
    rsCustomer.Close
    dbTemp.Close
    Set frmProgrammeSale = Nothing
End Sub

Private Sub List1_Click()
    On Error Resume Next
    rsProjects.Bookmark = bookmarksPro(List1.ItemData(List1.ListIndex))
End Sub

Private Sub Text1_Change(Index As Integer)

End Sub


