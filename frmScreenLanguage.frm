VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmScreenLanguage 
   BackColor       =   &H00404040&
   Caption         =   "Screen Language"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   10440
   Begin TabDlg.SSTab Tab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   33
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "Invoice Text"
      TabPicture(0)   =   "frmScreenLanguage.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Data1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Live Update"
      TabPicture(1)   =   "frmScreenLanguage.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Data2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "User Record"
      TabPicture(2)   =   "frmScreenLanguage.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Data3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Payment Due DateText"
      TabPicture(3)   =   "frmScreenLanguage.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid1(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Data4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Code Snippets"
      TabPicture(4)   =   "frmScreenLanguage.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid1(4)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Data5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "Code Type"
      TabPicture(5)   =   "frmScreenLanguage.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Data6"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Country"
      TabPicture(6)   =   "frmScreenLanguage.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Grid1(6)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Data7"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "Customers"
      TabPicture(7)   =   "frmScreenLanguage.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Grid1(7)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Data8"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).ControlCount=   2
      TabCaption(8)   =   "Database Password"
      TabPicture(8)   =   "frmScreenLanguage.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Grid1(8)"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Data9"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      TabCaption(9)   =   "Database Print"
      TabPicture(9)   =   "frmScreenLanguage.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Grid1(9)"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "Data10"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).ControlCount=   2
      TabCaption(10)  =   "Email"
      TabPicture(10)  =   "frmScreenLanguage.frx":0118
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Grid1(10)"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).Control(1)=   "Data11"
      Tab(10).Control(1).Enabled=   0   'False
      Tab(10).ControlCount=   2
      TabCaption(11)  =   "Invoice"
      TabPicture(11)  =   "frmScreenLanguage.frx":0134
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Grid1(11)"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "Data12"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).ControlCount=   2
      TabCaption(12)  =   "Licence"
      TabPicture(12)  =   "frmScreenLanguage.frx":0150
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Grid1(12)"
      Tab(12).Control(0).Enabled=   0   'False
      Tab(12).Control(1)=   "Data13"
      Tab(12).Control(1).Enabled=   0   'False
      Tab(12).ControlCount=   2
      TabCaption(13)  =   "Compact Database"
      TabPicture(13)  =   "frmScreenLanguage.frx":016C
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Grid1(13)"
      Tab(13).Control(0).Enabled=   0   'False
      Tab(13).Control(1)=   "Data14"
      Tab(13).Control(1).Enabled=   0   'False
      Tab(13).ControlCount=   2
      TabCaption(14)  =   "Mass Mail"
      TabPicture(14)  =   "frmScreenLanguage.frx":0188
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Grid1(14)"
      Tab(14).Control(0).Enabled=   0   'False
      Tab(14).Control(1)=   "Data15"
      Tab(14).Control(1).Enabled=   0   'False
      Tab(14).ControlCount=   2
      TabCaption(15)  =   "Search Code"
      TabPicture(15)  =   "frmScreenLanguage.frx":01A4
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Grid1(15)"
      Tab(15).Control(0).Enabled=   0   'False
      Tab(15).Control(1)=   "Data16"
      Tab(15).Control(1).Enabled=   0   'False
      Tab(15).ControlCount=   2
      TabCaption(16)  =   "MDI"
      TabPicture(16)  =   "frmScreenLanguage.frx":01C0
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Grid1(16)"
      Tab(16).Control(0).Enabled=   0   'False
      Tab(16).Control(1)=   "Data17"
      Tab(16).Control(1).Enabled=   0   'False
      Tab(16).ControlCount=   2
      TabCaption(17)  =   "Passwords"
      TabPicture(17)  =   "frmScreenLanguage.frx":01DC
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Grid1(17)"
      Tab(17).Control(0).Enabled=   0   'False
      Tab(17).Control(1)=   "Data18"
      Tab(17).Control(1).Enabled=   0   'False
      Tab(17).ControlCount=   2
      TabCaption(18)  =   "Payments"
      TabPicture(18)  =   "frmScreenLanguage.frx":01F8
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "Grid1(18)"
      Tab(18).Control(0).Enabled=   0   'False
      Tab(18).Control(1)=   "Data19"
      Tab(18).Control(1).Enabled=   0   'False
      Tab(18).ControlCount=   2
      TabCaption(19)  =   "Print Database"
      TabPicture(19)  =   "frmScreenLanguage.frx":0214
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Grid1(19)"
      Tab(19).Control(0).Enabled=   0   'False
      Tab(19).Control(1)=   "Data20"
      Tab(19).Control(1).Enabled=   0   'False
      Tab(19).ControlCount=   2
      TabCaption(20)  =   "Programming"
      TabPicture(20)  =   "frmScreenLanguage.frx":0230
      Tab(20).ControlEnabled=   0   'False
      Tab(20).Control(0)=   "Grid1(20)"
      Tab(20).Control(0).Enabled=   0   'False
      Tab(20).Control(1)=   "Data21"
      Tab(20).Control(1).Enabled=   0   'False
      Tab(20).ControlCount=   2
      TabCaption(21)  =   "Projects"
      TabPicture(21)  =   "frmScreenLanguage.frx":024C
      Tab(21).ControlEnabled=   0   'False
      Tab(21).Control(0)=   "Grid1(21)"
      Tab(21).Control(0).Enabled=   0   'False
      Tab(21).Control(1)=   "Data22"
      Tab(21).Control(1).Enabled=   0   'False
      Tab(21).ControlCount=   2
      TabCaption(22)  =   "Write To Me"
      TabPicture(22)  =   "frmScreenLanguage.frx":0268
      Tab(22).ControlEnabled=   0   'False
      Tab(22).Control(0)=   "Grid1(22)"
      Tab(22).Control(0).Enabled=   0   'False
      Tab(22).Control(1)=   "Data23"
      Tab(22).Control(1).Enabled=   0   'False
      Tab(22).ControlCount=   2
      TabCaption(23)  =   "Registration"
      TabPicture(23)  =   "frmScreenLanguage.frx":0284
      Tab(23).ControlEnabled=   0   'False
      Tab(23).Control(0)=   "Grid1(23)"
      Tab(23).Control(0).Enabled=   0   'False
      Tab(23).Control(1)=   "Data24"
      Tab(23).Control(1).Enabled=   0   'False
      Tab(23).ControlCount=   2
      TabCaption(24)  =   "Computer Password"
      TabPicture(24)  =   "frmScreenLanguage.frx":02A0
      Tab(24).ControlEnabled=   0   'False
      Tab(24).Control(0)=   "Grid1(24)"
      Tab(24).Control(0).Enabled=   0   'False
      Tab(24).Control(1)=   "Data25"
      Tab(24).Control(1).Enabled=   0   'False
      Tab(24).ControlCount=   2
      TabCaption(25)  =   "Key Ascii"
      TabPicture(25)  =   "frmScreenLanguage.frx":02BC
      Tab(25).ControlEnabled=   0   'False
      Tab(25).Control(0)=   "Grid1(25)"
      Tab(25).Control(0).Enabled=   0   'False
      Tab(25).Control(1)=   "Data26"
      Tab(25).Control(1).Enabled=   0   'False
      Tab(25).ControlCount=   2
      TabCaption(26)  =   "Program Supplier"
      TabPicture(26)  =   "frmScreenLanguage.frx":02D8
      Tab(26).ControlEnabled=   0   'False
      Tab(26).Control(0)=   "Grid1(26)"
      Tab(26).Control(0).Enabled=   0   'False
      Tab(26).Control(1)=   "Data27"
      Tab(26).Control(1).Enabled=   0   'False
      Tab(26).ControlCount=   2
      TabCaption(27)  =   "Print Customers"
      TabPicture(27)  =   "frmScreenLanguage.frx":02F4
      Tab(27).ControlEnabled=   0   'False
      Tab(27).Control(0)=   "Grid1(27)"
      Tab(27).Control(0).Enabled=   0   'False
      Tab(27).Control(1)=   "Data28"
      Tab(27).Control(1).Enabled=   0   'False
      Tab(27).ControlCount=   2
      TabCaption(28)  =   "Print Project Times"
      TabPicture(28)  =   "frmScreenLanguage.frx":0310
      Tab(28).ControlEnabled=   0   'False
      Tab(28).Control(0)=   "Grid1(28)"
      Tab(28).Control(0).Enabled=   0   'False
      Tab(28).Control(1)=   "Data29"
      Tab(28).Control(1).Enabled=   0   'False
      Tab(28).ControlCount=   2
      TabCaption(29)  =   "Prosjekt Statistic"
      TabPicture(29)  =   "frmScreenLanguage.frx":032C
      Tab(29).ControlEnabled=   0   'False
      Tab(29).Control(0)=   "Grid1(29)"
      Tab(29).Control(0).Enabled=   0   'False
      Tab(29).Control(1)=   "Data30"
      Tab(29).Control(1).Enabled=   0   'False
      Tab(29).ControlCount=   2
      TabCaption(30)  =   "Mail Code Snippets"
      TabPicture(30)  =   "frmScreenLanguage.frx":0348
      Tab(30).ControlEnabled=   0   'False
      Tab(30).Control(0)=   "Grid1(30)"
      Tab(30).Control(0).Enabled=   0   'False
      Tab(30).Control(1)=   "Data31"
      Tab(30).Control(1).Enabled=   0   'False
      Tab(30).ControlCount=   2
      TabCaption(31)  =   "Code Statistic"
      TabPicture(31)  =   "frmScreenLanguage.frx":0364
      Tab(31).ControlEnabled=   0   'False
      Tab(31).Control(0)=   "Grid1(31)"
      Tab(31).Control(0).Enabled=   0   'False
      Tab(31).Control(1)=   "Data32"
      Tab(31).Control(1).Enabled=   0   'False
      Tab(31).ControlCount=   2
      TabCaption(32)  =   "API "
      TabPicture(32)  =   "frmScreenLanguage.frx":0380
      Tab(32).ControlEnabled=   0   'False
      Tab(32).Control(0)=   "Grid1(32)"
      Tab(32).Control(0).Enabled=   0   'False
      Tab(32).Control(1)=   "Data33"
      Tab(32).Control(1).Enabled=   0   'False
      Tab(32).ControlCount=   2
      Begin VB.Data Data33 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Programing\ProgrammersHelper\CodeLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -69720
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmAPI"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data32 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programing\Source Code\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -68400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCodeStatistic"
         Top             =   2820
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data31 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programing\Source Code\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -68400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmSnippetMail"
         Top             =   2940
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data30 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmStats"
         Top             =   3180
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data29 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmWriteProgRep"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data28 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPrintCustomer"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data27 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmSupplier"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data26 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmKeyAscii"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data25 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPasword"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data24 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmRegistration"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data23 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmWriteToMe"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data22 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmProjects"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data21 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmProgramming"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data20 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPrintDB"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data19 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPayments"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data18 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmPasswords"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data17 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmMDI"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data16 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmSearchCode"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data15 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmMassMail"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data14 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmMaint"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data13 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmLicence"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data12 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmInvoice"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data11 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmEmail"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data10 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmDatabasePrint"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data9 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmDatabasePassword"
         Top             =   3120
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data8 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCustomer"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data7 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -66840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCountry"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data6 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCodeType"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data5 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -67200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmCodeSnippets"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data4 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -70080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DueDateText"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data3 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -68160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmUser"
         Top             =   2880
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data2 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -68400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "frmLiveUpdate"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   "D:\Programmering\Programmering\ProgramLang.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   7440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "InvoiceText"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":039C
         Height          =   3855
         Index           =   0
         Left            =   240
         OleObjectBlob   =   "frmScreenLanguage.frx":03B0
         TabIndex        =   1
         Top             =   2880
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":0D84
         Height          =   3735
         Index           =   1
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":0D98
         TabIndex        =   2
         Top             =   3000
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":176C
         Height          =   3735
         Index           =   2
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":1780
         TabIndex        =   3
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":2154
         Height          =   3855
         Index           =   3
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":2168
         TabIndex        =   4
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":2B3C
         Height          =   3855
         Index           =   4
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":2B50
         TabIndex        =   5
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":3524
         Height          =   3855
         Index           =   5
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":3538
         TabIndex        =   6
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":3F0C
         Height          =   3615
         Index           =   6
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":3F20
         TabIndex        =   7
         Top             =   3000
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":48F4
         Height          =   3735
         Index           =   7
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":4908
         TabIndex        =   8
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":52DC
         Height          =   3855
         Index           =   8
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":52F0
         TabIndex        =   9
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":5CC4
         Height          =   3735
         Index           =   9
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":5CD9
         TabIndex        =   10
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":66AD
         Height          =   3735
         Index           =   10
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":66C2
         TabIndex        =   11
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":7097
         Height          =   3855
         Index           =   11
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":70AC
         TabIndex        =   12
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":7A81
         Height          =   3615
         Index           =   12
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":7A96
         TabIndex        =   13
         Top             =   3000
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":846B
         Height          =   3735
         Index           =   13
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":8480
         TabIndex        =   14
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":8E55
         Height          =   3855
         Index           =   14
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":8E6A
         TabIndex        =   15
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":983F
         Height          =   3855
         Index           =   15
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":9854
         TabIndex        =   16
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":A229
         Height          =   3855
         Index           =   16
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":A23E
         TabIndex        =   17
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":AC13
         Height          =   3855
         Index           =   17
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":AC28
         TabIndex        =   18
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":B5FD
         Height          =   3615
         Index           =   18
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":B612
         TabIndex        =   19
         Top             =   3000
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":BFE7
         Height          =   3735
         Index           =   19
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":BFFC
         TabIndex        =   20
         Top             =   3000
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":C9D1
         Height          =   3735
         Index           =   20
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":C9E6
         TabIndex        =   21
         Top             =   3000
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":D3BB
         Height          =   3735
         Index           =   21
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":D3D0
         TabIndex        =   22
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":DDA5
         Height          =   3855
         Index           =   22
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":DDBA
         TabIndex        =   23
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":E78F
         Height          =   3855
         Index           =   23
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":E7A4
         TabIndex        =   24
         Top             =   2880
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":F179
         Height          =   3855
         Index           =   24
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":F18E
         TabIndex        =   25
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":FB63
         Height          =   3735
         Index           =   25
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":FB78
         TabIndex        =   26
         Top             =   2880
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":1054D
         Height          =   3615
         Index           =   26
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":10562
         TabIndex        =   27
         Top             =   3000
         Width           =   8415
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":10F37
         Height          =   3735
         Index           =   27
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":10F4C
         TabIndex        =   28
         Top             =   3000
         Width           =   8295
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":11921
         Height          =   3615
         Index           =   28
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":11936
         TabIndex        =   29
         Top             =   3000
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":1230B
         Height          =   3735
         Index           =   29
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":12320
         TabIndex        =   30
         Top             =   3000
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":12CF5
         Height          =   3615
         Index           =   30
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":12D0A
         TabIndex        =   31
         Top             =   3120
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":136DF
         Height          =   3615
         Index           =   31
         Left            =   -74760
         OleObjectBlob   =   "frmScreenLanguage.frx":136F4
         TabIndex        =   32
         Top             =   3000
         Width           =   8175
      End
      Begin MSDBGrid.DBGrid Grid1 
         Bindings        =   "frmScreenLanguage.frx":140C9
         Height          =   3855
         Index           =   32
         Left            =   -74880
         OleObjectBlob   =   "frmScreenLanguage.frx":140DE
         TabIndex        =   33
         Top             =   2880
         Width           =   8295
      End
   End
End
Attribute VB_Name = "frmScreenLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Update
    End With
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    Data1.Refresh
    Data2.Refresh
    Data3.Refresh
    Data4.Refresh
    Data5.Refresh
    Data6.Refresh
    Data7.Refresh
    Data8.Refresh
    Data9.Refresh
    Data10.Refresh
    Data11.Refresh
    Data12.Refresh
    Data13.Refresh
    Data14.Refresh
    Data15.Refresh
    Data16.Refresh
    Data17.Refresh
    Data18.Refresh
    Data19.Refresh
    Data20.Refresh
    Data21.Refresh
    Data22.Refresh
    Data23.Refresh
    Data24.Refresh
    Data25.Refresh
    Data26.Refresh
    Data27.Refresh
    Data28.Refresh
    Data29.Refresh
    Data30.Refresh
    Data31.Refresh
    Data32.Refresh
    Data33.Refresh
    Me.WindowState = vbMaximized
End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Data1.DatabaseName = m_strProgramLng
    Data2.DatabaseName = m_strProgramLng
    Data3.DatabaseName = m_strProgramLng
    Data4.DatabaseName = m_strProgramLng
    Data5.DatabaseName = m_strProgramLng
    Data6.DatabaseName = m_strProgramLng
    Data7.DatabaseName = m_strProgramLng
    Data8.DatabaseName = m_strProgramLng
    Data9.DatabaseName = m_strProgramLng
    Data10.DatabaseName = m_strProgramLng
    Data11.DatabaseName = m_strProgramLng
    Data12.DatabaseName = m_strProgramLng
    Data13.DatabaseName = m_strProgramLng
    Data14.DatabaseName = m_strProgramLng
    Data15.DatabaseName = m_strProgramLng
    Data16.DatabaseName = m_strProgramLng
    Data17.DatabaseName = m_strProgramLng
    Data18.DatabaseName = m_strProgramLng
    Data19.DatabaseName = m_strProgramLng
    Data20.DatabaseName = m_strProgramLng
    Data21.DatabaseName = m_strProgramLng
    Data22.DatabaseName = m_strProgramLng
    Data23.DatabaseName = m_strProgramLng
    Data24.DatabaseName = m_strProgramLng
    Data25.DatabaseName = m_strProgramLng
    Data26.DatabaseName = m_strProgramLng
    Data27.DatabaseName = m_strProgramLng
    Data28.DatabaseName = m_strProgramLng
    Data29.DatabaseName = m_strProgramLng
    Data30.DatabaseName = m_strProgramLng
    Data31.DatabaseName = m_strProgramLng
    Data32.DatabaseName = m_strProgramLng
    Data33.DatabaseName = m_strProgramLng
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmScreenLanguage")
    ReadText
    m_iFormNo = 20
End Sub

Private Sub Form_Resize()
    ResizeForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Data1.Recordset.Close
    Data2.Recordset.Close
    Data3.Recordset.Close
    Data4.Recordset.Close
    Data5.Recordset.Close
    Data6.Recordset.Close
    Data7.Recordset.Close
    Data8.Recordset.Close
    Data9.Recordset.Close
    Data10.Recordset.Close
    Data11.Recordset.Close
    Data12.Recordset.Close
    Data13.Recordset.Close
    Data14.Recordset.Close
    Data15.Recordset.Close
    Data16.Recordset.Close
    Data17.Recordset.Close
    Data18.Recordset.Close
    Data19.Recordset.Close
    Data20.Recordset.Close
    Data21.Recordset.Close
    Data22.Recordset.Close
    Data23.Recordset.Close
    Data24.Recordset.Close
    Data25.Recordset.Close
    Data26.Recordset.Close
    Data27.Recordset.Close
    Data28.Recordset.Close
    Data29.Recordset.Close
    Data30.Recordset.Close
    Data31.Recordset.Close
    Data32.Recordset.Close
    Data33.Recordset.Close
    rsLanguage.Close
    m_iFormNo = 0
    Set frmScreenLanguage = Nothing
End Sub


