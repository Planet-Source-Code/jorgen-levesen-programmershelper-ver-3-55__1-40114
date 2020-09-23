VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUser 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab Tab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   4210752
      TabCaption(0)   =   "Page 1"
      TabPicture(0)   =   "frmUser.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Page 2"
      TabPicture(1)   =   "frmUser.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Page 3"
      TabPicture(2)   =   "frmUser.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Backgrounds / Colors"
      TabPicture(3)   =   "frmUser.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "BackgroundPic"
      Tab(3).Control(1)=   "Line1(0)"
      Tab(3).Control(2)=   "Line1(1)"
      Tab(3).Control(3)=   "Line1(2)"
      Tab(3).Control(4)=   "Line1(3)"
      Tab(3).Control(5)=   "Label2"
      Tab(3).Control(6)=   "btnBackground"
      Tab(3).Control(7)=   "btnFrame"
      Tab(3).Control(8)=   "btnLabel"
      Tab(3).ControlCount=   9
      Begin VB.CommandButton btnLabel 
         Caption         =   "&Label color"
         Height          =   375
         Left            =   -74520
         TabIndex        =   68
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton btnFrame 
         Caption         =   "&Frame color"
         Height          =   375
         Left            =   -72840
         TabIndex        =   67
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton btnBackground 
         Caption         =   "&Background"
         Height          =   375
         Left            =   -71160
         TabIndex        =   66
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Frame Frame5 
         Height          =   6375
         Left            =   -74760
         TabIndex        =   45
         Top             =   480
         Width           =   5415
         Begin VB.Frame Frame7 
            Caption         =   "Mail Electronic Signature:"
            Height          =   2895
            Left            =   120
            TabIndex        =   59
            Top             =   3360
            Width           =   5055
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
               DataField       =   "ElectronicSign"
               DataSource      =   "rsUser"
               Height          =   2415
               Index           =   14
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   60
               Top             =   360
               Width           =   4695
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Making new Language Recordsets"
            Height          =   855
            Left            =   120
            TabIndex        =   55
            Top             =   2400
            Visible         =   0   'False
            Width           =   5055
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   255
               Left            =   240
               TabIndex        =   56
               Top             =   360
               Width           =   4575
               _ExtentX        =   8070
               _ExtentY        =   450
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H8000000A&
            Caption         =   "Company Logo"
            ForeColor       =   &H00000000&
            Height          =   2055
            Left            =   2520
            TabIndex        =   48
            Top             =   240
            Width           =   2655
            Begin VB.CommandButton btnPaste 
               Height          =   375
               Left            =   240
               Picture         =   "frmUser.frx":0070
               Style           =   1  'Graphical
               TabIndex        =   52
               ToolTipText     =   "Paste picture from the Clipboard"
               Top             =   360
               Width           =   495
            End
            Begin VB.CommandButton btnCopy 
               Height          =   375
               Left            =   240
               Picture         =   "frmUser.frx":0732
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Copythe picture to the Clipboard"
               Top             =   720
               Width           =   495
            End
            Begin VB.CommandButton btnDelete 
               Height          =   375
               Left            =   240
               Picture         =   "frmUser.frx":0DF4
               Style           =   1  'Graphical
               TabIndex        =   50
               ToolTipText     =   "Delete this Picture"
               Top             =   1080
               Width           =   495
            End
            Begin VB.CommandButton btnReadFromFile 
               Height          =   375
               Left            =   240
               Picture         =   "frmUser.frx":0F3E
               Style           =   1  'Graphical
               TabIndex        =   49
               ToolTipText     =   "Reada picturefrom a disk file"
               Top             =   1440
               Width           =   495
            End
            Begin VB.Image Image1 
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               DataField       =   "CompanyLogo"
               DataSource      =   "rsUser"
               Height          =   1455
               Left            =   840
               Stretch         =   -1  'True
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H8000000A&
            Caption         =   "Screen Language"
            ForeColor       =   &H00000000&
            Height          =   2055
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   2175
            Begin VB.ComboBox cmbLanguage 
               BackColor       =   &H00FFFFC0&
               DataField       =   "LanguageOnScreen"
               DataSource      =   "rsUser"
               Height          =   315
               Left            =   240
               TabIndex        =   47
               Top             =   960
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000A&
         Height          =   6495
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   5535
         Begin VB.ComboBox cboPreferedLanguage 
            BackColor       =   &H00FFFFC0&
            DataField       =   "PrefferedLanguage"
            DataSource      =   "rsUser"
            Height          =   315
            Left            =   2760
            TabIndex        =   65
            Top             =   3480
            Width           =   2655
         End
         Begin VB.CheckBox Check3 
            DataField       =   "OwnSnippetAsDefault"
            DataSource      =   "rsUser"
            Height          =   255
            Left            =   2760
            TabIndex        =   62
            Top             =   5760
            Width           =   255
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            DataField       =   "PrintUseWord"
            DataSource      =   "rsUser"
            Height          =   255
            Left            =   2760
            TabIndex        =   53
            Top             =   5280
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            DataField       =   "CopyWithAppro"
            DataSource      =   "rsUser"
            Height          =   375
            Left            =   2760
            TabIndex        =   43
            ToolTipText     =   "Copy as: ""fieldName"" = Selected, or rs!fieldName = Blank"
            Top             =   4800
            Width           =   255
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "FaxDirectory"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   19
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   41
            Top             =   3960
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "LetterDirectory"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   14
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   40
            Top             =   4320
            Width           =   2655
         End
         Begin MSComDlg.CommonDialog Cmd1 
            Left            =   120
            Top             =   240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyPaymentDays"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   17
            Left            =   2760
            TabIndex        =   37
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyLatePayment"
            DataSource      =   "rsUser"
            Height          =   1965
            Index           =   18
            Left            =   2760
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyPayment"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   16
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   33
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyVAT"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   15
            Left            =   2760
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Prefered code Language:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   25
            Left            =   360
            TabIndex        =   64
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Letter Directory Folder:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   360
            TabIndex        =   63
            Top             =   3960
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Use own Code Snippet database as default:"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   24
            Left            =   120
            TabIndex        =   61
            Top             =   5760
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Print using MS Word:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   54
            Top             =   5280
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Copy field names with aprostroph:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   44
            Top             =   4920
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax Directory Folder:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   21
            Left            =   360
            TabIndex        =   42
            Top             =   4320
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   3480
            TabIndex        =   39
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment days:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   38
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Late payment text:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   36
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment term:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   34
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   15
            Left            =   3480
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Country VAT:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5655
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyMailServerName"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   20
            Left            =   2400
            TabIndex        =   57
            Top             =   4800
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyName"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   0
            Left            =   2400
            TabIndex        =   15
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyAddress1"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   1
            Left            =   2400
            TabIndex        =   14
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyAddress2"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   2
            Left            =   2400
            TabIndex        =   13
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyZip"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   3
            Left            =   2400
            TabIndex        =   12
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyTown"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   4
            Left            =   2400
            TabIndex        =   11
            Top             =   1920
            Width           =   2415
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyCountry"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   5
            Left            =   2400
            TabIndex        =   10
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyPrefixPhone"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   6
            Left            =   2400
            TabIndex        =   9
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyPhoneNo"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   7
            Left            =   2400
            TabIndex        =   8
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyFaxNo"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   8
            Left            =   2400
            TabIndex        =   7
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyEMail"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   9
            Left            =   2400
            TabIndex        =   6
            Top             =   4080
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyURL"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   10
            Left            =   2400
            TabIndex        =   5
            Top             =   4440
            Width           =   3135
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyHourWage"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   11
            Left            =   2400
            TabIndex        =   4
            Top             =   5640
            Width           =   975
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "CompanyHourWageDim"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   12
            Left            =   4680
            TabIndex        =   3
            Top             =   5640
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            DataField       =   "NextInvoiceNo"
            DataSource      =   "rsUser"
            Height          =   285
            Index           =   13
            Left            =   2400
            TabIndex        =   2
            Top             =   6000
            Width           =   975
         End
         Begin VB.Data rsUser 
            Caption         =   "Data1"
            Connect         =   "Access 2000;"
            DatabaseName    =   "D:\Jørgen Programmer\ProgrammersHelper\Source\CodeMaster.mdb"
            DefaultCursorType=   0  'DefaultCursor
            DefaultType     =   2  'UseODBC
            Exclusive       =   0   'False
            Height          =   345
            Left            =   0
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   "User"
            Top             =   0
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Mail Server Name:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   58
            Top             =   4800
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "User Name:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Address:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   27
            Top             =   960
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   26
            Top             =   1560
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Town:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   25
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Country:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   24
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone Prefix:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   23
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Phone Number:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   22
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fax Number:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   21
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail Address:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   20
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Internet URL:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   19
            Top             =   4440
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Hourly Wage:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   18
            Top             =   5640
            Width           =   2295
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Currency:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   11
            Left            =   3480
            TabIndex        =   17
            Top             =   5640
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Next free Invoice No.:"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   16
            Top             =   6000
            Width           =   2175
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This is just a test label to control the label front color"
         Height          =   615
         Left            =   -73560
         TabIndex        =   69
         Top             =   3360
         Width           =   2775
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   -74280
         X2              =   -70080
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   -74280
         X2              =   -70080
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -70080
         X2              =   -70080
         Y1              =   2040
         Y2              =   5280
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -74280
         X2              =   -74280
         Y1              =   2040
         Y2              =   5280
      End
      Begin VB.Image BackgroundPic 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         DataField       =   "BackgroundPicture"
         DataSource      =   "rsUser"
         Height          =   3735
         Left            =   -74520
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vBookCountry() As Variant, boolFirst As Boolean
Dim rsLanguage As Recordset
Dim rsCodeLanguage As Recordset
Dim rsCountry As Recordset
Private Function IsLanguagePresent(strLanguage As String) As Boolean
    IsLanguagePresent = False
    On Error GoTo errLangPres
    With rsLanguage
        .MoveLast
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If .Fields("Language") = strLanguage Then
                IsLanguagePresent = True
                Exit For
            End If
        .MoveNext
        Next
    End With
    Exit Function
    
errLangPres:
    Beep
    MsgBox Err.Description, vbExclamation, "Is Language Present"
    Err.Clear
End Function


Private Sub LoadCodeLanguage()
    cboPreferedLanguage.Clear
    With rsCodeLanguage
        .MoveFirst
        Do While Not .EOF
            cboPreferedLanguage.AddItem .Fields("Language")
        .MoveNext
        Loop
    End With
End Sub
Private Function MakeNewLanguage(strLanguage As String)
Dim RS As Recordset, iNo As Integer, boolNotFound As Boolean, iCount As Integer
Dim db As DAO.Database
Dim tbl As DAO.TableDef
    Set db = DBEngine.OpenDatabase(App.Path & "\ProgramLang")
    ProgressBar1.Max = db.TableDefs.Count
    
    iCount = 0
    On Error GoTo errNewLanguage
    For Each tbl In db.TableDefs
        Select Case tbl.Name
        Case "MSysAccessObjects"
        Case "MSysObjects"
        Case "MSysQueries"
        Case "MSysRelationships"
        Case "MSysACEs"
        Case "SpellLanguage"
        Case Else
            Set RS = db.OpenRecordset(tbl.Name)
            RS.MoveLast
            iNo = RS.RecordCount
            iCount = iCount + 1
            ProgressBar1.Value = iCount
            RS.MoveFirst
            boolNotFound = True
            For i = 0 To iNo - 1
                If Trim(RS.Fields("Language")) = Trim(strLanguage) Then
                    boolNotFound = False
                    Exit For
                End If
            RS.MoveNext
            Next
            If boolNotFound Then
                Call MakeNewRecordset(RS, m_FileExt, strLanguage)
            End If
        End Select
    Next
    Exit Function
    
errNewLanguage:
    Beep
    MsgBox Err.Description, vbInformation, "New Language"
    Resume Next
End Function
Private Function MakeNewRecordset(RS As Recordset, strOldLanguage As String, strNewLanguage As String)
Dim rsClone As Recordset, fld As Field, n As Integer
    On Error GoTo errNewRecordset
    Set rsClone = RS.Clone()
    With rsClone
        .MoveLast
        .MoveFirst
        For i = 0 To .RecordCount - 1
            If .Fields("Language") = strOldLanguage Then
                RS.AddNew
                RS.Fields("Language") = Trim(strNewLanguage)
                For n = 1 To rsClone.Fields.Count - 1
                    RS.Fields(n) = rsClone.Fields(n)
                Next
                RS.Update
                Exit Function
            End If
            .MoveNext
        Next
    End With
    Exit Function
    
errNewRecordset:
    Beep
    MsgBox Err.Description, vbCritical, "New Recordset"
    Err.Clear
End Function

Private Sub ReadText()
Dim sHelp As String
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
                For i = 0 To 25
                    If IsNull(.Fields(i + 2)) Then
                        .Fields(1 + 2) = Label1(i).Caption
                    Else
                        Label1(i).Caption = .Fields(i + 2)
                    End If
                Next
                If IsNull(.Fields("Frame3")) Then
                    .Fields("Frame3") = Frame3.Caption
                Else
                    Frame3.Caption = .Fields("Frame3")
                End If
                If IsNull(.Fields("Frame4")) Then
                    .Fields("Frame4") = Frame4.Caption
                Else
                    Frame4.Caption = .Fields("Frame4")
                End If
                If IsNull(.Fields("btnBackground")) Then
                    .Fields("btnBackground") = btnBackground.Caption
                Else
                    btnBackground.Caption = .Fields("btnBackground")
                End If
                If IsNull(.Fields("btnFrame")) Then
                    .Fields("btnFrame") = btnFrame.Caption
                Else
                    btnFrame.Caption = .Fields("btnFrame")
                End If
                If IsNull(.Fields("btnLabel")) Then
                    .Fields("btnLabel") = btnLabel.Caption
                Else
                    btnLabel.Caption = .Fields("btnLabel")
                End If
                Tab1.Tab = 0
                If IsNull(.Fields("Tab10")) Then
                    .Fields("Tab10") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab10")
                End If
                Tab1.Tab = 1
                If IsNull(.Fields("Tab11")) Then
                    .Fields("Tab11") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab11")
                End If
                Tab1.Tab = 2
                If IsNull(.Fields("Tab12")) Then
                    .Fields("Tab12") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab12")
                End If
                Tab1.Tab = 3
                If IsNull(.Fields("Tab13")) Then
                    .Fields("Tab13") = Tab1.Caption
                Else
                    Tab1.Caption = .Fields("Tab13")
                End If
                Tab1.Tab = 0
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
        .Fields("Form") = Me.Caption
        For i = 0 To 25
            .Fields(i + 2) = Label1(i).Caption
        Next
        .Fields("Frame3") = Frame3.Caption
        .Fields("Frame4") = Frame4.Caption
        .Fields("btnBackground") = btnBackground.Caption
        .Fields("btnFrame") = btnFrame.Caption
        .Fields("btnLabel") = btnLabel.Caption
        Tab1.Tab = 0
        .Fields("Tab10") = Tab1.Caption
        Tab1.Tab = 1
        .Fields("Tab11") = Tab1.Caption
        Tab1.Tab = 2
        .Fields("Tab12") = Tab1.Caption
        Tab1.Tab = 3
        .Fields("Tab13") = Tab1.Caption
        Tab1.Tab = 0
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
            cmbLanguage.ItemData(cmbLanguage.NewIndex) = cmbLanguage.ListCount - 1
            vBookCountry(cmbLanguage.ListCount - 1) = .Bookmark
        .MoveNext
        Loop
    End With
End Sub

Private Sub btnBackground_Click()
        On Error Resume Next
        With CMD1
            .FileName = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .ShowOpen
        End With
        Set BackgroundPic.Picture = LoadPicture(CMD1.FileName)
        frmMDI.rsUser.Refresh
        frmMDI.Picture1.Refresh
End Sub

Private Sub btnCopy_Click()
    On Error Resume Next
    Clipboard.SetData Image1.Picture, vbCFDIB
End Sub

Private Sub btnDelete_Click()
    On Error Resume Next
    Set Image1.Picture = LoadPicture()
End Sub

Private Sub btnFrame_Click()
        On Error Resume Next
        CMD1.ShowColor
        For i = 0 To 3
            Line1(i).BorderColor = CMD1.Color
        Next
    With rsUser.Recordset
        .Edit
        .Fields("FrameColor") = CMD1.Color
        .Update
    End With
End Sub

Private Sub btnLabel_Click()
    On Error Resume Next
    CMD1.ShowColor
    Label2.ForeColor = CMD1.Color
    With rsUser.Recordset
        .Edit
        .Fields("LabelColor") = CMD1.Color
        .Update
    End With
End Sub

Private Sub btnPaste_Click()
        On Error Resume Next
        Image1.Picture = Clipboard.GetData(vbCFDIB)
End Sub

Private Sub btnReadFromFile_Click()
        On Error Resume Next
        With CMD1
            .FileName = ""
            .DialogTitle = "Load Picture from disk"
            .Filter = "Pictures (*.bmp; *.pcx;*.jpg;*.jpeg;*.gif)|*.bmp;*.pcx;*.jpg;*.jpeg;*.gif"
            .FilterIndex = 1
            .Action = 1
        End With
        Set Image1.Picture = LoadPicture(CMD1.FileName)
End Sub

Private Sub cmbLanguage_Click()
    On Error GoTo errLanguage
    If boolFirst Then Exit Sub
    rsCountry.Bookmark = vBookCountry(cmbLanguage.ItemData(cmbLanguage.ListIndex))
    cmbLanguage.Text = rsCountry.Fields("CountryFix")
    
    'check if the language is present in the database
    If IsLanguagePresent(cmbLanguage.Text) Then
        m_FileExt = Trim(cmbLanguage.Text)
        frmMDI.ReadText
        frmMDI.LoadMenu
        Exit Sub
    Else   'we did not have this language in the ProgramLang.mdb database
        Frame6.Visible = True
        MakeNewLanguage (cmbLanguage.Text) 'make new recordset for each form in database
        Frame6.Visible = False
    End If
    m_FileExt = Trim(cmbLanguage.Text)
    Exit Sub
    
    ReadText    'show this form-text
    frmMDI.ReadText
    frmMDI.LoadMenu
    Exit Sub
    
errLanguage:
    Beep
    MsgBox Err.Description, vbInformation, "Change Language"
    Err.Clear
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    rsUser.Refresh
    boolFirst = True
    ReadText
    LoadLanguage
    LoadCodeLanguage
    rsUser.Refresh
    DisableButtons 1
    boolFirst = False
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    rsUser.DatabaseName = m_strPrograming
    Set rsCodeLanguage = m_dbCodeSnippet.OpenRecordset("Language")
    Set rsCountry = m_dbPrograming.OpenRecordset("Country")
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmUser")
    m_iFormNo = 22
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Error$, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsUser.UpdateRecord
    rsUser.Recordset.Close
    rsCountry.Close
    rsCodeLanguage.Close
    rsLanguage.Close
    m_iFormNo = 0
    DisableButtons 2
    Set frmUser = Nothing
End Sub

Private Sub Tab1_Click(PreviousTab As Integer)
    If Tab1.Tab = 3 Then
        With rsUser.Recordset
            If Not IsNull(.Fields("FrameColor")) Then
                For i = 0 To 3
                    Line1(i).BorderColor = .Fields("FrameColor")
                Next
            End If
            If Not IsNull(.Fields("LabelColor")) Then
                Label2.ForeColor = .Fields("LabelColor")
            End If
        End With
    End If
End Sub
