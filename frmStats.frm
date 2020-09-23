VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStats 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6.0 Code Statistics"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10200
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Height          =   375
      Left            =   5280
      TabIndex        =   55
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6600
      TabIndex        =   54
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   53
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar pbrVariables 
      Height          =   255
      Left            =   7800
      TabIndex        =   52
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame fraVariables 
      BackColor       =   &H00404040&
      Caption         =   "Unused Variables"
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   50
      Top             =   3720
      Width           =   9975
      Begin VB.ListBox lstVars 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1710
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Frame framProj 
      BackColor       =   &H00404040&
      Caption         =   "Project Statistics"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4695
      Begin VB.Label lblControl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblDControl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "User Controls :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblClass 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblDClass 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Class Modules :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDMod 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Modules :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblForm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblDForm 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forms :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "v1.0.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDVer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Version :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Project1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblDName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraStructure 
      BackColor       =   &H00404040&
      Caption         =   "Code Structure"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   4695
      Begin VB.Label lblDProc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Procedures :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblProc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDFunc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Functions :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblDProp 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Properties :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblProp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDApi 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "API Declarations :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblApi 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame fraBreakdown 
      BackColor       =   &H00404040&
      Caption         =   "Code Breakdown"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   4920
      TabIndex        =   15
      Top             =   600
      Width           =   5175
      Begin VB.Label lblWhile 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   49
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblDWhile 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Do-While Loops :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblSelect 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   47
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDSelect 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Select Statments :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   46
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblFor 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblDFor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "For Loops :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblEnum 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblDEnum 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Enumerators Declared :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   32
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblDType 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Types Declared :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblIf 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   30
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblDIf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "If Statements :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblConst 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   28
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblDConstants 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Constants Declared :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblDVariables 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Variables Declared :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblVar 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraLines 
      BackColor       =   &H00404040&
      Caption         =   "Lines"
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   4920
      TabIndex        =   35
      Top             =   2520
      Width           =   5175
      Begin VB.Label lblDBlank 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Blank Lines :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   42
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDComm 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comment Lines :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblComm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         TabIndex        =   40
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lines :"
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
         Left            =   2520
         TabIndex        =   39
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   4320
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code Lines :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   120
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog cdgFiles 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FormStartCode = "Attribute VB_Exposed "
Const ModStartCode = "Attribute VB_Name "
Const ClsStartCode = "Attribute VB_Exposed"
Const CtlStartCode = "Attribute VB_Exposed"
Const VbpTitle = "Title"
Const VbpMajor = "MajorVer"
Const VbpMinor = "MinorVer"
Const VbpRevision = "RevisionVer"
Const VbpForm = "Form"
Const VbpMod = "Module"
Const VbpClass = "Class" 'This is actually made up of "Class="<object name>"; "<class filename>"
Const VbpControl = "UserControl"
Const BrowseFilter = "VB Project *.Vbp|*.Vbp|VB Modules *.Bas|*.Bas|VB Forms *.Frm|*.Frm|VB Class Modules (*.Cls)|*.Cls|VB User Controls (*.Ctl)|*.Ctl|All Files *.*|*.*"
Const FUNC = "Function "
Const PROC = "Sub "
Const PROP = "Property "

'the different styles of variable modes
Private Enum VarModeEnum
    varPublic = 0
    varPrivate = 1  'default variable declaration using Dim
    varGlobal = 2
    varStatic = 3
    varModule = 4
End Enum

'used to keep track of variables within the
'program and their locations.
Private Type TrackVarType
    strVarName As String
    strVarProc As String
    strVarLocation As String
    enmVarMode As VarModeEnum
    blnVarUsed As Boolean
End Type

'variable tracker
Private mudtVariables() As TrackVarType
Private mudtCurLoc As TrackVarType      'the current location within the project while scanning
Private mblnScanning As Boolean         'if the user is scanning for variables, then True

'code and project counters
Private NumBlank As Long
Private NumProc As Long
Private NumFunc As Long
Private NumComments As Long
Private NumForms As Long
Private NumModules As Long
Private NumClasses As Long
Private NumControls As Long
Private NumProperties As Long
Private NumCode As Long
Private NumVariables As Long
Private NumVarLines As Long
Private NumConst As Long
Private NumType As Long
Private NumEnum As Long
Private NumIf As Long
Private NumSel As Long
Private NumFor As Long
Private NumWhile As Long
Private NumAPI As Long
Private Version As String
Private ProjectName As String

Dim rsLanguage As Recordset
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
                If IsNull(.Fields("lblFile")) Then
                    .Fields("lblFile") = lblFile.Caption
                Else
                    lblFile.Caption = .Fields("lblFile")
                End If
                If IsNull(.Fields("lblDName")) Then
                    .Fields("lblDName") = lblDName.Caption
                Else
                    lblDName.Caption = .Fields("lblDName")
                End If
                If IsNull(.Fields("lblDVer")) Then
                    .Fields("lblDVer") = lblDVer.Caption
                Else
                    lblDVer.Caption = .Fields("lblDVer")
                End If
                If IsNull(.Fields("lblDForm")) Then
                    .Fields("lblDForm") = lblDForm.Caption
                Else
                    lblDForm.Caption = .Fields("lblDForm")
                End If
                If IsNull(.Fields("lblDMod")) Then
                    .Fields("lblDMod") = lblDMod.Caption
                Else
                    lblDMod.Caption = .Fields("lblDMod")
                End If
                If IsNull(.Fields("lblDClass")) Then
                    .Fields("lblDClass") = lblDClass.Caption
                Else
                    lblDClass.Caption = .Fields("lblDClass")
                End If
                If IsNull(.Fields("lblDControl")) Then
                    .Fields("lblDControl") = lblDControl.Caption
                Else
                    lblDControl.Caption = .Fields("lblDControl")
                End If
                If IsNull(.Fields("lblDVariables")) Then
                    .Fields("lblDVariables") = lblDVariables.Caption
                Else
                    lblDVariables.Caption = .Fields("lblDVariables")
                End If
                If IsNull(.Fields("lblDConstants")) Then
                    .Fields("lblDConstants") = lblDConstants.Caption
                Else
                    lblDConstants.Caption = .Fields("lblDConstants")
                End If
                If IsNull(.Fields("lblDType")) Then
                    .Fields("lblDType") = lblDType.Caption
                Else
                    lblDType.Caption = .Fields("lblDType")
                End If
                If IsNull(.Fields("lblDEnum")) Then
                    .Fields("lblDEnum") = lblDEnum.Caption
                Else
                    lblDEnum.Caption = .Fields("lblDEnum")
                End If
                If IsNull(.Fields("lblDIf")) Then
                    .Fields("lblDIf") = lblDIf.Caption
                Else
                    lblDIf.Caption = .Fields("lblDIf")
                End If
                If IsNull(.Fields("lblDSelect")) Then
                    .Fields("lblDSelect") = lblDSelect.Caption
                Else
                    lblDSelect.Caption = .Fields("lblDSelect")
                End If
                If IsNull(.Fields("lblDFor")) Then
                    .Fields("lblDFor") = lblDFor.Caption
                Else
                    lblDFor.Caption = .Fields("lblDFor")
                End If
                If IsNull(.Fields("lblDWhile")) Then
                    .Fields("lblDWhile") = lblDWhile.Caption
                Else
                    lblDWhile.Caption = .Fields("lblDWhile")
                End If
                If IsNull(.Fields("lblDProc")) Then
                    .Fields("lblDProc") = lblDProc.Caption
                Else
                    lblDProc.Caption = .Fields("lblDProc")
                End If
                If IsNull(.Fields("lblDFunc")) Then
                    .Fields("lblDFunc") = lblDFunc.Caption
                Else
                    lblDFunc.Caption = .Fields("lblDFunc")
                End If
                If IsNull(.Fields("lblDProp")) Then
                    .Fields("lblDProp") = lblDProp.Caption
                Else
                    lblDProp.Caption = .Fields("lblDProp")
                End If
                If IsNull(.Fields("lblDApi")) Then
                    .Fields("lblDApi") = lblDApi.Caption
                Else
                    lblDApi.Caption = .Fields("lblDApi")
                End If
                If IsNull(.Fields("lblDBlank")) Then
                    .Fields("lblDBlank") = lblDBlank.Caption
                Else
                    lblDBlank.Caption = .Fields("lblDBlank")
                End If
                If IsNull(.Fields("lblDComm")) Then
                    .Fields("lblDComm") = lblDComm.Caption
                Else
                    lblDComm.Caption = .Fields("lblDComm")
                End If
                If IsNull(.Fields("lblDCode")) Then
                    .Fields("lblDCode") = lblDCode.Caption
                Else
                    lblDCode.Caption = .Fields("lblDCode")
                End If
                If IsNull(.Fields("lblDTotal")) Then
                    .Fields("lblDTotal") = lblDTotal.Caption
                Else
                    lblDTotal.Caption = .Fields("lblDTotal")
                End If
                If IsNull(.Fields("framProj")) Then
                    .Fields("framProj") = framProj.Caption
                Else
                    framProj.Caption = .Fields("framProj")
                End If
                If IsNull(.Fields("fraBreakdown")) Then
                    .Fields("fraBreakdown") = fraBreakdown.Caption
                Else
                    fraBreakdown.Caption = .Fields("fraBreakdown")
                End If
                If IsNull(.Fields("fraStructure")) Then
                    .Fields("fraStructure") = fraStructure.Caption
                Else
                    fraStructure.Caption = .Fields("fraStructure")
                End If
                If IsNull(.Fields("fraLines")) Then
                    .Fields("fraLines") = fraLines.Caption
                Else
                    fraLines.Caption = .Fields("fraLines")
                End If
                If IsNull(.Fields("fraVariables")) Then
                    .Fields("fraVariables") = fraVariables.Caption
                Else
                    fraVariables.Caption = .Fields("fraVariables")
                End If
                'If IsNull(.Fields("cmdBrowse")) Then
                    '.Fields("cmdBrowse") = cmdBrowse.Caption
                'Else
                    'cmdBrowse.Caption = .Fields("cmdBrowse")
                'End If
                If IsNull(.Fields("cmdScan")) Then
                    .Fields("cmdScan") = cmdScan.Caption
                Else
                    cmdScan.Caption = .Fields("cmdScan")
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
        .Fields("Form") = Me.Caption
        .Fields("lblFile") = lblFile.Caption
        .Fields("lblDName") = lblDName.Caption
        .Fields("lblDVer") = lblDVer.Caption
        .Fields("lblDForm") = lblDForm.Caption
        .Fields("lblDMod") = lblDMod.Caption
        .Fields("lblDClass") = lblDClass.Caption
        .Fields("lblDControl") = lblDControl.Caption
        .Fields("lblDVariables") = lblDVariables.Caption
        .Fields("lblDConstants") = lblDConstants.Caption
        .Fields("lblDType") = lblDType.Caption
        .Fields("lblDEnum") = lblDEnum.Caption
        .Fields("lblDIf") = lblDIf.Caption
        .Fields("lblDSelect") = lblDSelect.Caption
        .Fields("lblDFor") = lblDFor.Caption
        .Fields("lblDWhile") = lblDWhile.Caption
        .Fields("lblDProc") = lblDProc.Caption
        .Fields("lblDFunc") = lblDFunc.Caption
        .Fields("lblDProp") = lblDProp.Caption
        .Fields("lblDApi") = lblDApi.Caption
        .Fields("lblDBlank") = lblDBlank.Caption
        .Fields("lblDComm") = lblDComm.Caption
        .Fields("lblDCode") = lblDCode.Caption
        .Fields("lblDTotal") = lblDTotal.Caption
        .Fields("framProj") = framProj.Caption
        .Fields("fraBreakdown") = fraBreakdown.Caption
        .Fields("fraStructure") = fraStructure.Caption
        .Fields("fraLines") = fraLines.Caption
        .Fields("fraVariables") = fraVariables.Caption
        '.Fields("cmdBrowse") = cmdBrowse.Caption
        .Fields("cmdScan") = cmdScan.Caption
        .Fields("Help") = sHelp
        .Update
    End With
End Sub

Public Sub ResetValues()
    'reset values and variables
    
    lblName.Caption = ""
    Version = ""
    lstVars.Clear
    fraVariables.Caption = "Unused Variables"
    ReDim mudtVariables(0)
    mudtCurLoc.strVarProc = "Module"
    NumBlank = 0
    NumProc = 0
    NumFunc = 0
    NumComments = 0
    NumForms = 0
    NumModules = 0
    NumCode = 0
    NumVariables = 0
    NumVarLines = 0
    NumClasses = 0
    NumProperties = 0
    NumAPI = 0
    NumControls = 0
    NumConst = 0
    NumType = 0
    NumEnum = 0
    NumIf = 0
    NumSel = 0
    NumFor = 0
    NumWhile = 0
End Sub

Public Sub DisplayValues(Optional ByVal blnNoList = False)
    'This will enter all the appropiate details into the lables and
    'total the number of lines found
    
    'half the number of properties countes because there are two property
    'statements per property, Let and Get.
    'NumProperties = NumProperties / 2
    
    'display results
    If Trim(lblName.Caption) = "" Then
        'if the project name is blank then use the default name
        lblName.Caption = "Project1"
    End If
    'If (LCase(Version) = "v") Or (LCase(lblVersion.Caption) = "v") Then
    If Version = "" Then
        'if version if blank, then set it to default
        Version = "v1.0.0"
    End If
    lblVersion.Caption = Version
    lblBlank.Caption = Format(NumBlank, "0")
    lblComm.Caption = Format(NumComments, "0")
    lblForm.Caption = Format(NumForms, "0")
    lblMod.Caption = Format(NumModules, "0")
    lblClass.Caption = Format(NumClasses, "0")
    lblProc.Caption = Format(NumProc, "0")
    lblFunc.Caption = Format(NumFunc, "0")
    lblProp.Caption = Format(NumProperties / 2, "0")
    lblCode.Caption = Format(NumCode, "0")
    lblVar.Caption = Format(NumVariables, "0")
    lblControl.Caption = Format(NumControls, "0")
    lblApi.Caption = Format(NumAPI, "0")
    lblConst.Caption = Format(NumConst, "0")
    lblType.Caption = Format(NumType, "0")
    lblEnum.Caption = Format(NumEnum, "0")
    lblIf.Caption = Format(NumIf, "0")
    lblSelect.Caption = Format(NumSel, "0")
    lblFor.Caption = Format(NumFor, "0")
    lblWhile.Caption = Format(NumWhile, "0")
    
    'total results accounting for headers/footers of procedures/functions, types, enumerators etc.
    lblTotal.Caption = Format(GetTotal, "0")
    
    'display unused variables (if any)
    If (Not blnNoList) And mblnScanning Then
        Call ShowUnusedVars
    End If
End Sub

Private Function GetTotal() As Long
    'This will total up the number of lines
    On Error GoTo errGetTotal
    GetTotal = (NumBlank + NumComments + _
                    ((NumProc + NumFunc + _
                      NumProperties + NumType + _
                      NumEnum) _
                     * 2) + _
                NumConst + NumAPI + _
                NumVarLines + NumCode)
    Exit Function
    
errGetTotal:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Total"
    Err.Clear
End Function

Public Sub ReadProject(ByVal strPath As String)
    'This will read an entire project and set the values for statistics
    On Error GoTo errReadProject
    Dim intFileNum As Integer 'used for the .vbp file
    Dim strLine As String
    Dim blnStartScan As Boolean
    
    'if path is invalid, then quit
    If Dir(strPath) = "" Then
        Exit Sub
    End If
    
    Call ResetValues
    blnStartScan = False
    
    'open project
    intFileNum = FreeFile
    Open strPath For Input As #intFileNum
        While Not EOF(intFileNum)
            Line Input #intFileNum, strLine
            
            Select Case GetBefore(strLine)
            Case VbpTitle
                lblName.Caption = GetAfter(strLine)
            
            Case VbpMajor
                Version = "v" & GetAfter(strLine) & "."
            
            Case VbpMinor
                Version = Version & GetAfter(strLine) & "."
            
            Case VbpRevision
                Version = Version & GetAfter(strLine)
            
            Case VbpForm
                'scan form
                NumForms = NumForms + 1
                Call ScanFile(AddFile(GetPath(strPath), _
                                      GetAfter(strLine)), _
                              FormStartCode)
                
            Case VbpMod
                'scan module
                NumModules = NumModules + 1
                Call ScanFile(AddFile(GetPath(strPath), GetMod(strLine)), ModStartCode)
            
            Case VbpClass
                'scan class module
                NumClasses = NumClasses + 1
                Call ScanFile(AddFile(GetPath(strPath), GetClass(strLine)), ClsStartCode)
                
            Case VbpControl
                'scan user control
                NumControls = NumControls + 1
                Call ScanFile(AddFile(GetPath(strPath), GetAfter(strLine)), CtlStartCode)
            
            End Select
            
        Wend
    Close #intFileNum
    
    Call DisplayValues
    Exit Sub
    
errReadProject:
    Beep
    MsgBox Err.Description, vbExclamation, "Read Project"
    Err.Clear
End Sub

Public Sub IncrementVal(ByVal strLine As String)
    'This will increment the appropiate values based on the text
    
    'the following constants are keywords to look for in the text
    Const EndProc = "End Sub"
    Const EndFunc = "End Function"
    Const EndProp = "End Property"
    Const DecApi = "Declare "
    Const LibApi = " Lib "
    Const VarA = "Public"
    Const VarB = "Private"
    Const VarC = "Global"
    Const VarD = "Dim"
    Const VarE = "Static"
    Const VarAs = " As "
    Const Constant = "Const "
    Const EndType = "End Type"
    Const EndEnum = "End Enum"
    Const EdIf = "End If"
    Const EndSel = "End Select"
    Const ForL = "For "             'For loop
    Const DoL = "Do "               'Do or Do While loop
    Const WhileL = "While "         'While loop
    Const Comment = "'"
    Const Blank = ""
    
    On Error GoTo errIncrementVal
    Static strNextLine As String    'used to temperorily hold sections of a line until they are loaded. strLine sections are added by checking for the "_" character at the end of the line
    
    'continue line character ("_") - the underscore
    If Right(strLine, 1) = "_" Then
        'don't count anything, but remember the
        'line section
        strNextLine = strNextLine & Left(strLine, Len(strLine) - 1)
        Exit Sub
    Else
        'if the current line section is empty
        'then don't do anything, other wise
        'we have reached the end of the line
        'section. Process the entire line
        If strNextLine <> "" Then
            'complete the line section
            strNextLine = strNextLine & strLine
            
            'process the complete line
            strLine = strNextLine
            
            'line section has been completed
            'ad is about to be processed, we do
            'not need to hold it any more
            strNextLine = ""
        End If
    End If
    
    'Comments
    If Left(strLine, 1) = Comment Then
        NumComments = NumComments + 1
        Exit Sub
    End If
    
    'Blanks
    If strLine = Blank Then
        NumBlank = NumBlank + 1
        Exit Sub
    End If
    
    'the footers of the functions, procedures and properties.
    'I'm counting the footers because they are always the
    'same no matter what keywords the title has.
    If Left(strLine, Len(EndProc)) = EndProc Then
        NumProc = NumProc + 1
        
        'code num as already counted the header, so subtract this.
        NumCode = NumCode - 1
        
        'set the current location within the project
        mudtCurLoc.strVarName = ""
        mudtCurLoc.enmVarMode = varModule
        Exit Sub
    End If
    If Left(strLine, Len(EndFunc)) = EndFunc Then
        NumFunc = NumFunc + 1
        
        'code num as already counted the header, so subtract this.
        NumCode = NumCode - 1
        
        'set the current location within the project
        mudtCurLoc.strVarName = ""
        mudtCurLoc.enmVarMode = varModule
        Exit Sub
    End If
    If Left(strLine, Len(EndFunc)) = EndProp Then
        NumProperties = NumProperties + 1
        
        'code num as already counted the header, so subtract this.
        NumCode = NumCode - 1
        
        'set the current location within the project
        mudtCurLoc.strVarName = ""
        mudtCurLoc.enmVarMode = varModule
        Exit Sub
    End If
    
    'check for api declarations
    If (InStr(1, strLine, DecApi) <> 0) And IsNotInQuote(strLine, DecApi) And (InStr(1, strLine, LibApi) <> 0) Then
        NumAPI = NumAPI + 1
        Exit Sub
    End If
    
    'constant declarations
    If (InStr(1, strLine, Constant) <> 0) And IsNotInQuote(strLine, Constant) Then
        NumConst = NumConst + 1
        Exit Sub
    End If
    
    'get the procedure and function names for tracking
    'variables
    If IsNotInQuote(strLine, FUNC) _
       And (InStr(strLine, FUNC) <> 0) _
       And (Left(Trim(strLine), 4) <> "Exit") Then
        
        'set the current location within the project
        mudtCurLoc.strVarName = GetName(strLine, FUNC)
        mudtCurLoc.enmVarMode = varPrivate
    End If
    If IsNotInQuote(strLine, PROC) _
       And (InStr(strLine, PROC) <> 0) _
       And (Left(Trim(strLine), 4) <> "Exit") _
       And IsWord(strLine, PROC) Then
        
        'set the current location within the project
        mudtCurLoc.strVarName = GetName(strLine, PROC)
        mudtCurLoc.enmVarMode = varPrivate
    End If
    If IsNotInQuote(strLine, PROP) _
       And (InStr(strLine, PROP) <> 0) _
       And (Left(Trim(strLine), 4) <> "Exit") Then
        
        'set the current location within the project
        mudtCurLoc.strVarName = GetName(strLine, PROP)
        mudtCurLoc.enmVarMode = varPrivate
    End If
    
    'variable declarations
    'if the left part of the string contains one of the variable decalration
    'keywords and also contains the keyword " As " and does not contain
    'the api declaration keyword "Declare", then the string is a variable
    'declaration.
    'NOTE: These variables do NOT count procedure/function parameters.
    'Also, the number of variables is not the same as the number of
    'lines used to declare them eg,
    'Dim MyVar1, MyVar2, MyVar3 As Integer
    If ((Left(strLine, Len(VarA)) = VarA) _
        Or (Left(strLine, Len(VarB)) = VarB) _
        Or (Left(strLine, Len(VarC)) = VarC) _
        Or (Left(strLine, Len(VarD)) = VarD) _
        Or (Left(strLine, Len(VarE)) = VarE)) _
       And (InStr(1, strLine, VarAs) <> 0) _
       And (InStr(1, strLine, DecApi) = 0) Then
        
        'get the variable names
        If mblnScanning Then
            Call GetVarNames(strLine)
        End If
        
        NumVariables = NumVariables + 1 + CommaCount(strLine)
        NumVarLines = NumVarLines + 1
        Exit Sub
    End If
    
    'defined Types
    If Left(strLine, Len(EndType)) = EndType Then
        NumType = NumType + 1
        NumCode = NumCode - 1
        Exit Sub
    End If
    
    'enumerators
    If Left(strLine, Len(EndEnum)) = EndEnum Then
        NumEnum = NumEnum + 1
        NumCode = NumCode - 1
        Exit Sub
    End If
    
    'else the line is code
    NumCode = NumCode + 1
    Call UpdateVars(strLine)
    
    'the following are counted as code, but we want to count them seperatly
    
    'If statements
    If Left(strLine, Len(EdIf)) = EdIf Then
        NumIf = NumIf + 1
        Exit Sub
    End If
    
    'Select statements
    If Left(strLine, Len(EndSel)) = EndSel Then
        NumSel = NumSel + 1
        Exit Sub
    End If
    
    'For loops
    If Left(strLine, Len(ForL)) = ForL Then
        NumFor = NumFor + 1
        Exit Sub
    End If
    
    'Do, Loop and While loops
    If (Left(strLine, Len(DoL)) = DoL) Or (Left(strLine, Len(WhileL)) = WhileL) Then
        NumWhile = NumWhile + 1
    End If
    Exit Sub
    
errIncrementVal:
    Beep
    MsgBox Err.Description, vbExclamation, "Increment values"
    Err.Clear
End Sub

Public Function IsNotInQuote(ByVal strText As String, _
                             ByVal strWords As String) _
                             As Boolean
    'This function will tell you if the specified text is in quotes within
    'the second string. It does this by counting the number of quotation
    'marks before the specified strWords. If the number is even, then the
    'strWords are not in qototes, otherwise they are.
    
    'the quotation mark, " , is ASCII character 34
    
    Dim lngGotPos As Long
    Dim lngCounter As Long
    Dim lngNextPos As Long
    
    On Error GoTo errIsNotInQuote
    'find where the position of strWords in strText
    lngGotPos = InStr(1, strText, strWords)
    If lngGotPos = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'start counting the number of quotation marks
    lngNextPos = 0
    Do
        lngNextPos = InStr(lngNextPos + 1, strText, Chr(34))
        
        If (lngNextPos <> 0) And (lngNextPos < lngGotPos) Then
            'quote found, add to total
            lngCounter = lngCounter + 1
        End If
    Loop Until (lngNextPos = 0) Or (lngNextPos >= lngGotPos)
    
    'no quotes at all found
    If lngCounter = 0 Then
        IsNotInQuote = True
        Exit Function
    End If
    
    'if the number of quotes is even, then return true, else return false
    If lngCounter Mod 2 = 0 Then
        IsNotInQuote = True
    End If
    Exit Function
    
errIsNotInQuote:
    Beep
    MsgBox Err.Description, vbExclamation, "Is Not In Quote"
    Err.Clear
End Function

Public Function GetPath(ByVal strAddress As String) _
                        As String
    'This function returns the path from a string containing the full
    'path and filename of a file.
    
    Dim intLastPos As Integer
    On Error GoTo errGetpath
    'find the position of the last "\" mark in the string
    intLastPos = InStrRev(strAddress, "\")
    
    'if no \ found in the string, then
    If intLastPos = 0 Then
        'return the whole string
        intLastPos = Len(strAddress) + 1
    End If
    
    'return everything before the last "\" mark
    GetPath = Left(strAddress, (intLastPos - 1))
    Exit Function
    
errGetpath:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Path"
    Err.Clear
End Function

Public Function GetBefore(ByVal strSentence As String) _
                          As String
    'This procedure returns all the character of a
    'string before the "=" sign.
    
    Const Sign = "="
    
    Dim intCounter As Integer
    Dim strBefore As String
    
    On Error GoTo errGetBefore
    'find the position of the equals sign
    intCounter = InStr(1, strSentence, Sign)
    
    If (intCounter <> Len(strSentence)) And (intCounter <> 0) Then
        strBefore = Left(strSentence, (intCounter - 1))
    Else
        strBefore = ""
    End If
    
    GetBefore = strBefore
    Exit Function
    
errGetBefore:
    Beep
    MsgBox Err.Description, vbExclamation, "GetBefore"
    Err.Clear
End Function

Public Function GetAfter(ByVal strSentence As String, _
                         Optional ByVal strCharacter As String = "=") _
                         As String
    'This procedure returns all the character of a
    'string after the "=" sign.
    
    'Const strSign = "="
    
    Dim intCounter As Integer
    Dim strRest As String
    Dim strSign As String
    
    strSign = strCharacter
    On Error GoTo errGetAfter
    'find the last position of the specified sign
    intCounter = InStrRev(strSentence, strSign)
    
    If intCounter <> Len(strSentence) Then
        strRest = Right(strSentence, (Len(strSentence) - (intCounter + Len(strCharacter) - 1)))
    Else
        strRest = ""
    End If
    
    GetAfter = strRest
    Exit Function
    
errGetAfter:
    Beep
    MsgBox Err.Description, vbExclamation, "Get After"
    Err.Clear
End Function

Public Function GetMod(ByVal strSentence As String) _
                       As String
    'This procedure returns all the character of a
    'string after the ";" sign.
    
    Const ModName = ";"
    
    Dim strRest As String
    Dim intModPos As Integer
    
    On Error GoTo errGetMod
    'find the position of the ; sign
    intModPos = InStr(1, strSentence, ModName) + 1
    
    If intModPos <> Len(strSentence) Then
        strRest = Right(strSentence, (Len(strSentence) - intModPos))
    Else
        strRest = ""
    End If
    
    GetMod = strRest
    Exit Function
    
errGetMod:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Mod"
    Err.Clear
End Function

Public Function GetClass(ByVal strSentence As String) _
                         As String
    'This procedure returns all the character of a
    'string after the "; " sign.
    
    Const ClassName = "; "
    
    Dim strRest As String
    Dim intClassPos As Integer
    
    On Error GoTo errGetClass
    'find the position of the ; sign
    intClassPos = InStr(1, strSentence, ClassName) + 1
    
    If intClassPos <> Len(strSentence) Then
        strRest = Right(strSentence, (Len(strSentence) - intClassPos))
    Else
        strRest = ""
    End If
    
    GetClass = strRest
    Exit Function
    
errGetClass:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Class"
    Err.Clear
End Function

Private Sub cmdBrowse_Click()
    cdgFiles.Filter = BrowseFilter
    cdgFiles.InitDir = GetPath(txtPath.Text)
    cdgFiles.ShowOpen
    txtPath.Text = cdgFiles.FileName
End Sub

Private Sub cmdScan_Click()
    'Try to scan the file specified in the text box
    
    Const ProjExt = "vbp"
    Const FormExt = "frm"
    Const ModuleExt = "bas"
    Const ClassExt = "cls"
    Const ControlExt = "ctl"
    
    Dim strExtention As String
    Dim strFilePath As String
    
    strFilePath = txtPath.Text
    strExtention = GetAfter(strFilePath, ".")
    
    'don't try to scan file if it doesn't exist
    If (Dir(strFilePath) = "") Or (strFilePath = "") Then
        Exit Sub
    End If
    
    'scan each file type differently
    Select Case LCase(strExtention)
    Case LCase(ProjExt)
        'scan an entire project
        Call ReadProject(strFilePath)
    
    Case LCase(FormExt)
        'scan one form
        Call ResetValues
        NumForms = NumForms + 1
        Call ScanFile(strFilePath, FormStartCode)
        Call DisplayValues
    
    Case LCase(ModuleExt)
        'scan one module
        Call ResetValues
        NumModules = NumModules + 1
        Call ScanFile(strFilePath, ModStartCode)
        Call DisplayValues
    
    Case LCase(ClassExt)
        'scan one class
        Call ResetValues
        NumClasses = NumClasses + 1
        Call ScanFile(strFilePath, ClsStartCode)
        Call DisplayValues
        
    Case LCase(ControlExt)
        'scan one user control
        Call ResetValues
        NumControls = NumControls + 1
        Call ScanFile(strFilePath, CtlStartCode)
        Call DisplayValues
        
    End Select
    FileFind
End Sub

Private Sub ScanFile(ByVal strPath As String, _
                     ByVal strStart As String)
    'This procedure will scan a file starting at the first point with the
    'specified starting string.
    
    Dim intFileNum As Integer
    Dim strLine As String
    Dim blnStartScan As Boolean
    
    intFileNum = FreeFile
    
    On Error GoTo errScanFile
    If Dir(strPath) = "" Then
        'invalid path
        Exit Sub
    End If
    
    'remember the file we are scanning
    mudtCurLoc.strVarLocation = GetAfter(strPath, "\")
    mudtCurLoc.enmVarMode = varModule
    
    Open strPath For Input As #intFileNum
        'scan file
        While Not EOF(intFileNum)
            Line Input #intFileNum, strLine
            If blnStartScan Then
                Call IncrementVal(LTrim(strLine))
            End If
            
            If Left(strLine, Len(strStart)) = strStart Then
                'scan code
                blnStartScan = True
            End If
            
            If mblnScanning Then
                If GetTotal <= pbrVariables.Max Then
                    pbrVariables.Value = GetTotal
                    DoEvents
                End If
            End If
        Wend
    Close #intFileNum
    Exit Sub
    
errScanFile:
    Beep
    MsgBox Err.Description, vbExclamation, "Scan File"
    Err.Clear
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
    frmMDI.Toolbar1.Buttons(11).Enabled = True
End Sub
Private Sub Form_Load()
    On Error GoTo errForm_Load
    Me.Move 0, 0
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmStats")
    txtPath.Text = App.Path
    txtPath.SelLength = Len(txtPath.Text)
    m_iFormNo = 31
    Exit Sub
    
errForm_Load:
    Beep
    MsgBox Err.Description, vbCritical, "Load Form"
    Err.Clear
    Unload Me
End Sub

Public Function CommaCount(ByVal strLine As String) _
                           As Integer
    'This will return the number of commas found in the string. Mainly
    'use to find the number of variables declared on the same line
    
    Dim intCounter As Integer
    Dim intLastPos As Integer
    Dim intCommaNum As Integer
    
    intLastPos = 0
    On Error GoTo errCommaCount
    Do
        intCounter = InStr(intLastPos + 1, strLine, ",")
        
        If intCounter <> 0 Then
            intCommaNum = intCommaNum + 1
        End If
        intLastPos = intCounter
    Loop Until intLastPos = 0
    
    'return result
    CommaCount = intCommaNum
    Exit Function
    
errCommaCount:
    Beep
    MsgBox Err.Description, vbExclamation, "Comma Count"
    Err.Clear
End Function

Public Function AddFile(ByVal strDirectory As String, _
                        ByVal strFileName As String) _
                        As String
    'This will add a file name to a directory path to create a full filepath.
    On Error GoTo errAddFile
    If Right(strDirectory, 1) <> "\" Then
        'insert a backslash
        strDirectory = strDirectory & "\"
    End If
    
    'append the file name to the directory path now
    AddFile = strDirectory & strFileName
    Exit Function
    
errAddFile:
    Beep
    MsgBox Err.Description, vbExclamation, "Add File"
    Err.Clear
End Function

Private Function GetName(ByVal strLine As String, _
                         ByVal strMode As String) _
                         As String
    On Error GoTo errGetName
    'This will get the procedure, function tr
    'property name from a line of text
    
    Dim strName As String
    
    strName = Trim(GetAfter(strLine, strMode))
    
    GetName = Left(strName, InStr(strName, "(") - 1)
    Exit Function
    
errGetName:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Name"
    Err.Clear
End Function

Private Sub GetVarNames(ByVal strLine As String)
    'This procedure will get the variable names
    'from the string provided and add them into
    'the array
    
    Dim lngCounter As Long      'used to cycle through the array
    Dim strVars() As String     'a list of variables found in the array
    Dim lngMode As VarModeEnum  'the mode of the variable(s)
    Dim lngCommaCount As Long   'the number of commas in the string
    Dim lngOldSize As Long      'the current size of the variable array before resizing to add new variables
    Dim strVarName As String    'the variable name
    Dim lngOffset As Long       'the number of rejected variable names
    
    On Error GoTo errGetVarNames
    If mudtCurLoc.strVarLocation = "" Then
        Exit Sub
    End If
    
    'strip any comments from the end of the string
    If InStr(strLine, "'") > 0 Then
        strLine = Trim(Left(strLine, InStr(strLine, "'") - 1))
    End If
    
    'check for the level of the variable
    Select Case Left(strLine, InStr(strLine, " ") - 1)
    Case "Public"
        lngMode = varPublic
    
    Case "Private"
        If mudtCurLoc.strVarName = "" Then
            lngMode = varModule
        Else
            lngMode = varPrivate
        End If
    
    Case "Static"
        lngMode = varStatic
    
    Case "Dim"
        If mudtCurLoc.strVarName = "" Then
            lngMode = varModule
        Else
            lngMode = varPrivate
        End If
    
    Case "Global"
        lngMode = varGlobal
    
    Case Else
        'not a variable
        Exit Sub
    End Select
    
    If InStr(strLine, "(") > 0 Then
        lngMode = varPrivate
    
        'strip the first word from the string as we do
        'not need it
        strLine = Replace(strLine, "ByVal ", "")
        strLine = Replace(strLine, "ByRef ", "")
        strLine = Replace(strLine, "Optional ", "")
        strLine = Replace(strLine, "Friend ", "")
        strLine = Replace(strLine, "Static ", "")
        strLine = Replace(strLine, "ParamArray ", "")
        strLine = Trim(Mid(strLine, InStr(strLine, "(") + 1))
    
    Else
        'strip the first word from the string as we do
        'not need it
        strLine = Trim(Mid(strLine, InStr(strLine, " ")))
    End If
    
    'if there is more than one variable declared
    'in the line, then store all of them in the array
    lngCommaCount = CommaCount(strLine)
    If lngCommaCount > 0 Then
        'put the list of variables into the array
        ReDim strVars(lngCommaCount)
        strVars() = Split(strLine, ",")
        
        'resize the variable tracking array
        lngOldSize = UBound(mudtVariables)
        ReDim Preserve mudtVariables(lngOldSize + lngCommaCount + 1)
        
        'enter the variables into the array
        For lngCounter = 0 To lngCommaCount
            'get the variable name
            strVarName = Trim(Left(strVars(lngCounter + lngOffset), _
                                   InStr(Trim(strVars(lngCounter + lngOffset)), _
                                         " ")))
            
            'account for array brackets by
            'removing them
            If Right(strVarName, 2) = "()" Then
                strVarName = Left(strVarName, Len(strVarName) - 2)
            End If
            
            If strVarName <> "" Then
                With mudtVariables(lngOldSize + lngCounter + 1)
                    .strVarLocation = mudtCurLoc.strVarLocation
                    .strVarProc = mudtCurLoc.strVarName
                    .enmVarMode = lngMode
                    .strVarName = strVarName
                End With
            Else
                'a rejected variable name
                lngOffset = lngOffset + 1
            End If
        Next lngCounter
        
        'resize to account for rejected variable names
        ReDim Preserve mudtVariables(lngOldSize + lngCommaCount + 1 - lngOffset)
    Else
        'just enter one new variable
        ReDim Preserve mudtVariables(UBound(mudtVariables) + 1)
        
        With mudtVariables(UBound(mudtVariables))
            .strVarLocation = mudtCurLoc.strVarLocation
            .strVarProc = mudtCurLoc.strVarName
            .enmVarMode = lngMode
            .strVarName = Trim(Left(strLine, InStr(strLine, " ")))
        End With
    End If
    
    'find uncounted variables and notify programmer
    If UBound(mudtVariables) <> (NumVariables + 1 + lngCommaCount) Then
        With mudtVariables(UBound(mudtVariables))
            'Debug.Print .strVarLocation, .strVarProc, .strVarName
            'Stop
            'NumVariables = NumVariables - 1
        End With
    End If
    Exit Sub
    
errGetVarNames:
    Beep
    MsgBox Err.Description, vbExclamation, "Get Variable Names"
    Err.Clear
End Sub

Private Sub UpdateVars(ByVal strLine As String)
    'This will remove any variables from the array
    'if they are found within the specified string
    
    'first check private level variables
    Call UpdateByLevel(strLine, varPrivate)
    Call UpdateByLevel(strLine, varStatic)
    
    'next check module level variables
    Call UpdateByLevel(strLine, varModule)
    
    'check public level variables last
    Call UpdateByLevel(strLine, varPublic)
    Call UpdateByLevel(strLine, varGlobal)
End Sub

Private Sub UpdateByLevel(ByVal strLine As String, _
                          ByVal lngVarLevel As VarModeEnum)
    'This will remove any variable in the array
    'that appears in the string if it is a specified
    'level
    
    Dim lngCounter As Long      'used to cycle through the array
    Dim lngNumVars As Long      'the number of elements in the array
    Dim lngNumDel As Long       'the number of array elements deleted
    
    On Error GoTo errUpdateByLevel
    'get the number of variables in the array
    lngNumVars = UBound(mudtVariables)
    
    'search through the array
    For lngCounter = 0 To (lngNumVars - lngNumDel)
        'if we are deleting values, then we need to
        'move the array elements down
        If (lngCounter > (lngNumVars - lngNumDel)) Then
            Exit For
        End If
        mudtVariables(lngCounter) = mudtVariables(lngCounter + lngNumDel)
        
        With mudtVariables(lngCounter)
            'check to see if the variable is already used
            If (Not .blnVarUsed) And (.enmVarMode = lngVarLevel) Then
                'check if the variable is in the string
                If IsWord(strLine, .strVarName) Then
                    'the word is use, set the flag
                    .blnVarUsed = True
                    lngNumDel = lngNumDel + 1
                    lngCounter = lngCounter - 1
                End If
            End If
        End With
    Next lngCounter
    
    'resize the array to remove unwanted array
    'elements
    ReDim Preserve mudtVariables(lngNumVars - lngNumDel)
    Exit Sub
    
errUpdateByLevel:
    Beep
    MsgBox Err.Description, vbExclamation, "Update By Level"
    Err.Clear
End Sub

Private Sub ShowUnusedVars()
    'This will display a list of unused variables and
    'their location
    
    Dim lngVarCount As Long         'the size of the array of variable names
    Dim lngCounter As Long          'used to cycle through the array
    Dim lngNumUnused As Long        'the number of unused variables
    
    On Error GoTo errShowUnusedVars
    'get the total number of variables
    lngVarCount = UBound(mudtVariables)
    
    lstVars.Clear
    
    For lngCounter = 0 To lngVarCount
        With mudtVariables(lngCounter)
            If (Not .blnVarUsed) _
               And (.strVarLocation <> "") _
               And (.strVarName <> "") Then
                'display variable in the list box
                Call lstVars.AddItem(PaddString(.strVarLocation, 30) _
                                     & " " & _
                                     PaddString(.strVarProc, 30) _
                                     & " " & _
                                     PaddString(.strVarName, 30))
                lngNumUnused = lngNumUnused + 1
            End If
        End With
    Next lngCounter
    
    If lngNumUnused = 0 Then
        fraVariables.Enabled = False
    Else
        fraVariables.Enabled = True
    End If
    
    'display the number of unused variables
    fraVariables.Caption = "Unused Variables : " & Format(lngNumUnused, "##,##0")
    Exit Sub
    
errShowUnusedVars:
    Beep
    MsgBox Err.Description, vbExclamation, "Show Unused Variables"
    Err.Clear
End Sub

Private Function PaddString(ByVal strText As String, _
                            ByVal lngTotalChar As Long) _
                            As String
    'This will padd a string with trailing spaces so
    'that the returned string matches the total
    'number of characters specified. If the string
    'passed is bigger than the total number of
    'characters, then the string is truncated and then
    'returned.
    
    Dim lngLenText As Long  'the length of the text string passed
    
    'if the number of characters is zero, then
    'return nothing
    On Error GoTo errPaddString
    If lngTotalChar = 0 Then
        PaddString = ""
        Exit Function
    End If
    
    'get the length of the string
    lngLenText = Len(strText)
    
    If lngLenText >= lngTotalChar Then
        'return a trucated string
        PaddString = Left(strText, lngTotalChar)
    Else
        'padd the string out with spaces
        PaddString = strText & Space(lngTotalChar - lngLenText)
    End If
    Exit Function
    
errPaddString:
    Beep
    MsgBox Err.Description, vbExclamation, "Padd String"
    Err.Clear
End Function

Private Function IsWord(ByVal strLine As String, _
                        ByVal strWord As String) _
                        As Boolean
    'This function will return True if the
    'specified word is not part of another
    'word
    
    Dim blnLeftOk As Boolean    'the left side of the word is valid
    Dim blnRightOk As Boolean   'the right side of the word is valid
    Dim lngWordPos As Long      'the position of the specified word in the string
    
    If (Len(strWord) > Len(strLine)) _
       Or (strLine = "") _
       Or (strWord = "") Then
        'invalid parameters
        IsWord = False
        Exit Function
    End If
    
    On Error GoTo errIsWord
    lngWordPos = InStr(UCase(strLine), UCase(strWord))
    
    If lngWordPos = 0 Then
        'word not found
        IsWord = False
        Exit Function
    End If
    
    'check the left side of the word
    If lngWordPos = 1 Then
        'word is on the left side of the string
        blnLeftOk = True
    Else
        'check the character to the left of the word
        Select Case UCase(Mid(strLine, lngWordPos - 1, 1))
        Case "A" To "Z", "0" To "9"
        Case Else
            blnLeftOk = True
        End Select
    End If
    
    'check the right side of the word
    If (lngWordPos + Len(strWord)) = Len(strLine) Then
        'word is on the left side of the string
        blnLeftOk = True
    Else
        'check the character to the left of the word
        'Debug.Print strWord, UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Select Case UCase(Mid(strLine, lngWordPos + Len(strWord), 1))
        Case "A" To "Z", "0" To "9"
            'Stop
        Case Else
            blnRightOk = True
        End Select
    End If
    
    'if both sides are OK, then return True
    If blnLeftOk And blnRightOk Then
        IsWord = True
    End If
    Exit Function
    
errIsWord:
    Beep
    MsgBox Err.Description, vbExclamation, "Is Word"
    Err.Clear
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    m_iFormNo = 0
    frmMDI.Toolbar1.Buttons(11).Enabled = False
    Set frmStats = Nothing
End Sub

Private Sub FileFind()
    'scan for unused variables
    
    'if invalid path, then exit
    If (txtPath.Text = "") _
       Or (Dir(txtPath.Text) = "") Then
       Exit Sub
    End If
    
    'display the progress bar
    If GetTotal > 0 Then
        'find unused variables
        pbrVariables.Max = GetTotal
        pbrVariables.Visible = True
        mblnScanning = True
        
        Call ReadProject(txtPath.Text)
        
        'hide the progress bar
        pbrVariables.Visible = False
        mblnScanning = False
    End If
End Sub
