VERSION 5.00
Begin VB.Form frmPrinterSetUp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Printer Setup"
   ClientHeight    =   4890
   ClientLeft      =   3525
   ClientTop       =   2940
   ClientWidth     =   6690
   Icon            =   "frmPrinterSetUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   465
      Left            =   5220
      TabIndex        =   22
      Top             =   2745
      Width           =   1380
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   465
      Left            =   5220
      TabIndex        =   1
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Frame fraCopies 
      Caption         =   "Copies"
      Height          =   1515
      Left            =   3375
      TabIndex        =   18
      Top             =   2655
      Width           =   1710
      Begin VB.VScrollBar VScroll 
         Height          =   390
         Left            =   945
         Max             =   9
         Min             =   1
         TabIndex        =   20
         Top             =   375
         Value           =   1
         Width           =   270
      End
      Begin VB.TextBox txtCopies 
         Height          =   285
         Left            =   285
         TabIndex        =   19
         Text            =   "1"
         Top             =   435
         Width           =   615
      End
      Begin VB.Image imgCopies 
         Height          =   450
         Left            =   105
         Picture         =   "frmPrinterSetUp.frx":030A
         Top             =   945
         Width           =   1470
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   1215
      TabIndex        =   14
      Top             =   240
      Width           =   4845
   End
   Begin VB.TextBox txtDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   4860
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Width           =   4860
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   465
      Left            =   5220
      TabIndex        =   0
      Top             =   3225
      Width           =   1380
   End
   Begin VB.Frame fraQuality 
      Caption         =   "Quality"
      Height          =   1515
      Left            =   225
      TabIndex        =   8
      Top             =   2655
      Width           =   3045
      Begin VB.OptionButton optQuality 
         Caption         =   "Best"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optQuality 
         Caption         =   "Normal"
         Height          =   375
         Index           =   1
         Left            =   1050
         TabIndex        =   10
         Top             =   960
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.OptionButton optQuality 
         Caption         =   "Draft"
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Top             =   960
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   150
         Picture         =   "frmPrinterSetUp.frx":04CD
         Top             =   405
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1215
         Picture         =   "frmPrinterSetUp.frx":0620
         Top             =   405
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2280
         Picture         =   "frmPrinterSetUp.frx":07B9
         Top             =   405
         Width           =   480
      End
   End
   Begin VB.Frame fraDuplex 
      Caption         =   "Duplex"
      Height          =   1065
      Left            =   210
      TabIndex        =   3
      Top             =   1515
      Width           =   3045
      Begin VB.OptionButton optDuplex 
         Caption         =   "Double Sided Book"
         Height          =   225
         Index           =   2
         Left            =   885
         TabIndex        =   21
         Top             =   750
         Width           =   2100
      End
      Begin VB.OptionButton optDuplex 
         Caption         =   "Double Sided Tablet"
         Height          =   225
         Index           =   1
         Left            =   885
         TabIndex        =   7
         Top             =   480
         Width           =   2100
      End
      Begin VB.OptionButton optDuplex 
         Caption         =   "Single Sided"
         Height          =   225
         Index           =   0
         Left            =   885
         TabIndex        =   6
         Top             =   210
         Value           =   -1  'True
         Width           =   2100
      End
      Begin VB.Image imgDuplex 
         Height          =   300
         Index           =   2
         Left            =   300
         Picture         =   "frmPrinterSetUp.frx":0962
         Top             =   345
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgDuplex 
         Height          =   300
         Index           =   0
         Left            =   300
         Picture         =   "frmPrinterSetUp.frx":0A44
         Top             =   345
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgDuplex 
         Height          =   465
         Index           =   1
         Left            =   300
         Picture         =   "frmPrinterSetUp.frx":0B1E
         Top             =   345
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgPrinterDuplex 
         Height          =   300
         Left            =   300
         Picture         =   "frmPrinterSetUp.frx":0BF6
         Top             =   345
         Width           =   405
      End
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Orientation"
      Height          =   1065
      Left            =   3360
      TabIndex        =   2
      Top             =   1515
      Width           =   3285
      Begin VB.OptionButton optOrien 
         Caption         =   "Landscape"
         Height          =   255
         Index           =   1
         Left            =   1170
         TabIndex        =   5
         Top             =   705
         Width           =   2070
      End
      Begin VB.OptionButton optOrien 
         Caption         =   "Portrait"
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   4
         Top             =   375
         Value           =   -1  'True
         Width           =   1950
      End
      Begin VB.Image imgOrien 
         Height          =   480
         Index           =   0
         Left            =   240
         Picture         =   "frmPrinterSetUp.frx":0CD0
         Top             =   405
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Image imgOrien 
         Height          =   390
         Index           =   1
         Left            =   240
         Picture         =   "frmPrinterSetUp.frx":0DB8
         Top             =   405
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgPrinterOrien 
         Height          =   480
         Left            =   240
         Picture         =   "frmPrinterSetUp.frx":0E99
         Top             =   405
         Width           =   390
      End
   End
   Begin VB.Frame fraColor 
      Height          =   510
      Left            =   225
      TabIndex        =   23
      Top             =   4110
      Width           =   3045
      Begin VB.OptionButton optColor 
         Caption         =   "Grayscale"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   25
         Top             =   210
         Width           =   1200
      End
      Begin VB.OptionButton optColor 
         Caption         =   "Color"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   24
         Top             =   195
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Printer:"
      Height          =   375
      Index           =   0
      Left            =   255
      TabIndex        =   17
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      Height          =   255
      Index           =   1
      Left            =   255
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Port:"
      Height          =   255
      Index           =   2
      Left            =   255
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmPrinterSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1996-2002
'/*************************************/
Option Explicit
Const MaxCopies As Integer = 999
Dim PrinterName As String
Dim PrinterSetupFormLoaded As Boolean
Dim rsLanguage As Recordset
Private Sub ReadText()
    'find YOUR sLanguage text
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
                If IsNull(.Fields("lblLabels(0)")) Then
                    .Fields("lblLabels(0)") = lblLabels(0).Caption
                Else
                    lblLabels(0).Caption = .Fields("lblLabels(0)")
                End If
                If IsNull(.Fields("lblLabels(1)")) Then
                    .Fields("lblLabels(1)") = lblLabels(1).Caption
                Else
                    lblLabels(1).Caption = .Fields("lblLabels(1)")
                End If
                If IsNull(.Fields("lblLabels(2)")) Then
                    .Fields("lblLabels(2)") = lblLabels(2).Caption
                Else
                    lblLabels(2).Caption = .Fields("lblLabels(2)")
                End If
                If IsNull(.Fields("fraDuplex")) Then
                    .Fields("fraDuplex") = fraDuplex.Caption
                Else
                    fraDuplex.Caption = .Fields("fraDuplex")
                End If
                If IsNull(.Fields("fraOrientation")) Then
                    .Fields("fraOrientation") = fraOrientation.Caption
                Else
                    fraOrientation.Caption = .Fields("fraOrientation")
                End If
                If IsNull(.Fields("fraQuality")) Then
                    .Fields("fraQuality") = fraQuality.Caption
                Else
                    fraQuality.Caption = .Fields("fraQuality")
                End If
                If IsNull(.Fields("fraCopies")) Then
                    .Fields("fraCopies") = fraCopies.Caption
                Else
                    fraCopies.Caption = .Fields("fraCopies")
                End If
                If IsNull(.Fields("optDuplex(0)")) Then
                    .Fields("optDuplex(0)") = optDuplex(0).Caption
                Else
                    optDuplex(0).Caption = .Fields("optDuplex(0)")
                End If
                If IsNull(.Fields("optDuplex(1)")) Then
                    .Fields("optDuplex(1)") = optDuplex(1).Caption
                Else
                    optDuplex(1).Caption = .Fields("optDuplex(1)")
                End If
                If IsNull(.Fields("optDuplex(2)")) Then
                    .Fields("optDuplex(2)") = optDuplex(2).Caption
                Else
                    optDuplex(2).Caption = .Fields("optDuplex(2)")
                End If
                If IsNull(.Fields("optOrien(0)")) Then
                    .Fields("optOrien(0)") = optOrien(0).Caption
                Else
                    optOrien(0).Caption = .Fields("optOrien(0)")
                End If
                If IsNull(.Fields("optOrien(1)")) Then
                    .Fields("optOrien(1)") = optOrien(1).Caption
                Else
                    optOrien(1).Caption = .Fields("optOrien(1)")
                End If
                If IsNull(.Fields("optQuality(0)")) Then
                    .Fields("optQuality(0)") = optQuality(0).Caption
                Else
                    optQuality(0).Caption = .Fields("optQuality(0)")
                End If
                If IsNull(.Fields("optQuality(1)")) Then
                    .Fields("optQuality(1)") = optQuality(1).Caption
                Else
                    optQuality(1).Caption = .Fields("optQuality(1)")
                End If
                If IsNull(.Fields("optQuality(2)")) Then
                    .Fields("optQuality(2)") = optQuality(2).Caption
                Else
                    optQuality(2).Caption = .Fields("optQuality(2)")
                End If
                If IsNull(.Fields("optColor(0)")) Then
                    .Fields("optColor(0)") = optColor(0).Caption
                Else
                    optColor(0).Caption = .Fields("optColor(0)")
                End If
                If IsNull(.Fields("optColor(1)")) Then
                    .Fields("optColor(1)") = optColor(1).Caption
                Else
                    optColor(1).Caption = .Fields("optColor(1)")
                End If
                If IsNull(.Fields("cmdPreview")) Then
                    .Fields("cmdPreview") = cmdPreview.Caption
                Else
                    cmdPreview.Caption = .Fields("cmdPreview")
                End If
                If IsNull(.Fields("cmdPrint")) Then
                    .Fields("cmdPrint") = cmdPrint.Caption
                Else
                    cmdPrint.Caption = .Fields("cmdPrint")
                End If
                If IsNull(.Fields("cmdQuit")) Then
                    .Fields("cmdQuit") = cmdQuit.Caption
                Else
                    cmdQuit.Caption = .Fields("cmdQuit")
                End If
                .Update
                Exit Sub
             End If
        .MoveNext
        Loop
                
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("lblLabels(0)") = lblLabels(0).Caption
        .Fields("lblLabels(1)") = lblLabels(1).Caption
        .Fields("lblLabels(2)") = lblLabels(2).Caption
        .Fields("fraDuplex") = fraDuplex.Caption
        .Fields("fraOrientation") = fraOrientation.Caption
        .Fields("fraQuality") = fraQuality.Caption
        .Fields("fraCopies") = fraCopies.Caption
        .Fields("optDuplex(0)") = optDuplex(0).Caption
        .Fields("optDuplex(1)") = optDuplex(1).Caption
        .Fields("optDuplex(2)") = optDuplex(2).Caption
        .Fields("optOrien(0)") = optOrien(0).Caption
        .Fields("optOrien(1)") = optOrien(1).Caption
        .Fields("optQuality(0)") = optQuality(0).Caption
        .Fields("optQuality(1)") = optQuality(1).Caption
        .Fields("optQuality(2)") = optQuality(2).Caption
        .Fields("optColor(0))") = optColor(0).Caption
        .Fields("optColor(1))") = optColor(1).Caption
        .Fields("cmdPreview") = cmdPreview.Caption
        .Fields("cmdPrint") = cmdPrint.Caption
        .Fields("cmdPrint") = cmdPrint.Caption
        .Update
    End With
End Sub

Private Sub cboPrinter_Click()
  Dim xPrinter As Printer
    
    On Local Error Resume Next
    
    For Each xPrinter In Printers
        If xPrinter.DeviceName = cboPrinter.Text Then
            
            Set Printer = xPrinter
            
            txtDriver = Printer.DriverName
            PrinterName = cboPrinter.Text
            txtPort = Printer.Port
            
            optDuplex(Printer.Duplex - 1).Value = True
            If Printer.Orientation = vbPRORPortrait Then
                optOrien(1) = False
                optOrien(0) = True
            Else
                optOrien(0) = True
                optOrien(1) = False
            End If
            Exit For
            
        End If
    Next

End Sub

Private Sub cmdPreview_Click()
    cPrint.SendToPrinter = False
    Call PrintPreview
End Sub

Private Sub cmdPrint_Click()
    cPrint.SendToPrinter = True
    Call PrintPreview
End Sub

Private Sub cmdQuit_Click()
    QuitCommand = True
    Me.Hide
End Sub

Private Sub Form_Activate()
    'On Error Resume Next
    ReadText
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
 Dim xPrinter As Printer, Index As Integer
    
    On Local Error Resume Next
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmPrinterSetUp")
    
    VScroll.Max = MaxCopies
    VScroll.Min = 1
    
    PrinterName = GetSetting(App.Title, "Options", "Printer", "None")
    txtCopies = GetSetting(App.Title, "Options", "Copies", "1")
    
    Index = -1
    For Each xPrinter In Printers
        cboPrinter.AddItem xPrinter.DeviceName
        If xPrinter.DeviceName = PrinterName Then Index = cboPrinter.NewIndex
        If xPrinter.DeviceName = Printer.DeviceName And Index = -1 Then Index = cboPrinter.NewIndex
    Next
    If Index >= 0 Then cboPrinter.ListIndex = Index
    
    PrinterSetupFormLoaded = True
    DoEvents

End Sub

Private Sub Form_Paint()
    Me.ZOrder
    QuitCommand = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsLanguage.Close
    Set frmPrinterSetUp = Nothing
End Sub

Private Sub optOrien_Click(Index As Integer)
    On Local Error Resume Next
    
    Printer.Orientation = Index + 1
    If Err.Number Then
       optOrien(0).Value = True
       Index = False
    End If
    
    imgPrinterOrien.Picture = imgOrien(Index).Picture

End Sub

Private Sub optDuplex_Click(Index As Integer)
    If Not PrinterSetupFormLoaded Then Exit Sub
    imgPrinterDuplex.Picture = imgDuplex(Index).Picture
End Sub

Private Sub optQuality_Click(Index As Integer)
    On Local Error Resume Next
    Select Case Index
    Case 0
        Printer.PrintQuality = vbPRPQDraft
    Case 1
        Printer.PrintQuality = vbPRPQMedium
    Case Else
        Printer.PrintQuality = vbPRPQHigh
    End Select

End Sub

Private Sub txtCopies_Change()
    On Local Error Resume Next
    
    If Val(txtCopies) > MaxCopies Then
        txtCopies = Format(MaxCopies)
    ElseIf Val(txtCopies) < 1 Then
        txtCopies = "1"
    End If
    VScroll.Value = Val(txtCopies)
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = False
    End If
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub VScroll_Change()
    txtCopies = Abs(VScroll.Value)
End Sub

Private Sub PrintPreview()
  Dim i As Byte
    
    On Local Error Resume Next
    For i = 0 To 2
        If optDuplex(i).Value Then
            Select Case i
            Case 1 '/* Double Sided Tablet
                If Printer.Orientation = vbPRORPortrait Then
                    Printer.Duplex = vbPRDPVertical
                Else
                    Printer.Duplex = vbPRDPHorizontal
                End If
            Case 2 '/* Double Sided Book
                If Printer.Orientation = vbPRORPortrait Then
                    Printer.Duplex = vbPRDPHorizontal
                Else
                    Printer.Duplex = vbPRDPVertical
                End If
            Case Else '/* Single Sided
                Printer.Duplex = vbPRDPSimplex
            End Select
        End If
    Next i
    
    If optColor(1).Value Then
        Printer.ColorMode = vbPRCMMonochrome
    Else
        Printer.ColorMode = vbPRCMColor
    End If
        
    Printer.Copies = Val(txtCopies)
    SaveSetting App.Title, "Options", "Printer", PrinterName
    SaveSetting App.Title, "Options", "Copies", txtCopies
    QuitCommand = False
    Unload Me

End Sub
