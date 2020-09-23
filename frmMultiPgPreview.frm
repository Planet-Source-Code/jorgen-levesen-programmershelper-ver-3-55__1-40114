VERSION 5.00
Begin VB.Form frmMultiPgPreview 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   ClientHeight    =   5415
   ClientLeft      =   3405
   ClientTop       =   2085
   ClientWidth     =   4755
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   Icon            =   "frmMultiPgPreview.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5415
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPrintPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   3885
      ScaleHeight     =   435
      ScaleWidth      =   255
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   4200
      ScaleHeight     =   5145
      ScaleWidth      =   555
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   555
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "&Goto"
         Height          =   240
         Left            =   45
         TabIndex        =   4
         ToolTipText     =   "Goto Page"
         Top             =   2595
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   285
         Picture         =   "frmMultiPgPreview.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Next Page"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   45
         Picture         =   "frmMultiPgPreview.frx":00C6
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Prev. Page"
         Top             =   2220
         UseMaskColor    =   -1  'True
         Width           =   225
      End
      Begin VB.CommandButton cmd_quit 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         Height          =   705
         Left            =   30
         Picture         =   "frmMultiPgPreview.frx":0180
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Close"
         Top             =   735
         Width           =   495
      End
      Begin VB.CommandButton cmd_print 
         Caption         =   "Print"
         Height          =   675
         Left            =   30
         Picture         =   "frmMultiPgPreview.frx":057F
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Send to Printer"
         Top             =   45
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2220
         LargeChange     =   10
         Left            =   105
         Max             =   100
         Min             =   -20
         TabIndex        =   5
         Top             =   2910
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "P 1"
         Height          =   600
         Left            =   45
         TabIndex        =   16
         Top             =   1500
         Width           =   465
      End
   End
   Begin VB.PictureBox picHScroll 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   4755
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5145
      Visible         =   0   'False
      Width           =   4755
      Begin VB.HScrollBar HScroll1 
         Height          =   270
         Left            =   0
         Max             =   100
         TabIndex        =   6
         Top             =   0
         Width           =   3765
      End
   End
   Begin VB.PictureBox picPrintOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   2355
      Left            =   555
      ScaleHeight     =   2325
      ScaleWidth      =   3150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   3180
      Begin VB.TextBox txtFrom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1695
         TabIndex        =   10
         Text            =   "1"
         Top             =   1095
         Width           =   420
      End
      Begin VB.TextBox txtTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2475
         TabIndex        =   11
         Text            =   "1"
         Top             =   1095
         Width           =   420
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Ok"
         Height          =   360
         Left            =   2145
         TabIndex        =   13
         Top             =   1815
         Width           =   705
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   270
         Picture         =   "frmMultiPgPreview.frx":0978
         Top             =   450
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy page to clipboard"
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
         Height          =   300
         Index           =   0
         Left            =   585
         TabIndex        =   7
         Top             =   480
         Width           =   2250
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Current Page"
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
         Height          =   300
         Index           =   1
         Left            =   585
         TabIndex        =   8
         Top             =   810
         Width           =   1965
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   270
         Picture         =   "frmMultiPgPreview.frx":0A15
         Top             =   780
         Width           =   300
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   2
         Left            =   270
         Picture         =   "frmMultiPgPreview.frx":0AB2
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print Pages"
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
         Height          =   300
         Index           =   2
         Left            =   585
         TabIndex        =   9
         Top             =   1110
         Width           =   1965
      End
      Begin VB.Image optPrint 
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   3
         Left            =   270
         Picture         =   "frmMultiPgPreview.frx":0B4F
         Top             =   1410
         Width           =   300
      End
      Begin VB.Label optText 
         BackStyle       =   0  'Transparent
         Caption         =   "Print All"
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
         Height          =   300
         Index           =   3
         Left            =   585
         TabIndex        =   12
         Top             =   1440
         Width           =   1965
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Height          =   270
         Left            =   2175
         TabIndex        =   21
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label lblPrintingPg 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   255
         TabIndex        =   20
         Top             =   1995
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   315
         Left            =   135
         TabIndex        =   18
         Top             =   30
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF00FF&
         Height          =   2250
         Left            =   30
         Top             =   30
         Width           =   3090
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4845
      Left            =   0
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   3765
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   0
      Picture         =   "frmMultiPgPreview.frx":0BEC
      Top             =   4860
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image optArt 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   555
      Picture         =   "frmMultiPgPreview.frx":0C99
      Top             =   4875
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "frmMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1999-2002
'/*************************************/
Option Explicit

Public PageNumber As Integer
Private ViewPage As Integer
Private TempDir As String
Private OptionV As Integer

Private Type PanState
   x As Long
   y As Long
End Type
Dim PanSet As PanState
Dim rsLanguage As Recordset
Private Sub ReadText()
    'find YOUR sLanguage text
    With rsLanguage
        .MoveFirst
        Do While Not .EOF
            If .Fields("Language") = m_FileExt Then
                .Edit
                If IsNull(.Fields("label2")) Then
                    .Fields("label2") = Label2.Caption
                Else
                    Label2.Caption = .Fields("label2")
                End If
                If IsNull(.Fields("label3")) Then
                    .Fields("label3") = Label3.Caption
                Else
                    Label3.Caption = .Fields("label3")
                End If
                If IsNull(.Fields("optText(0)")) Then
                    .Fields("optText(0)") = optText(0).Caption
                Else
                    optText(0).Caption = .Fields("optText(0)")
                End If
                If IsNull(.Fields("optText(1)")) Then
                    .Fields("optText(1)") = optText(1).Caption
                Else
                    optText(1).Caption = .Fields("optText(1)")
                End If
                If IsNull(.Fields("optText(2)")) Then
                    .Fields("optText(2)") = optText(2).Caption
                Else
                    optText(2).Caption = .Fields("optText(2)")
                End If
                If IsNull(.Fields("optText(3)")) Then
                    .Fields("optText(3)") = optText(3).Caption
                Else
                    optText(3).Caption = .Fields("optText(3)")
                End If
                If IsNull(.Fields("cmdPrint")) Then
                    .Fields("cmdPrint") = cmdPrint.Caption
                Else
                    cmdPrint.Caption = .Fields("cmdPrint")
                End If
                If IsNull(.Fields("cmd_quit")) Then
                    .Fields("cmd_quit") = cmd_quit.Caption
                Else
                    cmd_quit.Caption = .Fields("cmd_quit")
                End If
                If IsNull(.Fields("cmd_print")) Then
                    .Fields("cmd_print") = cmd_print.Caption
                Else
                    cmd_print.Caption = .Fields("cmd_print")
                End If
                .Update
                Exit Sub
             End If
        .MoveNext
        Loop
                
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("label2") = Label2.Caption
        .Fields("label3") = Label3.Caption
        .Fields("optText(0)") = optText(0).Caption
        .Fields("optText(1)") = optText(1).Caption
        .Fields("optText(2)") = optText(2).Caption
        .Fields("optText(3)") = optText(3).Caption
        .Fields("cmdPrint") = cmdPrint.Caption
        .Fields("cmd_quit") = cmd_quit.Caption
        .Fields("cmd_print") = cmd_print.Caption
        .Update
    End With
End Sub
Private Sub cmd_print_Click()
    txtTo.Text = PageNumber + 1
    OptionV = 3
    optPrint(3).Picture = optArt(1).Picture
    picPrintOptions.Left = Me.Width - (Picture2.Width + picPrintOptions.Width + 50)
    picPrintOptions.Visible = True
End Sub

Private Function IsNumber(ByVal CheckString As String, Optional KeyAscii As Integer = 0, Optional AllowDecPoint As Boolean = False, Optional AllowNegative As Boolean = False) As Boolean
    If KeyAscii > 0 And KeyAscii <> 8 Then
        If Not AllowNegative And KeyAscii = 45 Then KeyAscii = 0
        If Not AllowDecPoint And KeyAscii = 46 Then KeyAscii = 0
        If Not IsNumeric(CheckString & Chr(KeyAscii)) Then
            KeyAscii = False
            IsNumber = False
        Else
            IsNumber = True
        End If
    Else
        IsNumber = IsNumeric(CheckString)
    End If
End Function

Private Sub cmd_quit_Click()
    cPrint.SendToPrinter = False
    Unload Me
End Sub

Private Sub cmdGoTo_Click()
  Dim NewPageNo As Variant
    On Local Error Resume Next
    
    
    cmd_print.SetFocus
    
    NewPageNo = InputBox("Enter page number", "GoTo Page", 1)
    NewPageNo = Val(NewPageNo)
    
    If NewPageNo = 0 Then Exit Sub
    
    NewPageNo = NewPageNo - 1
    If NewPageNo > PageNumber Then NewPageNo = PageNumber
    ViewPage = NewPageNo
        
    Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    
    picPrintOptions.Visible = False
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages

End Sub

Private Sub cmdPrint_Click()
  Dim i As Integer
  
    '/* Prevent printing again until done
    cmd_print.SetFocus
    picPrintOptions.Enabled = False
    lblPrintingPg.Visible = True
    cmdPrint.Visible = False
    
    Select Case OptionV
    Case 0 '/* Copy to clipboard
        Clipboard.Clear
        Clipboard.SetData Picture1.Picture, vbCFBitmap
    Case 1 '/* Print current page
        lblPrintingPg.Caption = "Printing page " & ViewPage + 1
        lblPrintingPg.Refresh
        Call PrintPictureBox(Picture1, True, False)
    Case 2 '/* Print range
        For i = Val(txtFrom) - 1 To Val(txtTo) - 1
            lblPrintingPg.Caption = "Printing page " & CStr(i + 1) & " of " & txtTo
            lblPrintingPg.Refresh
            Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(i) & ".bmp")
            Call PrintPictureBox(Picture1, True, False)
        Next i
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Case Else '/* Print all
        cPrint.SendToPrinter = True '/* Send to Printer */
        Unload Me
    End Select
    
    '/* Restore normal view
    picPrintOptions.Enabled = True
    cmdPrint.Visible = True
    picPrintOptions.Visible = False
    lblPrintingPg.Visible = False

End Sub

Private Sub Command1_Click(Index As Integer)
    On Local Error Resume Next
    If Index = 0 Then
        ViewPage = ViewPage - 1
        If ViewPage < 0 Then ViewPage = 0
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    Else
        ViewPage = ViewPage + 1
        If ViewPage > PageNumber Then ViewPage = PageNumber
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
    End If
    
    Picture1.Top = 0
    'Picture1.Refresh
    picPrintOptions.Visible = False
    VScroll1.Value = 0
    HScroll1.Value = 0
    Call DisplayPages
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Screen.MousePointer = vbDefault
    Call DisplayPages
    If Picture1.Width < Me.Width - Picture2.Width Then
        Picture1.Move ((Me.Width - Picture2.Width) - Picture1.Width) \ 2, 0
    End If
    ReadText
End Sub

Private Sub Form_Click()
    picPrintOptions.Visible = False
End Sub


Private Sub Form_Load()
    Me.Move 0, 0, Screen.Width, Screen.Height
    Picture1.Move 0, 0

    VScroll1.Height = Me.Height - cmdGoTo.Top - cmdGoTo.Height - 500
    HScroll1.Width = Me.Width - Picture2.Width - 500
    
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmMultiPgPreview")
    TempDir = Environ("TEMP") & "\"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim tFileName As String
    On Error Resume Next
    '/* Remove preview pages
    tFileName = Dir(TempDir & "PPview*.bmp")
    If tFileName > vbNullString Then
        Do
            Kill TempDir & tFileName
            tFileName = Dir(TempDir & "PPview*.bmp")
        Loop Until tFileName = vbNullString
    End If
    
    PageNumber = 0
    ViewPage = 0
    rsLanguage.Close
    Set frmMultiPgPreview = Nothing
End Sub


Private Sub HScroll1_Change()
    On Local Error Resume Next
    Picture1.Left = -(HScroll1.Value)
    HScroll1.SetFocus
    On Local Error GoTo 0
End Sub

Private Sub HScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 37, 33 '/* Arrow left
        'Call Command1_Click(0)
    Case 39, 34 '/* Arrow right
        'Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select

End Sub


Private Sub optPrint_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 3
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub

Private Sub optText_Click(Index As Integer)
  Dim i As Byte
    OptionV = Index
    For i = 0 To 3
        If Index = i Then
            optPrint(i).Picture = optArt(1).Picture
        Else
            optPrint(i).Picture = optArt(0).Picture
        End If
    Next i

End Sub


Private Sub Picture1_Click()
    picPrintOptions.Visible = False
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 38 '/* Arrow up
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    Case 40 '/* Arrow down
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    Case 37, 33 '/* Arrow left
        Call Command1_Click(0)
    Case 39, 34 '/* Arrow right
        Call Command1_Click(1)
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
        
    End Select
    
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      PanSet.x = x
      PanSet.y = y
      MousePointer = vbSizePointer
   End If
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim nTop As Integer, nLeft As Integer

   On Local Error Resume Next

   If Button = vbLeftButton And Shift = 0 Then

      '/* new coordinates?
      With Picture1
         nTop = -(.Top + (y - PanSet.y))
         nLeft = -(.Left + (x - PanSet.x))
      End With

      '/* Check limits
      With VScroll1
         If .Visible Then
            If nTop < .Min Then
               nTop = .Min
            ElseIf nTop > .Max Then
               nTop = .Max
            End If
         Else
            nTop = -Picture1.Top
         End If
      End With

      With HScroll1
         If .Visible Then
            If nLeft < .Min Then
               nLeft = .Min
            ElseIf nLeft > .Max Then
               nLeft = .Max
            End If
         Else
            nLeft = -Picture1.Left
         End If
      End With

      Picture1.Move -nLeft, -nTop

   End If

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   On Local Error Resume Next
   If Button = vbLeftButton And Shift = 0 Then
      If VScroll1.Visible Then VScroll1.Value = -(Picture1.Top)
      If HScroll1.Visible Then HScroll1.Value = -(Picture1.Left)
   End If
   MousePointer = vbDefault
End Sub


Private Sub txtFrom_Change()
    If Val(txtFrom) < 1 Then txtFrom = 1
    If Val(txtFrom) > Val(txtTo) Then txtFrom = txtTo
End Sub

Private Sub txtFrom_GotFocus()
    txtFrom.SelStart = 0
    txtFrom.SelLength = Len(txtFrom)
End Sub


Private Sub txtFrom_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtFrom = txtFrom + 1
        KeyCode = False
    Case 40  '/* "-"
        txtFrom = txtFrom - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
    IsNumber txtFrom, KeyAscii, False, False
End Sub


Private Sub txtTo_Change()
    If Val(txtTo) > PageNumber + 1 Then txtTo = PageNumber + 1
    If Val(txtTo) < Val(txtFrom) Then txtTo = txtFrom
End Sub

Private Sub txtTo_GotFocus()
    txtTo.SelStart = 0
    txtTo.SelLength = Len(txtTo)
End Sub


Private Sub txtTo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 38  '/* "+"
        txtTo = txtTo + 1
        KeyCode = False
    Case 40  '/* "-"
        txtTo = txtTo - 1
        KeyCode = False
    End Select
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
    IsNumber txtTo, KeyAscii, False, False
End Sub


Private Sub VScroll1_Change()
    On Local Error Resume Next
    Picture1.Top = -(VScroll1.Value)
    VScroll1.SetFocus
    On Local Error GoTo 0
End Sub


Private Sub DisplayPages()
    Label1 = CStr(ViewPage + 1) & vbNewLine & "-- of --" & vbNewLine & CStr(PageNumber + 1)
    
    If Picture1.Width > Me.Width - Picture2.Width Then
        picHScroll.Visible = True
    Else
        picHScroll.Visible = False
    End If

    If Picture1.Height >= Me.Height Then
        VScroll1.Visible = True
    Else
        VScroll1.Visible = False
    End If
    Picture1.SetFocus

End Sub
Private Sub PrintPictureBox(pBox As PictureBox, _
                           Optional ScaleToFit As Boolean = True, _
                           Optional MaintainRatio As Boolean = True)
 
 Dim xmin As Single
 Dim ymin As Single
 Dim wid As Single
 Dim Hgt As Single
 Dim aspect As Single
 
    Screen.MousePointer = vbHourglass
    
    If Not ScaleToFit Then
        wid = Printer.ScaleX(pBox.ScaleWidth, pBox.ScaleMode, Printer.ScaleMode)
        Hgt = Printer.ScaleY(pBox.ScaleHeight, pBox.ScaleMode, Printer.ScaleMode)
        xmin = (Printer.ScaleWidth - wid) / 2
        ymin = (Printer.ScaleHeight - Hgt) / 2
    Else
        aspect = pBox.ScaleHeight / pBox.ScaleWidth
        wid = Printer.ScaleWidth
        Hgt = Printer.ScaleHeight
        
        If MaintainRatio Then
            If Hgt / wid > aspect Then
                Hgt = aspect * wid
                xmin = Printer.ScaleLeft
                ymin = (Printer.ScaleHeight - Hgt) / 2
            Else
                wid = Hgt / aspect
                xmin = (Printer.ScaleWidth - wid) / 2
                ymin = Printer.ScaleTop
            End If
        End If
    End If
    
    Printer.PaintPicture pBox.Picture, xmin, ymin, wid, Hgt
    Printer.EndDoc

    Screen.MousePointer = vbDefault

End Sub


Private Sub VScroll1_KeyUp(KeyCode As Integer, Shift As Integer)
   On Local Error Resume Next
    Select Case KeyCode
    Case 37, 33 '/* Arrow left
        If HScroll1.Visible = False Then
            Call Command1_Click(0)
        Else
            HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
        End If
    Case 39, 34 '/* Arrow right
        If HScroll1.Visible = False Then
            Call Command1_Click(1)
        Else
            HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
        End If
    Case 71 '/* G
        Call cmdGoTo_Click
    Case 35, 36 '/* Home, End
      Dim NewPageNo As Long
        If KeyCode = 36 Then
            NewPageNo = 0
        Else
            NewPageNo = PageNumber
        End If
        ViewPage = NewPageNo
        Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(ViewPage) & ".bmp")
        picPrintOptions.Visible = False
        VScroll1.Value = 0
        HScroll1.Value = 0
        Call DisplayPages
    End Select
End Sub


