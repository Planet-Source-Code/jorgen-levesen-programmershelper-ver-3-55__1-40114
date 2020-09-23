VERSION 5.00
Begin VB.Form frmPicViewer 
   BackColor       =   &H00404040&
   ClientHeight    =   7560
   ClientLeft      =   1140
   ClientTop       =   1230
   ClientWidth     =   10695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7560
   ScaleWidth      =   10695
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar vbar 
      Height          =   1335
      Left            =   4080
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar hbar 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.PictureBox picScroller 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   1155
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   120
      Width           =   1575
      Begin VB.PictureBox picImage 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1455
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.ComboBox PatternCombo 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.DirListBox DirList 
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.FileListBox FileList 
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "frmPicViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUser As Recordset
Private Sub LoadBackground()
    Picture1.Visible = False
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.Picture = frmMDI.Picture1.Picture
    TileForm Me, Picture1
End Sub
Private Sub DirList_Change()
    FileList.Path = DirList.Path
End Sub

Private Sub DriveList_Change()
    On Error GoTo DriveError
    DirList.Path = DriveList.Drive
    Exit Sub

DriveError:
    DriveList.Drive = DirList.Path
    Err.Clear
    Exit Sub
End Sub


Private Sub FileList_Click()
Dim fname As String

    On Error GoTo LoadPictureError

    fname = FileList.Path & "\" & FileList.FileName
    Caption = "Viewer [" & fname & "]"
    
    MousePointer = vbHourglass
    DoEvents
    picImage.Picture = LoadPicture(fname)
    MousePointer = vbDefault
    
    Exit Sub

LoadPictureError:
    Beep
    MousePointer = vbDefault
    Caption = "Viewer [Invalid picture]"
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set rsUser = m_dbPrograming.OpenRecordset("User")
    
    PatternCombo.AddItem "Bitmaps (*.bmp)"
    PatternCombo.AddItem "GIF (*.gif)"
    PatternCombo.AddItem "JPEG (*.jpg)"
    PatternCombo.AddItem "Icons (*.ico)"
    PatternCombo.AddItem "Matafiles (*.wmf)"
    PatternCombo.AddItem "DIBs (*.dib)"
    PatternCombo.AddItem "Graphic (*.gif;*.jpg;*.ico;*.bmp;*.wmf;*.dib)"
    PatternCombo.AddItem "All Files (*.*)"
    PatternCombo.ListIndex = 0

    DriveList.Drive = App.Path
    DirList.Path = App.Path
End Sub

' Set scroll bar parameters if necessary.
Private Sub ArrangeScrollbars()
Dim need_hgt As Single
Dim need_wid As Single
Dim got_hgt As Single
Dim got_wid As Single
Dim need_hbar As Boolean
Dim need_vbar As Boolean

    ' See which scroll bars we need.
    On Error Resume Next
    need_wid = picImage.Width
    need_hgt = picImage.Height
    got_wid = ScaleWidth - picScroller.Left
    got_hgt = ScaleHeight - picScroller.Top

    ' See if we need the horizontal scroll bar.
    need_hbar = (got_wid < need_wid)
    If need_hbar Then
        got_hgt = got_hgt - hbar.Height
    End If

    ' See if we need the vertical scroll bar.
    need_vbar = (got_hgt < need_hgt)
    If need_vbar Then
        got_wid = got_wid - vbar.Width

        ' See if we did not need the horizontal
        ' scroll bar but we now do.
        If Not need_hbar Then
            need_hbar = (got_wid < need_wid)
            If need_hbar Then
                got_hgt = got_hgt - hbar.Height
            End If
        End If
    End If
    If got_hgt < 120 Then got_hgt = 120
    If got_wid < 120 Then got_wid = 120

    ' Display the needed scroll bars.
    If need_hbar Then
        hbar.Move picScroller.Left, got_hgt, got_wid
        hbar.Min = 0
        hbar.Max = got_wid - need_wid
        hbar.SmallChange = got_wid / 3
        hbar.LargeChange = _
            (hbar.Max - hbar.Min) _
                * need_wid / _
                (got_wid - need_wid)
        hbar.Visible = True
    Else
        hbar.Visible = False
    End If
    If need_vbar Then
        vbar.Move picScroller.Left + got_wid, 0, vbar.Width, got_hgt
        vbar.Min = 0
        vbar.Max = got_hgt - need_hgt
        vbar.SmallChange = got_hgt / 3
        vbar.LargeChange = _
            (vbar.Max - vbar.Min) _
                * need_hgt / _
                (got_hgt - need_hgt)
        vbar.Visible = True
    Else
        vbar.Visible = False
    End If

    ' Arrange the window.
    picScroller.Move picScroller.Left, 0, got_wid, got_hgt
End Sub


Private Sub Form_Resize()
Const GAP = 60

Dim wid As Integer
Dim Hgt As Integer

    If WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    wid = DriveList.Width
    DriveList.Move GAP, GAP, wid
    PatternCombo.Move GAP, ScaleHeight - PatternCombo.Height, wid
    
    Hgt = (PatternCombo.Top - DriveList.Top - DriveList.Height - 3 * GAP) / 2
    If Hgt < 100 Then Hgt = 100
    DirList.Move GAP, DriveList.Top + DriveList.Height + GAP, wid, Hgt
    FileList.Move GAP, DirList.Top + DirList.Height + GAP, wid, Hgt

    ArrangeScrollbars
    LoadBackground
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    rsUser.Close
    Set frmPicViewer = Nothing
End Sub
Private Sub hbar_Change()
    picImage.Left = hbar.Value
End Sub

Private Sub hbar_Scroll()
    picImage.Left = hbar.Value
End Sub


Private Sub PatternCombo_Click()
Dim pat As String
Dim p1 As Integer
Dim p2 As Integer
    On Error Resume Next
    pat = PatternCombo.List(PatternCombo.ListIndex)
    p1 = InStr(pat, "(")
    p2 = InStr(pat, ")")
    FileList.Pattern = Mid$(pat, p1 + 1, p2 - p1 - 1)
End Sub


Private Sub vbar_Change()
    On Error Resume Next
    picImage.Top = vbar.Value
End Sub


Private Sub vbar_Scroll()
    On Error Resume Next
    picImage.Top = vbar.Value
End Sub


