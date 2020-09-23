Attribute VB_Name = "TreeViewGradient"
'---Bas module code---
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type
Private Declare Function BeginPaint Lib "user32" _
    (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" _
    (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Type TRIVERTEX
    x As Long
    y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type
    
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type

Const GRADIENT_FILL_TRIANGLE As Long = &H2

Private Declare Function GradientFillTri Lib "msimg32" Alias "GradientFill" (ByVal hDC As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_TRIANGLE, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC& Lib "user32" (ByVal hWnd As Long)
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ValidateRectBynum& Lib "user32" Alias "ValidateRect" (ByVal hWnd As Long, ByVal lpRect As Long)
Declare Function ReleaseDC& Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long)

Private Const GWL_WNDPROC = (-4)
Private Const WM_PAINT = &HF
Private Const WM_ERASEBKGND = &H14
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_SETREDRAW = &HB
Dim vert(3) As TRIVERTEX
Dim gTri(1) As GRADIENT_TRIANGLE
Dim OldProc As Long, bPainting As Boolean
Dim TVWidth As Long, TVHeight As Long

Public Sub SubClass(obj As Object)
   Dim h As Long
   On Error Resume Next
   h = obj.hWnd
   If Err Or (OldProc <> 0) Then Exit Sub
   PrepareVertex obj
   OldProc = SetWindowLong(h, GWL_WNDPROC, AddressOf WndProc)
End Sub

Public Sub UnSubClass(obj As Object)
   Dim h As Long
   On Error Resume Next
   h = obj.hWnd
   If Err Or (OldProc = 0) Then Exit Sub
   SetWindowLong h, GWL_WNDPROC, OldProc
   OldProc = 0
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim TVDC As Long, TempDC As Long
   Dim oldBMP As Long, TempBMP As Long
   Dim ps As PAINTSTRUCT
   Select Case wMsg
          Case WM_PAINT
               If bPainting = False Then
                     BeginPaint hWnd, ps
                     bPainting = True
                     TVDC = GetDC(hWnd)
                     TempDC = CreateCompatibleDC(TVDC)
                     TempBMP = CreateCompatibleBitmap(TVDC, TVWidth, TVHeight)
                     oldBMP = SelectObject(TempDC, TempBMP)
                     SendMessage hWnd, WM_PAINT, TempDC, ByVal 0&
                     GradientFillTri TVDC, vert(0), 4, gTri(0), 2, GRADIENT_FILL_TRIANGLE
                     TransparentBlt TVDC, 0, 0, TVWidth, TVHeight, TempDC, 0, 0, TVWidth, TVHeight, TranslateColor(vbWindowBackground)
                     SelectObject TempDC, oldBMP
                     DeleteObject TempBMP
                     ReleaseDC hWnd, TempDC
                     ReleaseDC hWnd, TVDC
                     WndProc = 0
                     bPainting = False
                     EndPaint hWnd, ps
                     Exit Function
               End If
           Case WM_ERASEBKGND
                WndProc = 1
                Exit Function
           Case WM_HSCROLL, WM_VSCROLL, WM_MOUSEWHEEL
                InvalidateRect hWnd, 0, False
           Case Else
   End Select
   WndProc = CallWindowProc(OldProc, hWnd, wMsg, wParam, lParam)
End Function

Private Sub PrepareVertex(tv As Object)
'!!!Play with colors!!!
TVWidth = tv.Width \ Screen.TwipsPerPixelX
TVHeight = tv.Height \ Screen.TwipsPerPixelY

With vert(0)
    .x = 0
    .y = 0
    .Red = 0&
    .Green = LongToUShort(&HFF00&) '0
    .Blue = 0&
    .Alpha = 0&
End With
With vert(1)
    .x = TVWidth
    .y = 0
    .Red = 0 'LongToUShort(&HFF00&)
    .Green = 0&
    .Blue = LongToUShort(&HFF00&)
    .Alpha = 0&
End With
With vert(2)
    .x = TVWidth
'    .x = Me.ScaleWidth
    .y = TVHeight
    .Red = LongToUShort(&HFF00&)
    .Green = 0&
    .Blue = 0 'LongToUShort(&HFF00&)
    .Alpha = 0&
End With
With vert(3)
    .x = 0
    .y = TVHeight
    .Red = 0 'LongToUShort(&HFF00&)
    .Green = LongToUShort(&HFF00&)
    .Blue = LongToUShort(&HFF00&)
    .Alpha = 0&
End With
gTri(0).Vertex1 = 0
gTri(0).Vertex2 = 1
gTri(0).Vertex3 = 2

gTri(1).Vertex1 = 0
gTri(1).Vertex2 = 2
gTri(1).Vertex3 = 3
End Sub

Private Function LongToUShort(ULong As Long) As Integer
   LongToUShort = CInt(ULong - &H10000)
End Function

Private Function TranslateColor(inCol As Long) As Long
   Dim retCol As Long
   OleTranslateColor inCol, 0&, retCol
   TranslateColor = retCol
End Function


