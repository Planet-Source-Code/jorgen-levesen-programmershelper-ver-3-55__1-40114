VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMultiPgPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*************************************/
'/* Author: Morgan Haueisen
'/*         morganh@hartcom.net
'/* Copyright (c) 1998-2002
'/*************************************/
Option Explicit

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'/* Flag indicating Printing or Previewing
Private PrintFlag As Boolean

'/* Object used for Print Preview
Private ObjPrint As Control

'/* Storage for the Printer's orignal scale mode
Private pSM As Integer
'/* Storage for the Object's orignal scale mode
Private oSM As Integer

'/* The size ratio between the actual page and the print preview object
'/* This is only used for Option 1 (See: pStartDoc)
Private Ratio As Single

'/* The actual printable area (something a little less then the paper size)
Private PgWidth As Single
Private PgHeight As Single
Private BotMargin As Single

Private PageNumber As Integer
Private TempDir As String

Public Enum PageOrientation
    PagePortrait = vbPRORPortrait
    PageLandscape = vbPRORLandscape
End Enum
Public Property Let PrintCopies(pNumber As Integer)
    On Local Error Resume Next
    Printer.Copies = pNumber
    On Local Error GoTo 0
End Property

Public Property Get PrintCopies() As Integer
    On Local Error Resume Next
    PrintCopies = Printer.Copies
    On Local Error GoTo 0
End Property

Public Sub pCancled()
    FontSize = 12
    FontBold = True
    ForeColor = vbRed
    pPrint
    pPrint "**** PRINTING CANCLED ****", 0.5
End Sub

Public Sub pMultiline(ByVal PrintVar As Variant, _
                      Optional ByVal LeftMargin As Single = -1, _
                      Optional ByVal RightMargin As Single = -1, _
                      Optional ByVal IndentChar As String = vbNullString, _
                      Optional SameLine As Boolean = False, _
                      Optional UsePageBreaks As Boolean = False)
 
  Dim StartChar As Integer
  Dim SecondLine As Boolean
  Dim CharLength As Single
  Dim CurrentPos As Single
  Dim TxtLen As Single
  Dim TxtWidth As Single
  Dim IndentText As String
  Dim x As Integer, y As Integer
 
    If LeftMargin = -1 Then LeftMargin = CurrentX
    If RightMargin = -1 Then RightMargin = PgWidth - 0.1
    RightMargin = RightMargin - LeftMargin
    
    If IsNull(PrintVar) Or Len(PrintVar) = 0 Then PrintVar = " "
    TxtLen = Len(PrintVar)
    StartChar = 1
    CurrentPos = 0
    CharLength = TxtLen
    IndentText = vbNullString
 
    On Local Error Resume Next
    TxtWidth = GetTextWidth(PrintVar)
    If TxtWidth <= RightMargin And Err.Number = 0 Then
        CurrentX = LeftMargin
        If SameLine Then
            pPrint PrintVar, , True
        Else
            pPrint PrintVar
        End If
        Exit Sub
    End If
    Err.Clear
    On Local Error GoTo 0
    
    For x = 1 To TxtLen
        y = x - CurrentPos
        If Mid(PrintVar, x, 1) < Chr(33) Then CharLength = y
        If GetTextWidth(IndentText) + GetTextWidth(Mid(PrintVar, StartChar, y)) >= RightMargin Then
            CurrentX = LeftMargin
            '/* If there are no spaces then break line here */
            If CharLength > y Then CharLength = y - Len(IndentText) - 1
            
            If PrintFlag Then
                Printer.Print IndentText & Mid(PrintVar, StartChar, CharLength)
            Else
                ObjPrint.Print IndentText & Mid(PrintVar, StartChar, CharLength)
            End If
            
            CurrentPos = CharLength + CurrentPos
            StartChar = CurrentPos + 1
            CharLength = TxtLen
            If Not SecondLine Then
                SecondLine = True
                IndentText = IndentChar
            End If
            
            If UsePageBreaks Then
                If pEndOfPage Then
                    pFooter
                    pNewPage
                End If
            End If
        End If
    Next x

    CurrentX = LeftMargin
    
    If SameLine Then
        If PrintFlag Then
            Printer.Print IndentText & Mid(PrintVar, StartChar);
        Else
            ObjPrint.Print IndentText & Mid(PrintVar, StartChar);
        End If
    Else
        If PrintFlag Then
            Printer.Print IndentText & Mid(PrintVar, StartChar)
        Else
            ObjPrint.Print IndentText & Mid(PrintVar, StartChar)
        End If
    End If
    
    
End Sub

Public Function GetRemoveCRLF(ByVal TextString As String) As String
  Dim i As Integer, FoundString As Boolean
  
    Do
        FoundString = False
        i = InStr(TextString, vbCr)
        If i Then Mid(TextString, i, 1) = " ": FoundString = True
        i = InStr(TextString, vbLf)
        If i Then Mid(TextString, i, 1) = " ": FoundString = True
    Loop Until FoundString = False
    GetRemoveCRLF = TextString
End Function

Public Sub pPrintPicture(NewPic As StdPicture, _
                        Optional LeftMargin As Single = -1, _
                        Optional TopMargin As Single = -1, _
                        Optional pWidth As Single = 0, _
                        Optional pHeight As Single = 0, _
                        Optional ScaleToFit As Boolean = False, _
                        Optional MaintainRatio As Boolean = True)
 
  Dim xmin As Single
  Dim ymin As Single
  Dim wid As Single
  Dim Hgt As Single
  Dim aspect As Single
  Dim picBox As PictureBox
 
    If pWidth = 0 Then pWidth = pHeight
    If pHeight = 0 Then pHeight = pWidth
    If pWidth = 0 And pHeight = 0 Then ScaleToFit = True
    
    If PrintFlag Then
        Load frmMultiPgPreview
        Set picBox = frmMultiPgPreview.picPrintPic
        picBox.Picture = NewPic
        
        aspect = picBox.ScaleHeight / picBox.ScaleWidth
        If ScaleToFit Then
            wid = Printer.ScaleWidth
            Hgt = Printer.ScaleHeight
        Else
            wid = pWidth
            Hgt = pHeight
        End If
        
        If MaintainRatio Then
            If Hgt / wid > aspect Then
                Hgt = aspect * wid
                If LeftMargin = -1 Then
                    xmin = Printer.ScaleLeft
                Else
                    xmin = LeftMargin
                End If
                If TopMargin = -1 Then
                    ymin = (Printer.ScaleHeight - Hgt) / 2
                Else
                    ymin = TopMargin
                End If
            Else
                wid = Hgt / aspect
                If LeftMargin = -1 Then
                    xmin = (Printer.ScaleWidth - wid) / 2
                Else
                    xmin = LeftMargin
                End If
                If TopMargin = -1 Then
                    ymin = Printer.ScaleTop
                Else
                    ymin = TopMargin
                End If
            End If
        Else
            If LeftMargin = -1 Then
                xmin = Printer.ScaleLeft
            Else
                xmin = LeftMargin
            End If
            If TopMargin = -1 Then
                ymin = (Printer.ScaleHeight - Hgt) / 2
            Else
                ymin = TopMargin
            End If
        End If
        
        Printer.PaintPicture picBox.Picture, xmin, ymin, wid, Hgt
        Unload frmMultiPgPreview
        
    Else
        
        Set picBox = frmMultiPgPreview.picPrintPic
        picBox.Picture = NewPic 'LoadPicture(Filename)
        
        aspect = picBox.ScaleHeight / picBox.ScaleWidth
        If ScaleToFit Then
            wid = ObjPrint.ScaleWidth
            Hgt = ObjPrint.ScaleHeight
        Else
            wid = pWidth
            Hgt = pHeight
        End If
        
        If MaintainRatio Then
            If Hgt / wid > aspect Then
                Hgt = aspect * wid
                If LeftMargin = -1 Then
                    xmin = ObjPrint.ScaleLeft
                Else
                    xmin = LeftMargin
                End If
                If TopMargin = -1 Then
                    ymin = (ObjPrint.ScaleHeight - Hgt) / 2
                Else
                    ymin = TopMargin
                End If
            Else
                wid = Hgt / aspect
                If LeftMargin = -1 Then
                    xmin = (ObjPrint.ScaleWidth - wid) / 2
                Else
                    xmin = LeftMargin
                End If
                If TopMargin = -1 Then
                    ymin = ObjPrint.ScaleTop
                Else
                    ymin = TopMargin
                End If
            End If
        Else
            If LeftMargin = -1 Then
                xmin = ObjPrint.ScaleLeft
            Else
                xmin = LeftMargin
            End If
            If TopMargin = -1 Then
                ymin = (ObjPrint.ScaleHeight - Hgt) / 2
            Else
                ymin = TopMargin
            End If
        End If
        
        ObjPrint.PaintPicture picBox.Picture, xmin, ymin, wid, Hgt
        picBox.Picture = Nothing
        
    End If
    
    Set picBox = Nothing
    
End Sub

Public Property Let FontStrikethru(YesNo As Boolean)
    If PrintFlag Then
        Printer.FontStrikethru = YesNo
    Else
        ObjPrint.FontStrikethru = YesNo
    End If
End Property

Public Property Get FontStrikethru() As Boolean
    If PrintFlag Then
        FontStrikethru = Printer.FontStrikethru
    Else
        FontStrikethru = ObjPrint.FontStrikethru
    End If
End Property

Public Function GetFormalCase(ByVal TextString As String) As String
  Dim x As Integer
 
    '/* Cap the first letter if each word
    On Local Error Resume Next
    
    TextString = UCase$(Left$(TextString, 1)) & LCase$(Mid$(TextString, 2))
    
    '/* Look for space
    x = InStr(TextString, " ")
    If x Then
        Do
            Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
            x = x + 1
            x = InStr(x, TextString, " ")
            If x = 0 Or x + 1 > Len(TextString) Then Exit Do
        Loop
    End If
    
    '/* Look for .
    x = InStr(TextString, ".")
    If x Then
        Do
            Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
            x = x + 1
            x = InStr(x, TextString, ".")
            If x = 0 Or x + 1 > Len(TextString) Then Exit Do
        Loop
    End If
    
    '/* Look for Mc
    x = InStr(TextString, "Mc")
    If x Then
        Do
            Mid$(TextString, x + 2, 1) = UCase$(Mid$(TextString, x + 2, 1))
            x = x + 2
            x = InStr(x, TextString, "Mc")
            If x = 0 Or x + 2 > Len(TextString) Then Exit Do
        Loop
    End If
    
    '/* Look for O'
    x = InStr(TextString, "O'")
    If x Then
        Do
            Mid$(TextString, x + 2, 1) = UCase$(Mid$(TextString, x + 2, 1))
            x = x + 2
            x = InStr(x, TextString, "O'")
            If x = 0 Or x + 2 > Len(TextString) Then Exit Do
        Loop
    End If
    
    '/* Look for -
    x = InStr(TextString, "-")
    If x Then
       Do
           Mid$(TextString, x + 1, 1) = UCase$(Mid$(TextString, x + 1, 1))
           x = x + 1
           x = InStr(x, TextString, "-")
           If x = 0 Or x + 1 > Len(TextString) Then Exit Do
       Loop
    End If
    
    GetFormalCase = LTrim$(TextString)
 
End Function

Public Sub pRightTab(ByVal PrintVar As Variant, _
                            Optional ByVal LMargin As Single = 0.1, _
                            Optional SameLine As Boolean = False)
 
    CurrentX = PgWidth - (GetTextWidth(PrintVar) + LMargin)
    
    If SameLine Then
        If PrintFlag Then
            Printer.Print PrintVar;
        Else
            ObjPrint.Print PrintVar;
        End If
    Else
        If PrintFlag Then
            Printer.Print PrintVar
        Else
            ObjPrint.Print PrintVar
        End If
    End If

 
End Sub

Public Sub pCenter(ByVal PrintVar As String, _
                   Optional SameLine As Boolean = False)
    
    If GetTextWidth(PrintVar) > PgWidth Then
        pMultiline PrintVar, 0.1, PgWidth - 0.1, , SameLine
        Exit Sub
    End If
    CurrentX = ((PgWidth - GetTextWidth(PrintVar)) / 2)
    
    pPrint PrintVar, , SameLine
    
End Sub

Public Sub pRightJust(ByVal PrintVar As Variant, _
                      Optional ByVal xVAL As Single = -1, _
                      Optional SameLine As Boolean = False)
  
  Dim TxtWidth As Single
  
    TxtWidth = GetTextWidth(PrintVar & "  ")
    If xVAL = -1 Then xVAL = CurrentX + TxtWidth
    CurrentX = xVAL - GetTextWidth(PrintVar)
    
    If SameLine Then
        If PrintFlag Then
            Printer.Print PrintVar;
        Else
            ObjPrint.Print PrintVar;
        End If
    Else
        If PrintFlag Then
            Printer.Print PrintVar
        Else
            ObjPrint.Print PrintVar
        End If
    End If


End Sub

Public Sub pBox(Optional ByVal bLeft As Single = -1, _
                Optional ByVal bTop As Single = -1, _
                Optional ByVal bWidth As Single = -1, _
                Optional ByVal bHeight As Single = -1, _
                Optional ByVal ColorLine As Long = -1, _
                Optional ByVal ColorFill As Long = -1, _
                Optional FilledBox As FillStyleConstants = vbFSTransparent)
   
  Dim x As Single, y As Single
  
    y = CurrentY
    x = CurrentX
    
    If ColorLine = -1 Then ColorLine = ForeColor
    If ColorFill = -1 Then ColorFill = ColorLine
    If bLeft = -1 Then bLeft = CurrentX
    If bTop = -1 Then bTop = CurrentY
    If bWidth = -1 Then bWidth = PgWidth
    If bHeight = -1 Then bHeight = GetTextHeight("Sample Text")
    
    If FilledBox <> vbFSTransparent Then
        If PrintFlag Then
            Printer.FillColor = ColorFill
            Printer.FillStyle = FilledBox
            Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
            Printer.FillStyle = vbFSTransparent
        Else
            ObjPrint.FillColor = ColorFill
            ObjPrint.FillStyle = FilledBox
            ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
            ObjPrint.FillStyle = vbFSTransparent
        End If
    Else
        If PrintFlag Then
            Printer.FillStyle = vbFSTransparent
            Printer.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
        Else
            ObjPrint.FillStyle = vbFSTransparent
            ObjPrint.Line (bLeft, bTop)-(bLeft + bWidth, bTop + bHeight), ColorLine, B
        End If
    End If
    
    CurrentX = x
    CurrentY = y
    
End Sub

Public Sub pCircle(ByVal bLeft As Single, _
                   ByVal bTop As Single, _
                   ByVal bRadius As Single, _
                   Optional ByVal ColorLine As Long = -1, _
                   Optional ByVal ColorFill As Long = -1, _
                   Optional FilledCircle As FillStyleConstants = vbFSTransparent)
                   
    If ColorLine = -1 Then ColorLine = ForeColor
    If ColorFill = -1 Then ColorFill = ColorLine
    
    If PrintFlag Then
        If FilledCircle <> vbFSTransparent Then
            Printer.FillStyle = FilledCircle
            Printer.FillColor = ColorFill
        End If
        
        Printer.Circle (bLeft, bTop), bRadius, ColorLine
        
        Printer.FillStyle = vbFSTransparent
    Else
        If FilledCircle <> vbFSTransparent Then
            ObjPrint.FillStyle = FilledCircle
            ObjPrint.FillColor = ColorFill
        End If
        
        ObjPrint.Circle (bLeft, bTop), bRadius, ColorLine
        
        ObjPrint.FillStyle = vbFSTransparent
    End If
    
End Sub

Public Sub pEndDoc()
  Dim i As Integer
    
    If PrintFlag Then
        Printer.EndDoc
        Printer.ScaleMode = pSM
        SendToPrinter = False
    Else
        On Local Error Resume Next
        ObjPrint.ScaleMode = oSM
        
        SavePicture ObjPrint.Image, TempDir & "PPview" & CStr(PageNumber) & ".bmp"
    
        frmMultiPgPreview.PageNumber = PageNumber
        frmMultiPgPreview.Picture1.Picture = LoadPicture(TempDir & "PPview" & CStr(0) & ".bmp")
        frmMultiPgPreview.Show vbModal
    End If
End Sub

Public Sub pFontName(Optional ByVal pFontName As String = "Times New Roman")
    If PrintFlag Then
        Printer.FontName = pFontName
        Printer.Print "";
    Else
        ObjPrint.FontName = pFontName
        ObjPrint.Print "";
    End If
End Sub

Public Property Let FontSize(pSize As Integer)
    If PrintFlag Then
        Printer.FontSize = pSize
    Else
        '/* Sized by ratio since Scale method does not effect FontSize
        ObjPrint.FontSize = pSize * Ratio
    End If
End Property
Public Property Get FontSize() As Integer
    If PrintFlag Then
        FontSize = Printer.FontSize
    Else
        '/* Sized by ratio since Scale method does not effect FontSize
        FontSize = ObjPrint.FontSize / Ratio
    End If
End Property

Public Property Let ForeColor(NewColor As Long)
    If PrintFlag Then
        Printer.ForeColor = NewColor
    Else
        ObjPrint.ForeColor = NewColor
    End If
End Property
Public Property Get ForeColor() As Long
    If PrintFlag Then
        ForeColor = Printer.ForeColor
    Else
        ForeColor = ObjPrint.ForeColor
    End If
End Property

Public Sub pLine(Optional ByVal LeftPos As Single = 0, _
                 Optional ByVal RightPos As Single = 0, _
                 Optional ByVal LineWidth As Integer = 0, _
                 Optional IncludeQuarterSpace As Boolean = True)
  
  Dim eDrawWidth As Integer
    
    If IncludeQuarterSpace Then pQuarterSpace
    eDrawWidth = DrawWidth
    
    If LineWidth > 0 Then DrawWidth = LineWidth
    
    If PrintFlag Then
        If RightPos <= LeftPos Then RightPos = PgWidth
        Printer.Line (LeftPos, Printer.CurrentY)-(RightPos, Printer.CurrentY)
    Else
        If RightPos <= LeftPos Then RightPos = PgWidth
        ObjPrint.Line (LeftPos, ObjPrint.CurrentY)-(RightPos, ObjPrint.CurrentY)
    End If
    
    If LineWidth > 0 Then DrawWidth = eDrawWidth
    If IncludeQuarterSpace Then pQuarterSpace
End Sub

Public Sub pNewPage()

    On Local Error Resume Next
    If PrintFlag Then
        Printer.NewPage
    Else
        SavePicture ObjPrint.Image, TempDir & "PPview" & CStr(PageNumber) & ".bmp"
        ObjPrint.Cls
        PageNumber = PageNumber + 1
    End If
End Sub

Public Sub pPrint(Optional ByVal PrintVar As String = vbNullString, _
                  Optional ByVal LMargin As Single = -1, _
                  Optional SameLine As Boolean = False)

    If PrintVar = vbNullString Then
        '/* Empty String */
    Else
        If LMargin = -1 Then LMargin = CurrentX
        If GetTextWidth(PrintVar) + LMargin > PgWidth Then
            pMultiline PrintVar, LMargin, PgWidth - 0.1, , SameLine
            Exit Sub
        End If
    End If
    
    If LMargin >= 0 Then CurrentX = LMargin
    
    If SameLine Then
        If PrintFlag Then
            Printer.Print PrintVar;
        Else
            ObjPrint.Print PrintVar;
        End If
    Else
        If PrintFlag Then
            Printer.Print PrintVar
        Else
            ObjPrint.Print PrintVar
        End If
    End If

End Sub

Public Sub pPrintedDate(Optional PrintCentered As Boolean = False, _
                        Optional ByVal LMargin As Single = -1, _
                        Optional SameLine As Boolean = False)
  
  Dim PrintVar As String
  Dim FSize As Integer, FBold As Boolean, FItalic As Boolean
  Dim FUnderline As Boolean, FStrikethru As Boolean
  
    FSize = FontSize
    FBold = FontBold
    FItalic = FontItalic
    FUnderline = FontUnderline
    FStrikethru = FontStrikethru
    
    FontSize = 9
    FontBold = False
    FontUnderline = False
    FontItalic = False
    FontStrikethru = False
    PrintVar = "Printed: " & Format(Now, "ddd. mmm. d, yyyy \@ h:mm ampm")
    
    If PrintCentered Then
        pCenter PrintVar, SameLine
    Else
        pPrint PrintVar, LMargin, SameLine
    End If
    
    FontSize = FSize
    FontBold = FBold
    FontItalic = FItalic
    FontUnderline = FUnderline
    FontStrikethru = FStrikethru

End Sub

Public Sub pStartDoc()
  'Dim HeightRatio As Single
  'Dim WidthRatio As Single

    PageNumber = 0
    TempDir = Environ("TEMP") & "\"
    
    On Local Error Resume Next
    
    '/* Set the Printer's scale mode to inches
    pSM = Printer.ScaleMode
    Printer.ScaleMode = vbInches
    
    '/* Set the physical page size
    PgWidth = Printer.ScaleWidth
    PgHeight = Printer.ScaleHeight
    '/* Option 1:   Ratio = PgHeight / PgWidth
    BotMargin = PgHeight - 0.4
    
    If PrintFlag Then
        '/* Initialize printer
        Printer.Print "";
    Else
        '/* Initialize the preview object
        Load frmMultiPgPreview
        Set ObjPrint = frmMultiPgPreview!Picture1
        
        '/* Scale Object to Printer's printable area in Inches
        '/* Compare the height and with ratios to determine the
        '/* Ratio to use and how to size the picture box
        oSM = ObjPrint.ScaleMode
        ObjPrint.ScaleMode = vbInches
        
        '/* Option 1 - Larger but does not match printout because
        '/* scaled font sizes must be intergers
        'ObjPrint.Height = ObjPrint.Width * Ratio
        'HeightRatio = ObjPrint.ScaleHeight / PgHeight
        'WidthRatio = ObjPrint.ScaleWidth / PgWidth
        'If HeightRatio < WidthRatio Then
        '    Ratio = HeightRatio
        'Else
        '    Ratio = WidthRatio
        'End If
        'Ratio = Ratio - 0.1
            
        '/* Option 2 - Full size (1440 twips = 1 inch)
        ObjPrint.Width = (PgWidth + 0.25) * 1440
        ObjPrint.Height = (PgHeight + 0.25) * 1440
        'If Printer.Orientation = vbPRORPortrait Then
        '    ObjPrint.Height = (PgHeight + 1.8) * 1440
        'Else
        '    ObjPrint.Height = (PgHeight + 1.25) * 1440
        'End If
        Ratio = 1

        '/* Set default properties of the scroll bars
        frmMultiPgPreview!VScroll1.Max = Val(ObjPrint.Height * 0.5)
        frmMultiPgPreview!VScroll1.Min = -500
        frmMultiPgPreview!VScroll1.SmallChange = Val(frmMultiPgPreview!VScroll1.Max * 0.1)
        frmMultiPgPreview!VScroll1.LargeChange = Val(frmMultiPgPreview!VScroll1.Max * 0.5)
         
        frmMultiPgPreview!HScroll1.Max = Val(ObjPrint.Width * 0.25)
        frmMultiPgPreview!HScroll1.Min = -500
        frmMultiPgPreview!HScroll1.SmallChange = Val(frmMultiPgPreview!HScroll1.Max * 0.1)
        frmMultiPgPreview!HScroll1.LargeChange = Val(frmMultiPgPreview!HScroll1.Max * 0.5)
         
        '/* Set default properties of the object to match printer
        ObjPrint.Scale (0, 0)-(PgWidth, PgHeight)
        ObjPrint.FontName = Printer.FontName
        ObjPrint.FontSize = Printer.FontSize * Ratio
        ObjPrint.ForeColor = Printer.ForeColor
        ObjPrint.Picture = Nothing
        
    End If

End Sub

Public Property Let FontBold(YesNo As Boolean)
    If PrintFlag Then
        Printer.FontBold = YesNo
    Else
        ObjPrint.FontBold = YesNo
    End If
End Property

Public Property Get FontBold() As Boolean
    If PrintFlag Then
        FontBold = Printer.FontBold
    Else
        FontBold = ObjPrint.FontBold
    End If
End Property

Public Property Let FontItalic(YesNo As Boolean)
    If PrintFlag Then
        Printer.FontItalic = YesNo
    Else
        ObjPrint.FontItalic = YesNo
    End If
End Property

Public Property Get FontItalic() As Boolean
    If PrintFlag Then
        FontItalic = Printer.FontItalic
    Else
        FontItalic = ObjPrint.FontItalic
    End If
End Property

Public Function pEndOfPage(Optional ByVal Less As Single = 0) As Boolean
  Dim n As Single
    
    Less = Less + 0.25
    If PrintFlag Then
        n = Printer.CurrentY + Less
    Else
        n = ObjPrint.CurrentY + Less
    End If
    
    If n >= BotMargin Then
        pEndOfPage = True
    Else
        pEndOfPage = False
    End If
    
End Function

Public Property Let FontUnderline(YesNo As Boolean)
    If PrintFlag Then
        Printer.FontUnderline = YesNo
    Else
        ObjPrint.FontUnderline = YesNo
    End If
End Property

Public Property Get FontUnderline() As Boolean
    If PrintFlag Then
        FontUnderline = Printer.FontUnderline
    Else
        FontUnderline = ObjPrint.FontUnderline
    End If
End Property

Public Sub pHalfSpace()
  Dim eFont As Integer
  Dim hFont As Integer
 
    eFont = FontSize
    hFont = eFont \ 2
    If hFont < 1 Then hFont = 1
    FontSize = hFont
    pPrint
    FontSize = eFont
    
End Sub

Public Sub pDoubleLine(Optional ByVal LeftPos As Single = 0, _
                       Optional ByVal RightPos As Single = 0, _
                       Optional ByVal LineWidth As Integer = 0)
 
  Dim eFont As Integer, eDrawWidth As Integer
   
    eDrawWidth = DrawWidth
    If LineWidth > 0 Then DrawWidth = LineWidth
    
    If PrintFlag Then
        If RightPos <= LeftPos Then RightPos = PgWidth
        eFont = Printer.FontSize
        Printer.FontSize = 2
        Printer.Line (LeftPos, Printer.CurrentY)-(PgWidth, Printer.CurrentY)
        Printer.Print
        Printer.Line (LeftPos, Printer.CurrentY)-(PgWidth, Printer.CurrentY)
        Printer.FontSize = eFont \ 2
        Printer.Print
        Printer.FontSize = eFont
    Else
        If RightPos <= LeftPos Then RightPos = PgWidth
        eFont = ObjPrint.FontSize
        ObjPrint.FontSize = 2 * Ratio
        ObjPrint.Line (LeftPos, ObjPrint.CurrentY)-(PgWidth, ObjPrint.CurrentY)
        ObjPrint.Print
        ObjPrint.Line (LeftPos, ObjPrint.CurrentY)-(PgWidth, ObjPrint.CurrentY)
        ObjPrint.FontSize = eFont \ 2
        ObjPrint.Print
        ObjPrint.FontSize = eFont
    End If
    
    If LineWidth > 0 Then DrawWidth = eDrawWidth
End Sub

Public Sub pVerticalLine(Optional ByVal LeftPos As Single = -1, _
                         Optional ByVal TopPos As Single = -1, _
                         Optional ByVal BottomPos As Single = -1, _
                         Optional ByVal LineWidth As Integer = 0)
 
  Dim eDrawWidth As Integer, cY As Single, cX As Single, tH As Single
   
    eDrawWidth = DrawWidth
    cX = CurrentX
    cY = CurrentY
    If LineWidth > 0 Then DrawWidth = LineWidth
    If BottomPos = -1 Then BottomPos = BotMargin
    BottomPos = BottomPos + (Printer.TextHeight("Text") / 2)
    
    If LeftPos = -1 Then LeftPos = CurrentX
    If LeftPos > PgWidth Then LeftPos = PgWidth - 0.01
    If TopPos = -1 Then TopPos = CurrentY

    If PrintFlag Then
        Printer.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
    Else
        ObjPrint.Line (LeftPos, TopPos)-(LeftPos, BottomPos)
    End If
    
    CurrentX = cX
    CurrentY = cY
   
    If LineWidth > 0 Then DrawWidth = eDrawWidth
End Sub

Public Sub pSpaces(Optional ByVal xEnd As Single = -1, _
                   Optional ByVal xStart As Single = -1, _
                   Optional UseSymbol As Boolean = False)
  
  Dim xFontname As String
  Dim xForeColor As Long
  Dim tString As String
    
    xFontname = FontName
    xForeColor = ForeColor
    
    If xStart <> -1 Then CurrentX = xStart
    If xEnd = -1 Then xEnd = PgWidth
    
    If UseSymbol Then
        FontName = "Symbol"
        tString = "\"
        ForeColor = vbGreen
    Else
        tString = " "
    End If
    
    If CurrentX >= xEnd Then GoTo ExitSpaceSub
    Do
        pPrint tString, , True
    Loop Until CurrentX >= xEnd
    
ExitSpaceSub:
    FontName = xFontname
    ForeColor = xForeColor

End Sub

Public Sub p15Space()
  Dim eFont As Integer
  Dim hFont As Integer
    
    pPrint
    eFont = FontSize
    hFont = eFont \ 2
    FontSize = hFont
    pPrint
    FontSize = eFont
End Sub

Public Sub pFooter()
  Dim eFontS As Integer
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean
  Const tMargin = 0.5
      
    '/* Save current setting
    eFontN = FontName
    eFontS = FontSize
    eFontB = FontBold
    eFontI = FontItalic
    eFontU = FontUnderline
    eFontK = FontStrikethru
    If CurrentY < BotMargin Then CurrentY = BotMargin
    
    '/* Change settings
    pFontName
    FontSize = 10
    FontBold = False
    FontItalic = False
    FontUnderline = False
    FontStrikethru = False
    pLine , , 6
    
    pPrint "Printed: " & Format(Now, "ddd. mmmm d, yyyy \@ h:mm ampm"), tMargin, True
    pRightJust "Pg. " & GetPage, PgWidth - tMargin
    
    '/* Restore setting
    FontName = eFontN
    FontSize = eFontS
    FontBold = eFontB
    FontItalic = eFontI
    FontUnderline = eFontU
    FontStrikethru = eFontK
   
End Sub
Public Sub pBullet(Optional ByVal LMargin As Single = -1)
  Dim eFontN As String
  Dim eFontB As Boolean
  Dim eFontI As Boolean
  Dim eFontU As Boolean
  Dim eFontK As Boolean
      
    If LMargin = -1 Then LMargin = CurrentX
    
    '/* Save current setting
    eFontN = FontName
    eFontB = FontBold
    eFontI = FontItalic
    eFontU = FontUnderline
    eFontK = FontStrikethru
    
    '/* Change settings
    'FontName = "Wingdings"
    FontName = "Symbol"
    FontBold = False
    FontItalic = False
    FontUnderline = False
    FontStrikethru = False
    
    'pPrint Chr(164) & "  ", LMargin, True
    pPrint Chr(183) & "  ", LMargin, True
    
    '/* Restore setting
    FontName = eFontN
    FontBold = eFontB
    FontItalic = eFontI
    FontUnderline = eFontU
    FontStrikethru = eFontK
   
End Sub

Public Sub pQuarterSpace()
  Dim eFont As Integer
  Dim hFont As Integer
 
    eFont = FontSize
    hFont = eFont \ 4
    If hFont < 1 Then hFont = 1
    FontSize = hFont
    pPrint
    FontSize = eFont
End Sub

Public Sub pDots(ByVal xEnd As Single, Optional ByVal xStart As Single = 0)
    If xStart > 0 Then CurrentX = xStart
    If CurrentX >= xEnd Then Exit Sub
    Do
        pPrint ".", , True
    Loop Until CurrentX >= xEnd
End Sub

Public Function GetPage() As Variant
    If PrintFlag Then
       GetPage = Printer.Page
    Else
       GetPage = PageNumber + 1
    End If
End Function

Public Property Get SendToPrinter() As Boolean
    SendToPrinter = PrintFlag
End Property

Public Property Let SendToPrinter(ByVal vNewValue As Boolean)
    PrintFlag = vNewValue
End Property

Public Function GetPaperHeight() As Single
    GetPaperHeight = PgHeight
End Function
Public Function GetPaperWidth() As Single
    GetPaperWidth = PgWidth
End Function

Public Property Get CurrentX() As Single
    If PrintFlag Then
        CurrentX = Printer.CurrentX
    Else
        CurrentX = ObjPrint.CurrentX
    End If
End Property

Public Property Let CurrentX(ByVal NewXvalue As Single)
    If PrintFlag Then
        Printer.CurrentX = NewXvalue
    Else
        ObjPrint.CurrentX = NewXvalue
    End If
End Property

Public Property Get CurrentY() As Single
    If PrintFlag Then
        CurrentY = Printer.CurrentY
    Else
        CurrentY = ObjPrint.CurrentY
    End If
End Property

Public Property Let CurrentY(ByVal NewYvalue As Single)
    If PrintFlag Then
        Printer.CurrentY = NewYvalue
    Else
        ObjPrint.CurrentY = NewYvalue
    End If
End Property

Public Function GetStripQuotes(ByVal TextString As String) As String
    If Left(TextString, 1) = Chr(34) Then TextString = Mid(TextString, 2)
    If Right(TextString, 1) = Chr(34) Then TextString = Left(TextString, Len(TextString) - 1)
    GetStripQuotes = TextString
End Function

Public Property Get Orientation() As PageOrientation
    Orientation = Printer.Orientation
End Property

Public Property Let Orientation(ByVal vNewValue As PageOrientation)
    Printer.Orientation = vNewValue
End Property

Public Property Get DrawWidth() As Integer
    If PrintFlag Then
        DrawWidth = Printer.DrawWidth
    Else
        DrawWidth = ObjPrint.DrawWidth * 2
    End If
End Property

Public Property Let DrawWidth(ByVal NewWidth As Integer)
    If NewWidth < 1 Then NewWidth = 1
    If PrintFlag Then
        Printer.DrawWidth = NewWidth
    Else
        NewWidth = NewWidth \ 2
        If NewWidth < 1 Then NewWidth = 1
        ObjPrint.DrawWidth = NewWidth
    End If

End Property

Public Property Get BackColor() As Long
    If PrintFlag Then
       BackColor = GetBkColor(Printer.hdc)
    Else
       BackColor = ObjPrint.BackColor
    End If
End Property

Public Property Let BackColor(ByVal NewColor As Long)
  Dim r As Long
 
    If PrintFlag Then
        If NewColor = -1 Then
            Call MakeTrans '/* I am not sure why this is necessary
            r = SetBkColor(Printer.hdc, vbWhite)
            Printer.FontTransparent = True
        Else
            Printer.FontTransparent = False
            r = SetBkColor(Printer.hdc, NewColor)
        End If
    Else
        If NewColor = -1 Then
            ObjPrint.FontTransparent = True
            r = SetBkColor(ObjPrint.hdc, vbWhite)
        Else
            ObjPrint.FontTransparent = False
            r = SetBkColor(ObjPrint.hdc, NewColor)
        End If
    End If
    
End Property

Public Property Get FontName() As String
    If PrintFlag Then
        FontName = Printer.FontName
    Else
        FontName = ObjPrint.FontName
    End If
End Property

Public Property Let FontName(ByVal NewFont As String)
    If PrintFlag Then
        Printer.FontName = NewFont
        Printer.Print "";
    Else
        ObjPrint.FontName = NewFont
        ObjPrint.Print "";
    End If
End Property

Public Function GetTextWidth(TextString As Variant) As Single
    If PrintFlag Then
        GetTextWidth = Printer.TextWidth(TextString)
    Else
        GetTextWidth = ObjPrint.TextWidth(TextString)
    End If
End Function
Public Function GetTextHeight(TextString As Variant) As Single
    If PrintFlag Then
        GetTextHeight = Printer.TextHeight(TextString)
    Else
        GetTextHeight = ObjPrint.TextHeight(TextString)
    End If
End Function



Public Property Get FontTransparent() As Boolean
    If PrintFlag Then
        FontTransparent = Printer.FontTransparent
    Else
        FontTransparent = ObjPrint.FontTransparent
    End If
End Property

Public Property Let FontTransparent(ByVal vNewValue As Boolean)
    If PrintFlag Then
        If vNewValue Then
            '/* I am not sure why this is necessary but it doesn't work without it
            Call MakeTrans
            BackColor = -1
        Else
            Printer.FontTransparent = False
        End If
    Else
        ObjPrint.FontTransparent = vNewValue
    End If
End Property


Private Sub MakeTrans()
  Dim x As Single, y As Single
    x = CurrentX
    y = CurrentY
    Sleep 1
    BackColor = vbWhite
    Sleep 1
    pQuarterSpace
    Sleep 1
    CurrentX = x
    CurrentY = y
    Sleep 1

End Sub
