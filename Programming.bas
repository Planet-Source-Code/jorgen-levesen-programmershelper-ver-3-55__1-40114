Attribute VB_Name = "Programming"
' This Application is written by Jørgen E. Levesen
' MailTo: jorgen@levesen.com - Url: http://www.levesen.com
' Please also see frmAbout for a list of participants, without those this
' program was not to be. Thanks to your all !

Option Explicit
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0&
Public Const REG_SZ = 1 ' Unicode nul terminated String
Public Const REG_DWORD = 4 ' 32-bit number

Public Declare Function GetModuleHandle Lib _
    "Kernel" (ByVal lpModuleName As String) As Integer

Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal _
    bRevert As Long) As Long

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal _
    hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpmii As MENUITEMINFO) As Long

Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal _
    hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
Public Type MENUITEMINFO
    cbSize        As Long
    fMask         As Long
    fType         As Long
    fState        As Long
    wid           As Long
    hSubMenu      As Long
    hbmpChecked   As Long
    hbmpUnchecked As Long
    dwItemData    As Long
    dwTypeData    As String
    cch           As Long
End Type

'Menu item constants.
Public Const SC_CLOSE       As Long = &HF060&
Public Const xSC_CLOSE   As Long = -10
'SetMenuItemInfo fMask constants.
Public Const MENU_STATE     As Long = &H1&
Public Const MENU_ID        As Long = &H2&

'SetMenuItemInfo fState constants.
Public Const MFS_GRAYED     As Long = &H3&
Public Const MFS_CHECKED    As Long = &H8&
Public Const WM_NCACTIVATE  As Long = &H86

Public Declare Function ShellExceCute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
Declare Function WNetGetUserA Lib "mpr" (ByVal lpName As String, ByVal _
    lpUsername As String, lpnLength As Long) As Long

Public Declare Function SendMessage Lib _
    "user32" Alias "SendMessageA" (ByVal hWnd As _
   Long, ByVal wMsg As Long, ByVal wParam As _
   Long, lParam As Any) As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetDeviceCaps Lib "GDI32" (ByVal _
  hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
  ByVal hdc As Long) As Long

Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, _
   ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hSrcDC As Long, _
   ByVal XSrc As Long, ByVal YSrc As Long, _
   ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020

Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Const LB_GETITEMHEIGHT = &H1A1
Global Const PD_PRINTSETUP = &H40&

Global m_strProgramLng As String
Global m_dbLanguage As Database

Global m_strPrograming As String
Global m_dbPrograming As Database

Global m_strMyCodeSnippet As String
Global m_strCodeSnippet As String
Global m_dbMyCodeSnippet As Database
Global m_dbCodeSnippet As Database

Global m_strCodeZip As String
Global m_dbCodeZip As Database

Global m_FileExt As String
Global wdApp As Word.Application

Global m_lSnippet As Long
Global dateFromDate As String
Global dateToDate As String
Global m_boolSnippet As Boolean
Public Const WM_PASTE = &H302

Global Const SW_SHOWNORMAL = 1
Global i As Long
Global n As Integer
Global a As Long
Global m_iFormNo As Integer

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message
Global Const MB_DEFBUTTON2 = 256       ' Second button is default

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IdNo = 7                  ' No button pressed
Public Sub TileForm(frm As Form, pb As PictureBox)
Dim x As Long
Dim y As Long
Dim z As Long
    frm.AutoRedraw = True
    frm.ScaleMode = vbPixels
    pb.ScaleMode = vbPixels
    For x = 0 To frm.ScaleWidth Step pb.Width
        For y = 0 To frm.ScaleHeight Step pb.Height
            z = BitBlt(frm.hdc, x, y, pb.Height, pb.Width, pb.hdc, 0, 0, SRCCOPY)
        Next y
    Next x
    frm.Refresh
End Sub

Function ExtractFileName(FileName As String) As String
'Extract the File name from a full file name
    Dim pos As Integer
    Dim PrevPos As Integer

    pos = InStr(FileName, "\")
    If pos = 0 Then
        ExtractFileName = ""
        Exit Function
    End If
    
    Do While pos <> 0
        PrevPos = pos
        pos = InStr(pos + 1, FileName, "\")
    Loop

    ExtractFileName = Right(FileName, Len(FileName) - PrevPos)
End Function

Public Function IsDateBetween(dBeginDate As Date, dEndDate As Date, dToCheck As Date) As Boolean
    Dim intTotalDays As Integer
    Dim i As Integer
    Dim dNewDate As Date
    intTotalDays = DateDiff("d", dBeginDate, dEndDate)

    For i = 1 To intTotalDays
        dNewDate = DateAdd("d", i, dBeginDate)

        If dNewDate = dToCheck Then
            IsDateBetween = True
            Exit Function
        End If
    Next i
    IsDateBetween = False
End Function

Public Function IsOutlookPresent() As Boolean
Dim rVal As Variant, i As Integer
    rVal = getstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\mailto\shell\open\command", "")
    i = InStr(rVal, "OUTLOOK.EXE")
    If i <> 0 Then
        IsOutlookPresent = True
    Else
        IsOutlookPresent = False
    End If
End Function
Public Function MakeNewRecordset(RS As Recordset, strOldLanguage As String, strNewLanguage As String)
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

Public Function IsLanguagePresent(rsLanguage As Recordset, strLanguage As String) As Boolean
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

Public Function getstring(hKey As Long, strPath As String, strValue As String)
    'EXAMPLE:
    'text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    Dim r
    Dim lValueType
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                getstring = Left$(strBuf, intZeroPos - 1)
            Else
                getstring = strBuf
            End If
        End If
    End If
End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
Dim dwFlags As Long
Dim WebTest As Boolean
    ConnType = ""
    WebTest = InternetGetConnectedState(dwFlags, 0&)
    Select Case WebTest
        Case dwFlags And CONNECT_LAN: ConnType = "LAN"
        Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
        Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
        Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
        Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
        'Case dwflags And CONNECT_RAS:
        'ConnType = "Remote"
    End Select
    IsWebConnected = WebTest
End Function
Public Function FileExists(s As String) As Boolean
    If Dir(s) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function AppDir() As String
    '//(This sub makes sure an "\" is added to the app.path)
    If Right$(App.Path, 1) <> "\" Then
        AppDir = App.Path & "\"
    Else
        AppDir = App.Path
    End If
End Function
Public Function IsEMailAddress(ByVal sEmail As String, _
    Optional ByRef sReason As String) As Boolean
Dim sPreffix As String
Dim sSuffix As String
Dim sMiddle As String
Dim nCharacter As Integer
Dim sBuffer As String

    sEmail = Trim(sEmail)

    If Len(sEmail) < 8 Then
        IsEMailAddress = False
        sReason = "Too short"
        Exit Function
    End If

    If InStr(sEmail, "@") = 0 Then
        IsEMailAddress = False
        sReason = "Missing the @"
        Exit Function
    End If

    If InStr(InStr(sEmail, "@") + 1, sEmail, "@") <> 0 Then
        IsEMailAddress = False
        sReason = "Too many @"
        Exit Function
    End If

    If InStr(sEmail, ".") = 0 Then
        IsEMailAddress = False
        sReason = "Missing the period"
        Exit Function
    End If

    If InStr(sEmail, "@") = 1 Or InStr(sEmail, "@") = Len(sEmail) Or _
        InStr(sEmail, ".") = 1 Or InStr(sEmail, ".") = Len(sEmail) Then
        IsEMailAddress = False
        sReason = "Invalid format"
    Exit Function
End If

For nCharacter = 1 To Len(sEmail)
    sBuffer = Mid$(sEmail, nCharacter, 1)
    If Not (LCase(sBuffer) Like "[a-z]" Or sBuffer = "@" Or _
    sBuffer = "." Or sBuffer = "-" Or sBuffer = "_" Or _
    IsNumeric(sBuffer)) Then: IsEMailAddress = _
    False: sReason = "Invalid character": Exit Function
Next nCharacter

nCharacter = 0

On Error Resume Next

sBuffer = Right(sEmail, 4)
If InStr(sBuffer, ".") = 0 Then GoTo TooLong:
If Left(sBuffer, 1) = "." Then sBuffer = Right(sBuffer, 3)
If Left(Right(sBuffer, 3), 1) = "." Then sBuffer = Right(sBuffer, 2)
If Left(Right(sBuffer, 2), 1) = "." Then sBuffer = Right(sBuffer, 1)

If Len(sBuffer) < 2 Then
    IsEMailAddress = False
    sReason = "Suffix too short"
    Exit Function
End If

TooLong:
   If Len(sBuffer) > 3 Then
      IsEMailAddress = False
      sReason = "Suffix too long"
      Exit Function
   End If
   sReason = Empty
   IsEMailAddress = True
End Function
Public Function IsMicrosoftMailRunning() As Boolean
    'On Error GoTo IsMicrosoftMailRunning_Err
    'IsMicrosoftMailRunning = GetModuleHandle("MSMail")
    
'IsMicrosoftMailRunning_Err:
    'If Err Then
        'IsMicrosoftMailRunning = False
        'Err.Clear
    'Else
        'IsMicrosoftMailRunning = True
    'End If
End Function
Public Sub OpenDatabasePrint(dbString As String, rptReport As String)
Dim ac As New Access.Application, view As Long
    On Error GoTo errPrintDatabase
    ac.OpenCurrentDatabase (dbString)
    'You can use the follow constants:
    'acViewDesign
    'acViewNormal - prints the report immediately
    'acViewPreview
    view = acViewNormal
    ac.DoCmd.OpenReport rptReport, view
    ac.CloseCurrentDatabase
    Set ac = Nothing
    Beep
    MsgBox "Printing finished !"
    Exit Sub
    
errPrintDatabase:
    Beep
    MsgBox Err.Description, vbExclamation, "Print ...."
    Err.Clear
End Sub

Public Function SpellCheck(rf As RichTextBox, SpellID As Long)
Dim stText As String
Dim stNew_Text As String
Dim iPosition As Integer

    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText rf, vbCFRTF
    Set wdApp = New Word.Application
    With wdApp
        .Visible = True
        .Documents.Add
        .Selection.Paste
        .Selection.WholeStory
        .Selection.LanguageID = SpellID
        .Selection.NoProofing = False
        .Application.CheckLanguage = True
        .ActiveDocument.CheckSpelling
        .Selection.WholeStory
        'Clipboard.Clear
        'Clipboard.SetText .Selection, vbCFRTF
        'rf.TextRTF = Clipboard.GetText(vbCFRTF)
        stText = .Selection
        .Documents.Close 0
        .Quit
    End With
    Set wdApp = Nothing

    If Right$(stText, 1) = vbCr Then _
      stText = Left$(stText, Len(stText) - 1)
    stNew_Text = ""
    iPosition = InStr(stText, vbCr)
    Do While iPosition > 0
      stNew_Text = stNew_Text & Left$(stText, iPosition - 1) & vbCrLf
      stText = Right$(stText, Len(stText) - iPosition)
      iPosition = InStr(stText, vbCr)
    Loop
    
    stNew_Text = stNew_Text & stText
    rf = stNew_Text
    
End Function


Public Sub DisableButtons(n As Integer)
    On Error Resume Next
    With frmMDI.Toolbar1
        For i = 3 To 12
            Select Case n
            Case 1  'disable
                .Buttons(i).Enabled = False
            Case 2  'enable
                .Buttons(i).Enabled = True
            Case Else
            End Select
        Next
    End With
End Sub


Public Sub Dither(vForm As Form)
Dim intLoop As Integer
        vForm.AutoRedraw = True
        vForm.DrawStyle = vbInsideSolid
        vForm.DrawMode = vbCopyPen
        vForm.ScaleMode = vbPixels
        vForm.DrawWidth = 2
        vForm.ScaleHeight = 256
        For intLoop = 0 To 255
          vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
        Next intLoop
End Sub
Public Function GetFileSize(strFile As String) As String
Dim fso As New Scripting.FileSystemObject
Dim F As File
Dim lngBytes As Long
Const KB As Long = 1024
Const MB As Long = 1024 * KB
Const GB As Long = 1024 * MB
    Set F = fso.GetFile(fso.GetFile(strFile))
    lngBytes = F.Size
    If lngBytes < KB Then
        GetFileSize = Format(lngBytes) & " bytes"
    ElseIf lngBytes < MB Then
        GetFileSize = Format(lngBytes / KB, "0.00") & " KB"
    ElseIf lngBytes < GB Then
        GetFileSize = Format(lngBytes / MB, "0.00") & " MB"
    Else
        GetFileSize = Format(lngBytes / GB, "0.00") & " GB"
    End If
End Function
Public Function MakeNewLanguage(strLanguage As String)
Dim RS As Recordset, iNo As Integer, boolNotFound As Boolean, iCount As Integer
Dim db As DAO.Database
Dim tbl As DAO.TableDef
    Set db = DBEngine.OpenDatabase(App.Path & "\ScheduleLang.mdb")
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
                Call MakeNewRecordset(RS, "ENG", strLanguage)
            End If
        End Select
    Next
    Exit Function
    
errNewLanguage:
    Beep
    MsgBox Err.Description, vbInformation, "New Language"
    Resume Next
End Function

Public Sub SendOutlookMail(Subject As String, Recipient As String, Message As String)
    On Error GoTo ErrorHandler
    Dim oLapp As Object
    Dim oItem As Object
    
    Set oLapp = CreateObject("Outlook.application")
    Set oItem = oLapp.CreateItem(0)
   
    With oItem
       .Subject = Subject
       .To = Recipient
       .Body = Message
       .Send
    End With
    
ErrorHandler:
    Set oLapp = Nothing
    Set oItem = Nothing
    Exit Sub
End Sub

Sub ResizeControls(frmName As Form, winstate As Integer)
    On Error Resume Next
    Dim designwidth As Integer, designheight As Integer, designfontsize As Integer
    Dim currentfontsize As Integer
    Dim numofcontrols As Integer, a As Integer
    Dim movetype As String, moveamount As Integer
    Dim GetResolutionX As Integer, GetResolutionY As Integer
    Dim ratiox As Integer, ratioy As Integer, fontratio As Integer
    
    'Change the designwidth and the designheight according to the resolution that the form was
    'designed at
    designwidth = 1024
    designheight = 768
    designfontsize = 96
    
    GetResolutionX = Screen.Width / Screen.TwipsPerPixelX
    GetResolutionY = Screen.Height / Screen.TwipsPerPixelY
    
    'Work out the ratio for resizing the controls
    ratiox = GetResolutionX / designwidth
    ratioy = GetResolutionY / designheight
    'check to see what size of fonts are being used

    If IsScreenFontSmall Then
        currentfontsize = 96
    Else
        currentfontsize = 120
    End If
    
    'work out the ratio for the fontsize
    fontratio = designfontsize / currentfontsize
    If ratiox = 1 And ratioy = 1 And fontratio = 1 Then Exit Sub
    numofcontrols = frmName.Controls.Count - 1 'count the number of controls on the form

    If winstate = 0 Then 'if the form isn't fullscreen then
        frmName.Height = frmName.Height * ratioy
        frmName.Width = frmName.Width * ratiox

        If frmName.Tag <> "" Then
            movetype = Left(frmName.Tag, 1)
            moveamount = Mid(frmName.Tag, 2, Len(frmName.Tag))

            Select Case movetype
                Case "L"
                frmName.Left = frmName.Left + moveamount
                Case "T"
                frmName.Top = frmName.Top + moveamount
                Case "H"
                frmName.Height = frmName.Height + moveamount
                Case "W"
                frmName.Width = frmName.Width + moveamount
            End Select
    End If
ElseIf winstate = 2 Then 'otherwise if it is fullscreen then
    frmName.Width = Screen.Width
    frmName.Height = Screen.Height
    frmName.Top = 0
    frmName.Left = 0
End If

For a = 0 To numofcontrols 'loop through Each control

    If frmName.Controls(a).Font.Size <= 8 And ratiox < 1 Then
        frmName.Controls(a).Font.Name = "Small Fonts"
        frmName.Controls(a).Font.Size = frmName.Controls(a).Font.Size - 0.5
    Else
        frmName.Controls(a).Font.Size = frmName.Controls(a).Font.Size * ratiox
    End If

    If TypeOf frmName.Controls(a) Is Line Then
        frmName.Controls(a).X1 = frmName.Controls(a).X1 * ratiox
        frmName.Controls(a).Y1 = frmName.Controls(a).Y1 * ratioy
        frmName.Controls(a).X2 = frmName.Controls(a).X2 * ratiox
        frmName.Controls(a).Y2 = frmName.Controls(a).Y2 * ratioy
    ElseIf TypeOf frmName.Controls(a) Is PictureBox Then
        frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
        frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
        frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
        frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
        frmName.Controls(a).ScaleHeight = frmName.Controls(a).ScaleHeight * ratioy
        frmName.Controls(a).ScaleWidth = frmName.Controls(a).ScaleWidth * ratiox
    'ElseIf TypeOf frmName.Controls(a) Is Toolbar Then
        'frmName.Controls(a).ButtonHeight = frmName.Controls(a).ButtonHeight * ratioy
        'frmName.Controls(a).ButtonWidth = frmName.Controls(a).ButtonWidth * ratiox
        'frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
        'frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
        'frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
        'frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
    'ElseIf TypeOf frmName.Controls(a) Is MSFlexGrid Then
        'frmName.Controls(a).ColWidth = frmName.Controls(a).ColWidth * ratiox
        'frmName.Controls(a).RowHeight = frmName.Controls(a).RowHeight * ratioy
        'frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
        'frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
        'frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
        'frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
    Else
        frmName.Controls(a).Width = frmName.Controls(a).Width * ratiox
        frmName.Controls(a).Height = frmName.Controls(a).Height * ratioy
        frmName.Controls(a).Top = frmName.Controls(a).Top * ratioy
        frmName.Controls(a).Left = frmName.Controls(a).Left * ratiox
    End If

    If frmName.Controls(a).Tag <> "" Then
        movetype = Left(frmName.Controls(a).Tag, 1)
        moveamount = Mid(frmName.Controls(a).Tag, 2, Len(frmName.Controls(a).Tag))

        Select Case movetype
            Case "L"
            frmName.Controls(a).Left = frmName.Controls(a).Left + moveamount
            Case "T"
            frmName.Controls(a).Top = frmName.Controls(a).Top + moveamount
            Case "H"
            frmName.Controls(a).Height = frmName.Controls(a).Height + moveamount
            Case "W"
            frmName.Controls(a).Width = frmName.Controls(a).Width + moveamount
        End Select
End If
Next a

If fontratio <> 1 Then

If winstate = 0 Then
    frmName.Height = frmName.Height * fontratio
    frmName.Width = frmName.Width * fontratio

    If frmName.Tag <> "" Then
        movetype = Left(frmName.Tag, 1)
        moveamount = Mid(frmName.Tag, 2, Len(frmName.Tag))


        Select Case movetype
            Case "L"
            frmName.Left = frmName.Left + moveamount
            Case "T"
            frmName.Top = frmName.Top + moveamount
            Case "H"
            frmName.Height = frmName.Height + moveamount
            Case "W"
            frmName.Width = frmName.Width + moveamount
        End Select
End If
ElseIf winstate = 2 Then
frmName.Width = Screen.Width
frmName.Height = Screen.Height
frmName.Top = 0
frmName.Left = 0
End If

For a = 0 To numofcontrols

If frmName.Controls(a).Font.Size <= 8 And fontratio < 1 Then
    frmName.Controls(a).Font.Name = "Small Fonts"
    frmName.Controls(a).Font.Size = frmName.Controls(a).Font.Size - 0.5
Else
    frmName.Controls(a).Font.Size = frmName.Controls(a).Font.Size * fontratio
End If

If TypeOf frmName.Controls(a) Is Line Then
    frmName.Controls(a).X1 = frmName.Controls(a).X1 * fontratio
    frmName.Controls(a).Y1 = frmName.Controls(a).Y1 * fontratio
    frmName.Controls(a).X2 = frmName.Controls(a).X2 * fontratio
    frmName.Controls(a).Y2 = frmName.Controls(a).Y2 * fontratio
ElseIf TypeOf frmName.Controls(a) Is PictureBox Then
    frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
    frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
    frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
    frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
    frmName.Controls(a).ScaleHeight = frmName.Controls(a).ScaleHeight * fontratio
    frmName.Controls(a).ScaleWidth = frmName.Controls(a).ScaleWidth * fontratio
'ElseIf TypeOf frmName.Controls(a) Is Toolbar Then
    'frmName.Controls(a).ButtonHeight = frmName.Controls(a).ButtonHeight * fontratio
    'frmName.Controls(a).ButtonWidth = frmName.Controls(a).ButtonWidth * fontratio
    'frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
    'frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
    'frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
    'frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
'ElseIf TypeOf frmName.Controls(a) Is MSFlexGrid Then
    'frmName.Controls(a).ColWidth = frmName.Controls(a).ColWidth * fontratio
    'frmName.Controls(a).RowHeight = frmName.Controls(a).RowHeight * fontratio
    'frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
    'frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
    'frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
   'frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
Else
    frmName.Controls(a).Width = frmName.Controls(a).Width * fontratio
    frmName.Controls(a).Height = frmName.Controls(a).Height * fontratio
    frmName.Controls(a).Top = frmName.Controls(a).Top * fontratio
    frmName.Controls(a).Left = frmName.Controls(a).Left * fontratio
End If
Next a
End If
End Sub
Public Function IsScreenFontSmall() As Boolean
    Dim hWndDesk As Long
    Dim hDCDesk As Long
    Dim logPix As Long
    Dim r As Long
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    r = ReleaseDC(hWndDesk, hDCDesk)
    If logPix = 96 Then IsScreenFontSmall = True
    Exit Function
End Function
