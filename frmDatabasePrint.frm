VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDatabasePrint 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Access Database Structure Print"
   ClientHeight    =   3360
   ClientLeft      =   2565
   ClientTop       =   2565
   ClientWidth     =   5190
   Icon            =   "frmDatabasePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Height          =   615
      Left            =   3840
      Picture         =   "frmDatabasePrint.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Height          =   615
      Left            =   120
      Picture         =   "frmDatabasePrint.frx":2004
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2640
      Width           =   4935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   90
      TabIndex        =   3
      Top             =   1560
      Width           =   5010
      Begin VB.CheckBox chkSeparated 
         BackColor       =   &H00000000&
         Caption         =   "Separate Page Per Table"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkSystemTables 
         BackColor       =   &H00000000&
         Caption         =   "Print System Tables"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton optHTML 
         BackColor       =   &H00000000&
         Caption         =   "HTML"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton optPrinter 
         BackColor       =   &H00000000&
         Caption         =   "Printer"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   4560
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Access Database"
      Filter          =   "Access Databases *.mdb |*.mdb"
      InitDir         =   "C:\"
   End
   Begin VB.TextBox txtDBPath 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   128
      TabIndex        =   0
      Top             =   960
      Width           =   3630
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4950
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 - Click Print"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   885
      TabIndex        =   2
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 - Set Your Print Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   885
      TabIndex        =   1
      Top             =   240
      Width           =   2190
   End
End
Attribute VB_Name = "frmDatabasePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code Author : Joseph B. Surls
'Author E-Mail: joseph.surls@verizon.net
'Submission Date : March 17, 2002       Happy St. Pat's!!

'I wrote this code for work. When you don't have Access
'on a user's machine where your VB program is, you can
'still check out its structure to pinpoint a problem.

'I hope this code helps somebody. Feel free to use this code
'and/or change it in any way. Drop me an E and let me know
'if I can help in any way.
'----------------------------------------------------------
'this code modified 2002.03.26 By J.Levesen

Option Explicit

Dim dbAccess As DAO.Database    'Database Object
Dim rsAccess As DAO.Recordset   'Recordset Object
Dim rsLanguage As Recordset
Dim i As Integer
Dim J As Long
Dim oTable As DAO.TableDef  'TableDef Object
Dim oField As DAO.Field     'Field Object
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
                If IsNull(.Fields("label1(0)")) Then
                    .Fields("Label1(0)") = Label1(0).Caption
                Else
                    Label1(0).Caption = .Fields("Label1(0)")
                End If
                If IsNull(.Fields("label1(1)")) Then
                    .Fields("Label1(1)") = Label1(1).Caption
                Else
                    Label1(1).Caption = .Fields("Label1(1)")
                End If
                If IsNull(.Fields("chkSeparated")) Then
                    .Fields("chkSeparated") = chkSeparated.Caption
                Else
                    chkSeparated.Caption = .Fields("chkSeparated")
                End If
                If IsNull(.Fields("chkSystemTables")) Then
                    .Fields("chkSystemTables") = chkSystemTables.Caption
                Else
                    chkSystemTables.Caption = .Fields("chkSystemTables")
                End If
                If IsNull(.Fields("optPrinter")) Then
                    .Fields("optPrinter") = optPrinter.Caption
                Else
                    optPrinter.Caption = .Fields("optPrinter")
                End If
                If IsNull(.Fields("optHTML")) Then
                    .Fields("optHTML") = optHTML.Caption
                Else
                    optHTML.Caption = .Fields("optHTML")
                End If
                If IsNull(.Fields("cmdBrowse")) Then
                    .Fields("cmdBrowse") = cmdBrowse.ToolTipText
                Else
                    cmdBrowse.ToolTipText = .Fields("cmdBrowse")
                End If
                If IsNull(.Fields("cmdPrint")) Then
                    .Fields("cmdPrint") = cmdPrint.ToolTipText
                Else
                    cmdPrint.ToolTipText = .Fields("cmdPrint")
                End If
                .Update
                Exit Sub
            End If
        .MoveNext
        Loop
        
        .AddNew
        .Fields("Language") = m_FileExt
        .Fields("Form") = Me.Caption
        .Fields("label1(0)") = Label1(0).Caption
        .Fields("label1(1)") = Label1(1).Caption
        .Fields("chkSeparated") = chkSeparated.Caption
        .Fields("chkSystemTables") = chkSystemTables.Caption
        .Fields("optPrinter") = optPrinter.Caption
        .Fields("optHTML") = optHTML.Caption
        .Fields("cmdBrowse") = cmdBrowse.ToolTipText
        .Fields("cmdPrint") = cmdPrint.ToolTipText
        .Update
    End With
End Sub

Private Sub cmdBrowse_Click()
On Error GoTo CancelBrowse
    'Open Common Dialog for User to input Database Path
    With dlgCommon
        .CancelError = True
        .InitDir = App.Path
        .DialogTitle = "Open Database..."
        .Filter = "Access Databases *.mdb|*.mdb"
        .FileName = ""
        .ShowOpen
        txtDBPath = .FileName
    End With
    Exit Sub
CancelBrowse:
    If Err.Number = 32755 Then 'User Pressed Cancel button
        Exit Sub
    Else
        MsgBox Err.Number & Chr(10) & _
            Err.Description
    End If
End Sub

Private Sub cmdPrint_Click()
On Error GoTo NoDB
    'If no printers on user's system, get out
    If Printers.Count < 1 Then
        Exit Sub
        Beep
        MsgBox "There are no printers connected !"
    End If
    
    'If no DB specified, get out
    If txtDBPath = "" Then Exit Sub
    
    Set dbAccess = OpenDatabase(Trim(txtDBPath), True, True)
    
    If optHTML.Value = True Then 'Print Structure in an HTML file
        PrintHTML
        Set dbAccess = Nothing
        Exit Sub
    Else                         'Print Structure to a printer
        Screen.MousePointer = vbHourglass
        Printer.Print Trim(txtDBPath)
        Printer.Print
        Printer.Print
        For Each oTable In dbAccess.TableDefs 'Loop through each table in the database
        
            'this next line determines whether to print the Access System tables or not
            If chkSystemTables.Value = vbChecked Or Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
                
                'Printer Setup Header Stuff
                Printer.FontSize = 14
                Printer.FontBold = True
                Printer.Print "TABLE NAME = " & oTable.Name
                Printer.FontSize = 8
                Printer.FontBold = False
                Printer.Print "======================================="
                Printer.Print "Date Created =" & oTable.DateCreated
                Printer.Print "Date Last Modified = " & oTable.LastUpdated
                Printer.Print "Records = " & oTable.RecordCount
                Printer.Print "---------------------------------------------------"
                Printer.Print ""
                Printer.Print ""
                
                'Dont print System table breakdown
                If Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
                    'open recordset on current table
                    Set rsAccess = dbAccess.OpenRecordset(oTable.Name, dbOpenTable)
                    'All this X and Y stuff sets up the Columns and headings
                    Printer.CurrentX = 500
                    Printer.FontBold = True
                    Printer.Print "Fields Listing"
                    Printer.FontBold = False
                    Printer.CurrentX = 1000
                    J = Printer.CurrentY
                    Printer.Print "Field Name"
                    Printer.CurrentX = 3000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "Type"
                    Printer.CurrentX = 5000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "Size"
                    Printer.CurrentX = 7000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "Required"
                    Printer.CurrentX = 9000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "Allow Null"
                    Printer.CurrentX = 1000
                    J = Printer.CurrentY
                    Printer.Print "-------------------"
                    Printer.CurrentX = 3000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "--------"
                    Printer.CurrentX = 5000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "--------"
                    Printer.CurrentX = 7000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "--------------"
                    Printer.CurrentX = 9000
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.Print "---------------"
                    i = 0
                    
                    'Loop thru each field in current table
                    'Line up columns and print field info
                    For Each oField In rsAccess.Fields
                        Printer.CurrentX = 1000
                        J = Printer.CurrentY
                        Printer.Print oField.Name
                        Printer.CurrentX = 3000
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        
                        'convert datatype into English
                        Printer.Print GetFieldType(oField.Type)
                        
                        Printer.CurrentX = 5000
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        Printer.Print oField.Size
                        Printer.CurrentX = 7000
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        Printer.Print oField.Required
                        Printer.CurrentX = 9000
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        Printer.Print oField.AllowZeroLength
                        i = i + 1
                    Next
                End If
                
                'Get any indexes for current table
                If oTable.Indexes.Count > 0 Then
                    Printer.Print ""
                    Printer.CurrentX = 500
                    Printer.FontBold = True
                    Printer.Print "Index Listing"
                    Printer.FontBold = False
                    J = Printer.CurrentY
                    Printer.CurrentX = 1000
                    Printer.Print "Index Name"
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.CurrentX = 3000
                    Printer.Print "Fields"
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.CurrentX = 6000
                    Printer.Print "Unique"
                    J = Printer.CurrentY
                    Printer.CurrentX = 1000
                    Printer.Print "----------------"
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.CurrentX = 3000
                    Printer.Print "----------"
                    If Printer.CurrentY < J Then
                        J = Printer.CurrentY
                    End If
                    Printer.CurrentY = J
                    Printer.CurrentX = 6000
                    Printer.Print "----------"
                    
                    'loop thru table Indexes (if any)
                    For i = 0 To oTable.Indexes.Count - 1
                        J = Printer.CurrentY
                        Printer.CurrentX = 1000
                        Printer.Print oTable.Indexes(i).Name
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        Printer.CurrentX = 3000
                        Printer.Print oTable.Indexes(i).Fields
                        If Printer.CurrentY < J Then
                            J = Printer.CurrentY
                        End If
                        Printer.CurrentY = J
                        Printer.CurrentX = 6000
                        Printer.Print oTable.Indexes(i).Unique
                    Next i
                End If
                
                'Clear recordset for next table
                Set rsAccess = Nothing
                
                'Print each table on separate page or not
                If chkSeparated.Value = vbChecked Then
                    Printer.EndDoc
                Else
                    Printer.Print ""
                    Printer.Print ""
                End If
            End If
        Next
        If Not chkSeparated.Value = vbChecked Then
            Printer.EndDoc
        End If
        
        'Clear database variable
        Set dbAccess = Nothing
        MsgBox "Your Access Structure has been printed to " & Printer.DeviceName, vbInformation, "Complete"
        Screen.MousePointer = vbDefault
        Unload Me
        Exit Sub
    End If
NoDB:
    If Err.Number = 3031 Then 'Database needs a password
        frmDatabasePassword.Show vbModal
        If frmDatabasePassword.pblnCancel = True Then Exit Sub
        cmdPrint_Click
        Err.Clear
        Exit Sub
    End If
    MsgBox Err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Function GetFieldType(TypeCode As Integer)
'This routine accepts the Fieldtype variable and returns the
'English version for printing
    Select Case TypeCode
        Case dbBinary
            GetFieldType = "Binary"
        Case dbBoolean
            GetFieldType = "Boolean"
        Case dbByte
            GetFieldType = "Byte"
        Case dbChar
            GetFieldType = "Character"
        Case dbCurrency
            GetFieldType = "Currency"
        Case dbDate
            GetFieldType = "Date/Time"
        Case dbDecimal
            GetFieldType = "Decimal"
        Case dbDouble
            GetFieldType = "Double"
        Case dbFloat
            GetFieldType = "Float"
        Case dbGUID
            GetFieldType = "GUID"
        Case dbInteger
            GetFieldType = "Integer"
        Case dbLong
            GetFieldType = "Long"
        Case dbLongBinary
            GetFieldType = "OLE Object"
        Case dbMemo
            GetFieldType = "Memo"
        Case dbNumeric
            GetFieldType = "Numeric"
        Case dbSingle
            GetFieldType = "Single"
        Case dbText
            GetFieldType = "Text"
        Case dbTime
            GetFieldType = "Time"
        Case dbTimeStamp
            GetFieldType = "TimeStamp"
        Case dbVarBinary
            GetFieldType = "VarBinary"
        Case Else
            GetFieldType = "Undetermined"
    End Select
End Function

Private Sub PrintHTML()
'this routine prints the Access Structure to an HTML file
Dim SaveFile As String
On Error GoTo CancelHTML
    
    'More Common Dialog
    With dlgCommon
        .CancelError = True
        .DialogTitle = "Save HTML Page As..."
        .Filter = "Web Page *.htm|*.htm;*.html"
        .InitDir = "C:\"
        .FileName = "Structure.htm"
        .ShowSave
        SaveFile = .FileName
    End With
    DoEvents
    Open SaveFile For Output As #2
    
    'Set database Object
    Set dbAccess = OpenDatabase(Trim(txtDBPath), True, True)
    
    'HTML Template stuff
    Print #2, "<html>"
    Print #2, "<head>"
    Print #2, "<meta name='Access Structure Print' content=Jørgen Levesen'>"
    Print #2, "<title>" & "Access Structure for " & Trim(txtDBPath) & "</title>"
    Print #2, "</head>"
    Print #2, "<body bgcolor='#0099FF'>"
    Print #2, "<p><font size='1'>"
    Print #2, Trim(txtDBPath)
    Print #2, "</a></font></p>"
    
    'Loop thru each table in Database
    For Each oTable In dbAccess.TableDefs
        Print #2, "<p><b><u><font size='4' color='#000000'>"
        Print #2, "Table " & oTable.Name & "</font><br>"
        Print #2, "</u></b><font size='2'>"
        Print #2, "Date Created - " & oTable.DateCreated & "<br>"
        Print #2, "Date Last Modified - " & oTable.LastUpdated & "<br>"
        Print #2, "Records - " & oTable.RecordCount & "<br>"
        Print #2, "-----------------------------------------------------------"
        Print #2, "</font></p>"
        
        'No System Tables
        If Not UCase(Left(oTable.Name, 4)) = "MSYS" Then
            
            'open recordset for each table
            Set rsAccess = dbAccess.OpenRecordset(oTable.Name, dbOpenTable)
            Print #2, "<p>&nbsp;&nbsp; <font size='2'> </font><b><font size='3'>Fields Listing</font></b></p>"
            Print #2, "<table border='0' width='100%'>"
            Print #2, "<tr><td width='10%' align='center'></td>"
            Print #2, "<td width='30%' align='center'>"
            Print #2, "<p align='center'><u>Field Name</u></td>"
            Print #2, "<td width='20%' align='center'><u>Type</u></td>"
            Print #2, "<td width='10%' align='center'><u>Size</u></td>"
            Print #2, "<td width='10%' align='center'><u>Required</u></td>"
            Print #2, "<td width='44%' align='center'><u>Allow Null</u></td>"
            Print #2, "<td width='16%' align='center'></td></tr>"
            
            'Loop thru each field in current table
            For Each oField In rsAccess.Fields
                Print #2, "<tr><td width='10%' align='center'></td>"
                Print #2, "<td width='30%' align='center'>"
                Print #2, oField.Name & "</td>"
                Print #2, "<td width='20%' align='center'>"
                
                'convert data type to English
                Print #2, GetFieldType(oField.Type) & "</td>"
                Print #2, "<td width='10%' align='center'>"
                Print #2, oField.Size & "</td>"
                Print #2, "<td width='10%' align='center'>"
                Print #2, oField.Required & "</td>"
                Print #2, "<td width='44%' align='center'>"
                Print #2, oField.AllowZeroLength & "</td>"
                Print #2, "<td width='16%' align='center'></td>"
                Print #2, "</tr>"
            Next
            Print #2, "</table>"
            
            'Table Indexes
            If oTable.Indexes.Count > 0 Then
                Print #2, "<p>&nbsp;&nbsp;&nbsp; <b>Index Listing</b></p>"
                Print #2, "<table border='0' width='100%'>"
                Print #2, "<tr>"
                Print #2, "<td width='7%' align='center'></td>"
                Print #2, "<td width='23%' align='center'><u>Index Name</u></td>"
                Print #2, "<td width='44%' align='center'><u>Fields</u></td>"
                Print #2, "<td width='19%' align='center'><u>Unique</u></td>"
                Print #2, "<td width='7%' align='center'></td>"
                Print #2, "</tr>"
                For i = 0 To oTable.Indexes.Count - 1
                    Print #2, "<tr>"
                    Print #2, "<td width='7%' align='center'></td>"
                    Print #2, "<td width='23%' align='center'>"
                    Print #2, oTable.Indexes(i).Name & "</td>"
                    Print #2, "<td width='44%' align='center'>"
                    Print #2, oTable.Indexes(i).Fields & "</td>"
                    Print #2, "<td width='19%' align='center'>"
                    Print #2, oTable.Indexes(i).Unique & "</td>"
                    Print #2, "<td width='7%' align='center'></td>"
                    Print #2, "</tr>"
                Next i
            End If
            Print #2, "</table>"
            Print #2, "<p>====================================================================================</p>"
        End If
    Next
    Print #2, "<p align='center'>End of Listing<br>"
    Print #2, "This Page Created by Access Structure Print Software - " & _
        Date & "</p>"
    Print #2, "</body>"
    Print #2, "</html>"
    Close #2
    MsgBox "Your HTML Listing has been saved as " & dlgCommon.FileName, vbInformation, "Complete"
    Unload Me
    Exit Sub
    
CancelHTML:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox Err.Number & Chr(10) & Err.Description
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    ReadText
End Sub

Private Sub Form_Load()
    On Error Resume Next    'it is just language text
    Set rsLanguage = m_dbLanguage.OpenRecordset("frmDatabasePrint")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set frmDatabasePrint = Nothing
    rsAccess.Close
    rsLanguage.Close
End Sub


Private Sub optHTML_Click()
    'Disable Irrelevant Check Buttons
    chkSeparated.Enabled = False
    chkSystemTables.Enabled = False
End Sub

Private Sub optPrinter_Click()
    'Enable Relevant Check Buttons
    chkSeparated.Enabled = True
    chkSystemTables.Enabled = True
End Sub
