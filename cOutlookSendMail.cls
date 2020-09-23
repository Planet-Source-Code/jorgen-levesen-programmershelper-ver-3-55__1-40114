VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutlookSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------Class - cOutlookSendMail.cls----------------------------------
'
' Class Module cOutlookSendMail
' File: cOutlookSendMail.cls
' Author: John McGlothlin - kelly-and-john@worldnet.att.net
' Date: 10/4/01
' Purpose: Automate sending of mail thru Outlook 2000
'
'------------------------------------------------------------------------
' Usage:
'    Dim clsOutlook As cOutlookSendMail
'
'    Set clsOutlook = New cOutlookSendMail
'
'    clsOutlook.StartOutlook     ' MUST be first call
'
'    clsOutlook.CreateNewMail    ' MUST be before addding any info
'
'    clsOutlook.Recipient_TO = "bob@worldnet.att.net"
'    clsOutlook.Recipient_TO = "joe@aol.com"
'    clsOutlook.Recipient_CC = "tom@att.com"
'    clsOutlook.Recipient_CC = "kelly@hotmail.com"
'    clsOutlook.Recipient_BCC = "andy@hotmail.com"
'    clsOutlook.Recipient_BCC = "mario@whereever.com"
'
'    clsOutlook.Subject = "Test"
'    clsOutlook.Body = "Hello world - how be you ???"
'    clsOutlook.Attachment "c:\config.sys", "display name"
'    clsOutlook.Attachment "c:\autoexec.bat", "display name 2"
'
'    'clsOutlook.DisplayMail     ' To display mail for user to send
'    clsOutlook.SendMail         ' Uhmmm....Duh
'    clsOutlook.CloseOutlook     ' Call to close and clean up
'
'    Set clsOutlook = Nothing    ' free memory
'
'---------------------------------------------------------------
' When you add a Recipient - the address is checked to see if it's valid
' Error messagebox shows if not
'----------------------------------------------------------------
' Code doesn't have error checking for YOU calling the subs in wrong order.
' ie: If you call CreateNewMail before StartOutlook, it just returns with no error.
' Use the order in sample above
'----------------------------------------------------------------
Option Explicit

Private m_objOutlookApplication As Outlook.Application  ' Main Outlook object
Private m_objOutlookNameSpace As Outlook.NameSpace      ' Outlook namespace
Private m_objOutlookMailItem As Outlook.MailItem        ' Outlook mail item
'------------------Procedure StartOutlook-------------------------------
'
' Procedure StartOutlook
' Author: John McGlothlin
' Date: 10/4/01
' Purpose: Get or start instance of Outlook
' Result:
' Input parameters:     String - Outlook Profile
'                       String - Profile Password
'                       Boolean - True displays logon dialog
'                       Boolean - True starts a new session
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub StartOutlook(Optional strProfile As String = "", _
Optional strPassword As String = "", _
Optional blnDisplayDialog As Boolean = False, _
Optional blnNewSession As Boolean = False)
    

Dim m_strProfile As String
Dim m_strPassword As String
Dim m_blnDisplay As Boolean
Dim m_blnNewSession As Boolean

    ' grab parameters - not really necessary but I prefer it
    m_strProfile = strProfile
    m_strPassword = strPassword
    m_blnDisplay = blnDisplayDialog
    m_blnNewSession = blnNewSession
            
    On Error Resume Next
    
    ' grab open instance of Outlook
    Set m_objOutlookApplication = GetObject(, "Outlook.Application")
    
    ' non zero means Outlook wasn't open
    If Err.Number <> 0 Then
    
        ' open Outlook
        Set m_objOutlookApplication = CreateObject("Outlook.Application")
    End If
    
    Err.Clear
    On Error GoTo StartOutlook_Error
    
    ' make sure Outlook was opened
    If (m_objOutlookApplication Is Nothing) Then
        ' do nothing
    Else
        'Return a reference to MAPI.
        Set m_objOutlookNameSpace = m_objOutlookApplication.GetNamespace("MAPI")
        
        ' if no Profile
        If m_strProfile = "" Then
            ' if no Password
            If m_strPassword = "" Then
                ' logon with default Profile and no Password
                m_objOutlookNameSpace.Logon , , m_blnDisplay, m_blnNewSession
            Else
                ' logon with default Profile with Password
                m_objOutlookNameSpace.Logon , m_strPassword, m_blnDisplay, m_blnNewSession
            End If
            
        Else ' we have a profile - just check Password
        
            If m_strPassword = "" Then
                ' logon to specified Profile with no Password
                m_objOutlookNameSpace.Logon m_strProfile, , m_blnDisplay, m_blnNewSession
            Else
                ' logon to specified Profile with Password
                m_objOutlookNameSpace.Logon m_strProfile, m_strPassword, m_blnDisplay, m_blnNewSession
            End If
        End If
    End If
    
    On Error GoTo 0
    Exit Sub
    
StartOutlook_Error:
    MsgBox "Error " & CStr(Err.Number) & " in procedure StartOutlook of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Sub
'------------------Procedure CreateNewMail-------------------------------
'
' Procedure CreateNewMail
' Author: John McGlothlin
' Date: 10/4/01
' Purpose: Creates a new Mail Item
' Result:
' Input parameters:     None
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub CreateNewMail()
   
    On Error GoTo CreateNewMail_Error:
    
    ' check if Outlook App and Outlook Namespace is valid
    If (m_objOutlookApplication Is Nothing) Or (m_objOutlookNameSpace Is Nothing) Then
        ' do nothing
    Else
        ' create new maiitem
        Set m_objOutlookMailItem = m_objOutlookApplication.CreateItem(olMailItem)
    End If
    
    On Error GoTo 0
    Exit Sub
    
CreateNewMail_Error:
    MsgBox "Error " & CStr(Err.Number) & " in procedure CreateNewMail of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Sub

'------------------Procedure CloseOutlook-------------------------------
'
' Procedure CloseOutlook
' Author: John McGlothlin
' Date: 10/4/01
' Purpose: Clean up
' Result:
' Input parameters:     None
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub CloseOutlook()

    Call Class_Terminate
    
End Sub

'------------------Procedure DisplayMail-------------------------------
'
' Procedure DisplayMail
' Author: John McGlothlin
' Date: 10/4/01
' Purpose: Shows the mail to user and lets them send it
' Result:
' Input parameters:     None
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub DisplayMail()

    On Error GoTo Display_Error
    
    ' make sure mailitem is valid
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
    Else
         ' display mail
        m_objOutlookMailItem.Display
    End If
    
    On Error GoTo 0
    Exit Sub
    
Display_Error:
    MsgBox "Error " & CStr(Err.Number) & " in procedure DisplayMail of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Sub

'------------------Procedure SendMail-------------------------------
'
' Procedure SendMail
' Author: John McGlothlin
' Date: 10/4/01
' Purpose: Well.....Duh
' Result:
' Input parameters:     None
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub SendMail()

    On Error GoTo SendMail_Error
    
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
    Else
         ' send mail
        m_objOutlookMailItem.Send
    End If
    
    On Error GoTo 0
    Exit Sub
    
SendMail_Error:
    MsgBox "Error " & CStr(Err.Number) & " in procedure SendMail of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Sub
Private Sub Class_Terminate()
    Set m_objOutlookApplication = Nothing
    Set m_objOutlookNameSpace = Nothing
    Set m_objOutlookMailItem = Nothing
End Sub

'------------------Procedure Attachment-------------------------------
'
' Procedure Attachment
' Author: John McGlothlin
' Date: 10/5/01
' Purpose: Add an attachment to the mailitem
' Result:
' Input parameters:     String with Path to file to attach.
'                       String for Display Name of that attachment
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Sub Attachment(strNewValue As String, Optional strNewDisplayName As String = "")
Dim m_strAttachment As String
Dim m_strDisplayName As String
    
    ' grab parameters
    m_strAttachment = strNewValue
    m_strDisplayName = strNewDisplayName
    
    On Error GoTo Attachment_Error
    
    ' check that MailItem is valid
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
    Else
        ' add Attachment
        With m_objOutlookMailItem.Attachments.Add(m_strAttachment)
            ' add display name if necessary
            If m_strDisplayName = "" Then
                ' do nothing
            Else
                ' display name
                .DisplayName = m_strDisplayName
            End If
        End With
    End If
    
    On Error GoTo 0
    Exit Sub
    
Attachment_Error:
    MsgBox "Error " & CStr(Err.Number) & " in procedure Attachment of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Sub

'------------------Class Properties from here on-------------------------------
'
' Procedure: Various
' Author: John McGlothlin
' Date: 10/5/01
' Purpose:
' Result:
' Input parameters:     None
'
' Output parameters:    None
'
'------------------------------------------------------------------------
Public Property Let Body(ByVal strNewValue As String)
Dim m_strBody As String
    
    m_strBody = strNewValue
    
    On Error GoTo Body_Error
    
    ' make sure mailitem is valid
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
    Else
         ' add the subject - add vbCrLf so any attachments looks good in display
         ' User signature may make this useles...depends on current need
        m_objOutlookMailItem.Body = m_strBody & vbCrLf & vbCrLf
    End If
    
    On Error GoTo 0
    Exit Property
    
Body_Error:
    MsgBox "Error " & CStr(Err.Number) & " in property Body of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Property
Public Property Let Subject(ByVal strNewValue As String)
Dim m_strSubject As String

    m_strSubject = strNewValue
    
    On Error GoTo Subject_Error
    
    ' make sure mailitem is valid
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
    Else
         ' add the subject
        m_objOutlookMailItem.Subject = m_strSubject
    End If
    
    On Error GoTo 0
    Exit Property
Subject_Error:
    MsgBox "Error " & CStr(Err.Number) & " in property Subject of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Property
Public Property Let Recipient_CC(ByVal strNewValue As String)
Dim m_strRecipientCC As String

    m_strRecipientCC = strNewValue
    
    On Error GoTo Recipient_CC_Error
    
    ' check that MailItem is valid
    If (m_objOutlookMailItem Is Nothing) Then
    
        ' do nothing
    Else
         ' add Recipient
        With m_objOutlookMailItem.Recipients.Add(m_strRecipientCC)
            ' Make it CC:
            .Type = olCC
            
            ' resolve it - make sure address is valid
            If Not .Resolve Then
                MsgBox "Unable to resolve address: " & m_strRecipientCC, vbInformation
            End If
        End With
    End If
    
    On Error GoTo 0
    Exit Property
    
Recipient_CC_Error:
    MsgBox "Error " & CStr(Err.Number) & " in property Recipient_CC of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Property
Public Property Let Recipient_BCC(ByVal strNewValue As String)
Dim m_strRecipientBCC As String

    m_strRecipientBCC = strNewValue
    
    On Error GoTo Recipient_BCC_Error
    
    ' check that MailItem is valid
    If (m_objOutlookMailItem Is Nothing) Then
    
        ' do nothing
    Else
         ' add Recipient
        With m_objOutlookMailItem.Recipients.Add(m_strRecipientBCC)
            ' Make it BCC:
            .Type = olBCC
            
            ' resolve it - make sure address is valid
            If Not .Resolve Then
                MsgBox "Unable to resolve address: " & m_strRecipientBCC, vbInformation
            End If
        End With
    End If
    
    On Error GoTo 0
    Exit Property
    
Recipient_BCC_Error:
    MsgBox "Error " & CStr(Err.Number) & " in property Recipient_BCC of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Property
Public Property Let Recipient_TO(ByVal strNewValue As String)
Dim m_strRecipientTO As String

    m_strRecipientTO = strNewValue

    On Error GoTo Recipient_TO_Error
    
    ' check that MailItem is valid
    If (m_objOutlookMailItem Is Nothing) Then
        ' do nothing
        
    Else
        ' add Recipient
        With m_objOutlookMailItem.Recipients.Add(m_strRecipientTO)
            ' Make it TO:
            .Type = olTo
            
            ' resolve it - make sure address is valid
            If Not .Resolve Then
                MsgBox "Unable to resolve address: " & m_strRecipientTO, vbInformation
            End If
        End With
    End If
    
    On Error GoTo 0
    Exit Property
    
Recipient_TO_Error:
    MsgBox "Error " & CStr(Err.Number) & " in property Recipient_TO of the Class Module cOutlookSendMail (filename cOutlookSendMail.cls): " & Err.Description
    On Error GoTo 0
End Property
