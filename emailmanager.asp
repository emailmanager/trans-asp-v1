<!--#include file="json2.asp" -->
<%

' ##############################################################################
'
' Classic ASP Class for Emailmanager <http://trans.emailmanager.com/>
'   17/06/2011 - v1 - Initial Release
'
' Copyright 2011, James Healey <www.jayhealey.co.uk>
'   Licensed under the MIT License.
'   Redistributions of files must retain the above copyright notice.
'   http://www.opensource.org/licenses/mit-license.php
'
' ##############################################################################
'
' This package also includes & requires json2.asp to serialize/deserialize data
'   Written by Fabio Zendhi Nagao <http://zend.lojcomm.com.br/> @ August 2010
'   https://github.com/nagaozen/asp-xtreme-evolution/blob/master/lib/axe/classes/Parsers/json2.asp
'
' ##############################################################################
'
' Example Usage:
'
' Step 1: 
' Modify the Emailmanager API key (EMAILMANAGER_API_KEY) in this file to be your key.
' If you don't do this, you won't be able to send mail.
'
' NOTE: The EMAILMANAGER_API_TESTMODE constant is set to True by default.
'       True:  Good for testing - you can send to Emailmanager and recieve 
'              successful response, but will NOT actually send the e-mail to recipient.
'       False: Use when going live - this will send to Emailmanager which will then send e-mails.
'
' Step 2:
' Include this file (emailmanager.asp) in your code.
'
' Step 3:
' Use the following example code below to start working with Emailmanager.
'
' There are a couple of functions to add multiple recipients, CC's and BCC's.
'   SetTo: Single recipient
'   SetToCC: Carbon Copy - See single recipient
'   SetToBCC: Blind Carbon Copy - Set single recipient
'   AddTo: Multiple recipients
'   AddToCC: Carbon Copy - Add another recipient
'   AddToBC: Blind Carbon Copy - Add another recipient
'
' <start code example>
'
' Dim EmailManagerEmail: Set EmailManagerEmail = new EmailManager
'
' EmailManagerEmail.SetTo("to@address.com") ' Use when you want to add a single recipient
' EmailManagerEmail.AddTo("to-another@address.com") ' Add another recipient
' EmailManagerEmail.SetFrom("from@address.com")
' EmailManagerEmail.SetSubject("Subject goes here")
' ' Plain text
' EmailManagerEmail.SetTextBody("Body of e-mail goes here")
' ' HTML content
' EmailManagerEmail.SetHTMLBody("<html><body><h1>Body of email goes here.</h1></body></html>")
' EmailManagerEmail.Send()
'
' If (EmailManagerEmail.SendSuccessful()) Then
'   response.write "E-mail was sent!<br />"
'   response.write EmailManagerEmail.GetMessageID &"<br />"
' Else
'   response.write "E-mail failed to send...<br />"
'   response.write EmailManagerEmail.GetErrorCode &"<br />"
'   response.write EmailManagerEmail.GetMessage &"<br />"
' End If
'
' Set EmailManagerEmail = Nothing
'
' <end code example>
'
' ##############################################################################

' ##############################################################################
' Emailmanager API Settings
' http://trans.emailmanager.com/build.html

Const EMAILMANAGER_API_KEY      = "XXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXX"
Const EMAILMANAGER_API_TEST_KEY = "EMAILMANAGER_API_TEST"
Const EMAILMANAGER_API_URL      = "http://trans.emailmanager.com/email"
Const EMAILMANAGER_API_TESTMODE = False

' ##############################################################################
' Emailmanager API Response Codes
' http://trans.emailmanager.com/build.html#api

' Your request did not submit the correct API token in the X-Emailmanager-Server-Token header.
Const EMAILMANAGER_RESPONSE_API = 0

' Validation failed for the email request JSON data that you provided.
Const EMAILMANAGER_RESPONSE_INVALID_EMAIL = 300

' You are trying to send email with a From address that does not have a sender signature.
Const EMAILMANAGER_RESPONSE_SIGNATURE_NOT_SET = 400

' You are trying to send email with a From address that does not have a
' corresponding confirmed sender signature.
Const EMAILMANAGER_RESPONSE_SIGNATURE_NOT_CONFIRMED = 401

' The JSON input you provided is syntactically incorrect.
Const EMAILMANAGER_RESPONSE_JSON_INVALID = 402

' The JSON input you provided is syntactically correct, but still not the one we expect.
Const EMAILMANAGER_RESPONSE_JSON_INCOMPATIBLE = 403

' You ran out of credits.
Const EMAILMANAGER_RESPONSE_NO_CREDITS = 405

' You tried to send to a recipient that has been marked as inactive.
' Inactive recipients are ones that have generated a hard bounce or a spam complaint.
Const EMAILMANAGER_RESPONSE_INVALID_RECIPIENT = 406

' You requested a bounce by ID, but we could not find an entry in our database.
Const EMAILMANAGER_RESPONSE_BOUNCE_NOT_FOUND = 407

' You provided bad arguments as a bounces filter.
Const EMAILMANAGER_RESPONSE_BOUNCE_QRY_BAD = 408

' Your HTTP request does not have the Accept and Content-Type headers set to application/json.
Const EMAILMANAGER_RESPONSE_JSON_REQUIRED = 409

' Your batched request contains more than 500 messages.
Const EMAILMANAGER_RESPONSE_TOO_MANY_REQUESTS = 410

' Text String Response from Test Mode response
Const EMAILMANAGER_RESPONSE_TEXT_TEST_OK = "Test job accepted"

' Text String Response on Successful response
Const EMAILMANAGER_RESPONSE_TEXT_OK = "OK"

' ##############################################################################
' Emailmanager Class

Class EmailManager

    Private EmailTo
    Private EmailToCC
    Private EmailToBCC
    Private EmailFrom
    Private Subject
    Private HTMLBody
    Private TextBody
    Private EmailsSent
    Private isHTML
    Private isSendSuccessful
    
    Private responseText        ' Complete response text from Emailmanager API
    Private responseErrorCode   ' Code number returned from the responseText
    Private responseMessage     ' Code number returned from the responseText
    Private responseMessageID   ' ID of Message from Emailmanager

  ' ############################################################################
  ' Declare initial state variables

    Private Sub Class_Initialize
        isHTML = False
        isSendSuccessful = False
        responseErrorCode = -1
        responseMessageID = -1
        responseMessage = ""
    End Sub

  ' ############################################################################
  ' Set Single Recipient

    Public Function SetTo(p_To)
        EmailTo = Trim(p_To)
    End Function

  ' ############################################################################
  ' Add Multiple Recipients
  
    Public Function AddTo(p_To)
        If (Len(EmailTo) > 0) Then
            EmailTo = EmailTo &","& Trim(p_To)
        Else
            EmailTo = Trim(p_To)
        End If
    End Function

  ' ############################################################################
  ' Set Carbon Copy Recipient

    Public Function SetToCC(p_ToCC)
        EmailToCC = Trim(p_ToCC)
    End Function

  ' ############################################################################
  ' Add Multiple Carbon Copy Recipients

    Public Function AddToCC(p_ToCC)
        If (Len(EmailToCC) > 0) Then
            EmailToCC = EmailToCC &","& Trim(p_ToCC)
        Else
            EmailToCC = Trim(p_ToCC)
        End If
    End Function

  ' ############################################################################
  ' Set Blind Carbon Copy Recipient

    Public Function SetToBCC(p_ToBCC)
        EmailToBCC = Trim(p_ToBCC)
    End Function

  ' ############################################################################
  ' Add Multiple Blind Carbon Copy Recipients

    Public Function AddToBCC(p_ToBCC)
        If (Len(EmailToBCC) > 0) Then
            EmailToBCC = EmailToBCC &","& Trim(p_ToBCC)
        Else
            EmailToBCC = Trim(p_ToBCC)
        End If
    End Function

  ' ############################################################################
  ' Set From Address

    Public Function SetFrom(p_From)
        EmailFrom = Trim(p_From)
    End Function

  ' ############################################################################
  ' Set E-mail Subject

    Public Function SetSubject(p_Subject)
        Subject = Trim(p_Subject)
    End Function

  ' ############################################################################
  ' SetHTMLBody(): Set HTML content for an e-mail

    Public Function SetHTMLBody(p_HTMLBody)
        HTMLBody = Trim(p_HTMLBody)
        isHTML = True
    End Function

  ' ############################################################################
  ' SetTextBody(): Set plain text for an e-mail

    Public Function SetTextBody(p_TextBody)
        TextBody = Trim(p_TextBody)
        isHTML = False
    End Function

  ' ############################################################################
  ' Send(): Put together the data into a JSON string and send to Emailmanager API
    
    Public Function Send()

        ' Declare JSON2.asp object

        dim JSON_Email : set JSON_Email = JSON
        dim JSON_Email_String:  JSON_Email_String = ""
        
        JSON_Email.set "From", EmailFrom
        JSON_Email.set "To", EmailTo
        JSON_Email.set "Subject", Subject

        ' Add Carbon Copy Recipients if set

        If Len(EmailToCC) > 0 Then
            JSON_Email.set "Cc", EmailToCC
        End If
        
        ' Add Blind Carbon Copy Recipients if set

        If Len(EmailToBCC) > 0 Then
            JSON_Email.set "Bcc", EmailToBCC
        End If
        
        ' If the E-mail Body is set to be HTML

        If (True = isHTML) Then
            JSON_Email.set "HTMLBody", HTMLBody
        Else
            JSON_Email.set "TextBody", TextBody
        End If

        JSON_Email_String = JSON.stringify(JSON_Email, null, 2)
        set JSON_Email = nothing

        ' Setup the HTTP Request & Headers for the Emailmanager API
        
        Set xmlHttp = Server.Createobject("MSXML2.ServerXMLHTTP")
        xmlHttp.Open "POST", EMAILMANAGER_API_URL, False
        xmlHttp.setRequestHeader "Accept", "application/json"
        xmlHttp.setRequestHeader "Content-Type", "application/json"

        ' If in Test mode, use the Emailmanager API Test Key
        
        If (True = EMAILMANAGER_API_TESTMODE) Then
            xmlHttp.setRequestHeader "X-Emailmanager-Server-Token", EMAILMANAGER_API_TEST_KEY
        Else
            xmlHttp.setRequestHeader "X-Emailmanager-Server-Token", EMAILMANAGER_API_KEY
        End If

        ' Send the request with the JSON
        
        xmlHttp.Send JSON_Email_String

        ' Recieve response from API

        responseText = xmlHttp.responseText

        xmlHttp.abort()
        set xmlHttp = Nothing

        ' Pass the JSON response on so we can evaluate if Emailmanager sent the message
        
        HandleResponse(responseText)
        
    End Function

  ' ############################################################################
  ' HandleResponse(): Parse the JSON recieved from Emailmanager API

    Private Function HandleResponse(p_responseText)
        
        dim JSON_Response: set JSON_Response = JSON.Parse(p_responseText)

        responseErrorCode = JSON_Response.ErrorCode
        responseMessage   = JSON_Response.Message

        If (CInt(responseErrorCode) = CInt(EMAILMANAGER_RESPONSE_API)) Then
            responseMessageID = JSON_Response.MessageID
        End If
        
        set JSON_Response = nothing

        ' If in test mode & response message is test message OR
        '   if live & response message OK, then we're done!

        If ((EMAILMANAGER_API_TESTMODE = False AND _
             CInt(EMAILMANAGER_RESPONSE_API) = CInt(responseErrorCode) AND _
             0 = StrComp(responseMessage, EMAILMANAGER_RESPONSE_TEXT_OK, 1)) OR _
            (EMAILMANAGER_API_TESTMODE = True AND _
             CInt(EMAILMANAGER_RESPONSE_API) = CInt(responseErrorCode)   AND _
             0 = StrComp(responseMessage, EMAILMANAGER_RESPONSE_TEXT_TEST_OK, 1))) Then
            isSendSuccessful = True
        End If

    End Function

  ' ############################################################################
  ' SendSuccessful(): Return boolean based on the Emailmanager API response

    Public Function SendSuccessful
        SendSuccessful = isSendSuccessful
    End Function

  ' ############################################################################
  ' GetErrorCode(): Return integer, all eventualities are listed at the top of this file

    Public Function GetErrorCode
        GetErrorCode = responseErrorCode
    End Function

  ' ############################################################################
  ' GetMessage(): Return the string message response 

    Public Function GetMessage
        GetMessage = responseMessage
    End Function

  ' ############################################################################
  ' GetMessageID(): Return the string which is the Emailmanager ID for the message that was sent out

    Public Function GetMessageID
        GetMessageID = responseMessageID
    End Function

End Class

%>