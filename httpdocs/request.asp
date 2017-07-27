<%@LANGUAGE="VBSCRIPT"%>
<html>
<head>
<title>Send Mail</title>
</head>
<body>
<% 
  ' Use an ASP command to check to see if data has been submitted to this page
  If request.form = "" then 
  ' no data was submitted to the form, so display the form to the user
%>	
<%
  Else
%>
<!--#include file="hdCaptchaVerifyInc.asp" -->
<%
    ' Data was submitted to this form, so send the results via email
    Set Mail = Server.CreateObject("Persits.MailSender")
    Mail.Host = "65.61.38.154" 
    Mail.From     = Request.Form("Email") 
    Mail.FromName = Request.Form("Name")
    Mail.AddAddress "jason@clearvisiontank.com"
    Mail.Subject  = "Website Request"
    Mail.Body = Mail.Body & "Full Name:" 		& Request.Form ("Name") & vbCRLF
    Mail.Body = Mail.Body & "Phone:" 			& Request.Form ("Phone") & vbCRLF
    Mail.Body = Mail.Body & "Email:" 			& Request.Form ("Email") & vbCRLF
    Mail.Body = Mail.Body & "message:" 			& Request.Form ("message") & vbCRLF
    Mail.Body = Mail.Body & "Time Sent:"		& FormatDateTime(Now()) & vbCRLF
    Mail.Body = Mail.Body & "Senders IP:"		& Request.ServerVariables("REMOTE_ADDR") & vbCRLF
    On Error Resume Next
    Mail.Send
    If Err = 0 Then
    	Response.Redirect("http://www.clearvisiontank.com/ThankYou.asp")
    else
      Response.Write "Message was not sent, an error encountered: " & Err.Description
    End If
    On Error Goto 0
    Set Mail = Nothing
  end if
%>
</body>
</html>