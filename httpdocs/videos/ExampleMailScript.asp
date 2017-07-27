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
    Mail.Host = "127.0.0.1" 
    Mail.From     = Request.Form("email") 
    Mail.FromName = Request.Form("firstname") & " " & Request.Form("lastname")
    Mail.AddAddress "jed@tychem.net"
    Mail.Subject = "Contact Request"
    Mail.Body = Mail.Body & "First Name:" 		& Request.Form ("firstname") & vbCRLF
    Mail.Body = Mail.Body & "Last Name:" 		& Request.Form ("lastname") & vbCRLF
    Mail.Body = Mail.Body & "Email:" 			& Request.Form ("email") & vbCRLF
	Mail.Body = Mail.Body & "Phone:" 			& Request.Form ("phone") & vbCRLF
    Mail.Body = Mail.Body & "Comments:" 		& Request.Form ("message") & vbCRLF
    Mail.Body = Mail.Body & "Time Sent:"		& FormatDateTime(Now()) & vbCRLF
    Mail.Body = Mail.Body & "Senders IP:"		& Request.ServerVariables("REMOTE_ADDR") & vbCRLF
    On Error Resume Next
    Mail.Send
    If Err = 0 Then
    	Response.Redirect("http://tychem.net/ThankYou.asp")
    else
      Response.Write "Message was not sent, an error encountered: " & Err.Description
    End If
    On Error Goto 0
    Set Mail = Nothing
  end if
%>
</body>
</html>