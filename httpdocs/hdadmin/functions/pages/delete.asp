<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

webpgID = CInt(request.querystring("id"))
webpgFileName = Trim(request.querystring("pg"))

If hdSecureCMSAdminPageSuperUser And webpgID <> 1 Then

    '##check if the form was submitted and update the user info.

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    qSQL = "DELETE FROM hdWebPage WHERE webpgID = " & webpgID & " AND webpgFileName = '" & webpgFileName & "'"

    objConn.Execute(qSQL)

    objConn.Close
    Set objConn = Nothing

    If hdDeleteDriverPage(webpgFileName) Then
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & webpgFileName & " deleted successfully.<br />"
    Else    
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "There was a problem deleting the file " & hdTheFileName & ".<br />"
    End If  '## no errors

End If

response.redirect("index.asp")

%>