<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

Call hdSecureCMSAdminPage

adminID = CInt(request.querystring("id"))

If adminID = 1 Then
    
    hdSubject = "User Edit Security Alert"
    hdBody = Session("adminPersonName") & " attempted to delete ID#1" & vbCrLf & vbCrLf & _
        "adminUsername: " & Session("adminUsername") & vbCrLf & _
        "adminRole: " & Session("adminRole")
        
    Call hdSendEmail(CStr("cms@" & Session("setSiteURL")), hdMasterUsersEmailAddr, hdSubject, hdBody, False)
    
ElseIf adminID = Session("adminID") Then

    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "You cannot delete your own username.<br />"

Else
    '##check if the form was submitted and update the user info.

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    qSQL = "DELETE FROM hdAdmin WHERE adminID = " & adminID & " AND adminRole >= " & Session("adminRole") 

    objConn.Execute(qSQL)

    objConn.Close
    Set objConn = Nothing

End If

response.redirect("index.asp")

%>