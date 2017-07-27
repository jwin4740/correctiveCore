<!--#include virtual="/includes/global.asp" -->

<%

webpgID = CInt(request.querystring("id"))
hdTheFileName = Trim(request.querystring("pg"))

If hdSecureCMSAdminPageSuperUser Then

    '##check if the form was submitted and update the user info.

    If hdCreateDriverPage(hdTheFileName) = 1 Then
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & hdTheFileName & " created successfully."
    Else    
        Session("hdAdminErrorMsg") = "There was a problem creating " & hdTheFileName & "<br />" & Session("hdAdminErrorMsg")
    End If  '## no errors

End If

response.redirect("edit.asp?id="&webpgID)

%>