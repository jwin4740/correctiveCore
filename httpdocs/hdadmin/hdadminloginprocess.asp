<!--#include virtual="/includes/global.asp" -->

<%

Session("hdAdminErrorMsg") = ""

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

    hdUsername = Trim(Left(Request.Form("hdusername"),20))
    hdPassword = Trim(Left(Request.Form("hdpassword"),20))
    
    If Not hdFieldValidate(7, hdUsername, 3, "") Then
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Invalid Username<br />"
    End If

    If Not hdFieldValidate(9, hdPassword, 3, "") Then
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Invalid Password<br />"
    End If
    
    '## if no problems with the username and password then see if a valid user
    If Session("hdAdminErrorMsg") = "" Then

        Set hdRs = Server.CreateObject("ADODB.Recordset")
        hdRs.ActiveConnection = hdDSN
        hdRs.Source = "SELECT hdAdmin.*, hdAdminRole.admrlsName FROM hdAdmin INNER JOIN hdAdminRole ON hdAdmin.adminRole = hdAdminRole.admrlsID WHERE adminIsEnabled = 1 AND adminUsername = '" & hdUsername & "' AND adminPassword = '" & hdPassword & "'"
        hdRs.Open()

        If hdRs.EOF Then
            Session("hdAdminIsLoggedIn") = ""
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Login Failed.<br />"
        Else
            '## set all the details for the logged in user
            Session("hdAdminIsLoggedIn") = "Yes"
            Session("adminPersonName") = hdRs("adminPersonName")
            Session("adminUsername") = hdRs("adminUsername")
            Session("adminEmail") = hdRs("adminEmail")
            Session("adminRole") = hdRs("adminRole")
            Session("adminID") = hdRs("adminID")
            Session("admrlsName") = hdRs("admrlsName")
           
        End If
       
        hdRs.Close()
        Set hdRs = Nothing
    
    End If   
End If

Response.Redirect("index.asp")

%>