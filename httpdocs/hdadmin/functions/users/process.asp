<!--#include virtual="/includes/global.asp" -->

<%

Call hdSecureCMSAdminPage

'##check if the form was submitted and update the user info.

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

    adminID = CInt(Request.Form("adminID"))

    '## don't allow a user to add a username with a role less then what they're logged in as.
    If CInt(Request.Form("adminRole")) < CInt(Session("adminRole")) Then
        '## send an email about missing or denied user access.
        hdSubject = "User Edit Security Alert"
        hdBody = Session("adminPersonName") & " attempted to add user with Role: " & CInt(Request.Form("adminRole")) & vbCrLf & vbCrLf & _
            "adminUsername: " & Session("adminUsername") & vbCrLf & _
            "adminRole: " & Session("adminRole")
            
        Call hdSendEmail(CStr("cms@" & Session("setSiteURL")), hdMasterUsersEmailAddr, hdSubject, hdBody, False)
        Response.Redirect("index.asp")
    End If

    Set rsUpdate = Server.CreateObject("ADODB.Recordset")
    rsUpdate.ActiveConnection = hdDSN
    
    '## ADD
    If adminID = 0 Then
        rsUpdate.Open "hdAdmin", , adOpenKeyset, adLockOptimistic, adCmdTable
        rsUpdate.AddNew
    Else
        '## or update the user.
	    qSQL = "SELECT * FROM hdAdmin WHERE adminID = " & adminID & " AND adminRole >= " & Session("adminRole")
	    rsUpdate.Open qSQL, , adOpenStatic, adLockOptimistic
    End If
    
    With Request.Form

        '## get the active status first and change it along the way if there's a problem
        
        adminIsEnabled = .Item("adminIsEnabled")
        
        adminPersonName = Trim(.Item("adminPersonName"))
        If hdFieldValidate(10, adminPersonName, 3, "") Then
            rsUpdate("adminPersonName") = adminPersonName
            If adminID = Session("adminID") Then Session("adminPersonName") = adminPersonName
        Else
            adminIsEnabled = False
            rsUpdate("adminPersonName") = "- invalid name -"
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Your name must be at least 3 alphanumeric characters only<br />"
        End If

        If Instr(adminPersonName, "- invalid") Then
            adminIsEnabled = False
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Please update the Persons Name!<br />"
        End If        

        '## make sure the email is formatted correctly and not assigned to anyone else
        adminEmail = Trim(.Item("adminEmail"))
        If hdIsValidEmailAddress(adminEmail) Then
            '## if new user then make sure it's not already assigned.
            If adminID = 0 Then
                If hdCheckIfAdminEmailExists(adminEmail) Then
                    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Email address " & adminEmail & " already assigned to another user.<br />"
                    adminIsEnabled = False
                Else
                    rsUpdate("adminEmail") = adminEmail
                End If   
            Else
                '## only update email if changed on existing user.
                If .Item("oldadminEmail") <> adminEmail Then
                    If hdCheckIfAdminEmailExists(adminEmail) Then
                        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Email address " & adminEmail & " already assigned to another user.<br />"
                        adminIsEnabled = False
                    Else
                        rsUpdate("adminEmail") = adminEmail
                        If adminID = Session("adminID") Then Session("adminEmail") = adminEmail
                    End If
                 End If '## changing email address
            End If  '## adminID = 0
        Else
            adminIsEnabled = False
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Invalid Email Address: " & adminEmail & "<br />"
        End If
        
        '## make sure the username is valid and doesn't exist.
        adminUsername = Trim(Left(.Item("adminUsername"),20))
        
        '## check if new user
        If adminID = 0 Then
            adminUsername = hdVerifyAdminUsernameUnique(adminUsername)
        Else
            '## if the username changed then validate
            If .Item("oldadminUsername") <> adminUsername Then 
                adminUsername = hdVerifyAdminUsernameUnique(adminUsername)
            End If
        End If
        
        rsUpdate("adminUsername") = adminUsername
        If adminID = Session("adminID") Then Session("adminUsername") = adminUsername

        If Instr(adminUsername, "AdminTemp") Then
            adminIsEnabled = False
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Please update your username!<br />"
        End If
        
        adminPassword = Trim(Left(.Item("adminPassword"),20))
        If adminPassword = "" Then
            '## if new user and blank password, don't allow
            If adminID = 0 Then
                adminIsEnabled = False
                adminPassword = "setpswd654321"
                rsUpdate("adminPassword") = adminPassword
                Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Your password must be 5-20 characters.  Character values can be alphanumeric and some special characters.  Cannot be blank.<br />"
            End If
        Else
            If hdFieldValidate(9, adminPassword, 5, "") Then
                rsUpdate("adminPassword") = adminPassword
            Else
                adminIsEnabled = False
                Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Your password must be 5-20 characters.  Character values can be alphanumeric and some special characters.<br />"
            End If
        End If
        
        rsUpdate("adminRole") = CInt(.Item("adminRole"))
        If adminID = Session("adminID") Then Session("adminRole") = CInt(.Item("adminRole"))
        
        If Not adminIsEnabled Then Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "User " & adminPersonName & " is Disabled.<br />"
        rsUpdate("adminIsEnabled") = adminIsEnabled
        
    End With
 	
  	rsUpdate.Update
  	
  	adminID = rsUpdate("adminID")
  	
	rsUpdate.Close()
	Set rsUpdate = Nothing 

    If Session("hdAdminErrorMsg") = "" Then
        Session("hdAdminErrorMsg") = "User updates completed."
        hdTheRedirect = "index.asp"
    Else
        hdTheRedirect = "edit.asp?id=" & adminID
    End If
End If

response.redirect(hdTheRedirect)


%>