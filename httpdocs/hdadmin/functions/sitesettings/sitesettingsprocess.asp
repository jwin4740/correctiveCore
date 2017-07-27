<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

'##check if the form was submitted and update the user info.

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

	Set rsUpdate = Server.CreateObject("ADODB.Recordset")
	rsUpdate.ActiveConnection = hdDSN    

    qSQL = "SELECT * FROM hdSiteSettings WHERE setID = 1"
    rsUpdate.Open qSQL, , adOpenStatic, adLockOptimistic
    
    With Request.Form
    
        setSiteURL = Trim(Left(.Item("setSiteURL"),255))
        If hdFieldValidate(12, setSiteURL, 6, "") Then
            rsUpdate("setSiteURL") = setSiteURL
        Else
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "It appears you entered an invalid domain name.<br />"
        End If

        setDefaultTitle = Trim(Left(.Item("setDefaultTitle"),255))
        If Not hdFieldValidate(11, setDefaultTitle, 0, "") Then
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Please enter only alphanumeric characters in your title.<br />"
        End If
        '## set title even if not "good" characters
        rsUpdate("setDefaultTitle") =  Server.HTMLEncode(setDefaultTitle)
                
        setDefaultMetaKeywords = Trim(Left(.Item("setDefaultMetaKeywords"),255))
        If Not hdFieldValidate(11, setDefaultMetaKeywords, 0, "") Then
            '## give a warning if special characters were used.
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Please avoid using special characters in your Keywords.<br />"
        End If
        '## set the keywords no matter what's entered
        rsUpdate("setDefaultMetaKeywords") = Server.HTMLEncode(setDefaultMetaKeywords)
                
        setDefaultMetaDescription = Trim(.Item("setDefaultMetaDescription"))
        If Not hdFieldValidate(11, setDefaultMetaDescription, 0, "") Then
            '## give a warning if special characters were used.
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Please avoid using special characters in your Description.<br />"
        End If
        '## set the keywords no matter what's entered
        rsUpdate("setDefaultMetaDescription") = Server.HTMLEncode(setDefaultMetaDescription)
    
        '## check if SuperUser is making updates
        If Session("adminRole") = 1 Then
            rsUpdate("setIsActive") = (.Item("setIsActive") = "True")
            rsUpdate("setMgtPages") = (.Item("setMgtPages") = "True")
            rsUpdate("setMgtBlog") = (.Item("setMgtBlog") = "True")
            rsUpdate("setMgtCalendar") = (.Item("setMgtCalendar") = "True")
            rsUpdate("setMgtContact") = (.Item("setMgtContact") = "True")
            rsUpdate("setMgtGallery") = (.Item("setMgtGallery") = "True")
            rsUpdate("setMgtNews") = (.Item("setMgtNews") = "True")
            rsUpdate("setMgtProjects") = (.Item("setMgtProjects") = "True")
            rsUpdate("setMgtBanners") = (.Item("setMgtBanners") = "True")
            rsUpdate("setMgtMailer") = (.Item("setMgtMailer") = "True")
            rsUpdate("setMgtTestimony") = (.Item("setMgtTestimony") = "True")
        End If  '## adminRole = 1
               
    End With
 	
  	rsUpdate.Update
  	
	rsUpdate.Close()
	Set rsUpdate = Nothing 

    If Session("hdAdminErrorMsg") = "" Then
        Session("hdAdminErrorMsg") = "Updates completed for your Default Site Settings."
    Else
        Session("hdAdminErrorMsg") = "Updates completed, some errors occured:<br />" & Session("hdAdminErrorMsg")
    End If
End If

response.redirect("index.asp")

%>