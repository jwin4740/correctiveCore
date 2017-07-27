<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtPages)

theRedirect = "index.asp"

'##check if the form was submitted and update the user info.
If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

'## Response.CacheControl = "No-cache"

    webpgID = Request.Form("webpgID")

    If webpgID = 0 And Session("adminRole") <> 1 Then
        Session("hdAdminErrorMsg") = "Not authorized to add pages."
        Response.Redirect("index.asp")
    End If
   
    Set theRS = Server.CreateObject("ADODB.Recordset")
    theRS.ActiveConnection = hdDSN
    
    '## ADD
    If webpgID = 0 Then
        theRS.Open "hdWebPage", , adOpenKeyset, adLockOptimistic, adCmdTable
        theRS.AddNew
    Else
        '## now update the project data.
	    qSQL = "SELECT * FROM hdWebPage WHERE webpgID = " & webpgID
	    theRS.Open qSQL, , adOpenStatic, adLockOptimistic
    End If

    With Request.Form
    
        webpgName = Trim(.Item("webpgName"))
        If webpgName = "" Then
            theRS("webpgName") = webpgName
        ElseIf hdFieldValidate(11, webpgName, 0, "") Then
            theRS("webpgName") = webpgName
        Else
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "WebPage name should only be alphanumeric (some special characters are allowed)<br />"
        End If 
    
        webpgTitle = Trim(.Item("webpgTitle"))
        If webpgTitle = "" Then
            theRS("webpgTitle") = webpgTitle
        ElseIf hdFieldValidate(11, webpgTitle, 0, "") Then
            theRS("webpgTitle") = webpgTitle
        Else
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Page Title should only be alphanumeric (some special characters are allowed)<br />"
        End If    

        webpgMetaKeywords = Trim(.Item("webpgMetaKeywords"))
        If webpgMetaKeywords = "" Then
            theRS("webpgMetaKeywords") = webpgMetaKeywords
        ElseIf hdFieldValidate(11, webpgMetaKeywords, 0, "") Then
            theRS("webpgMetaKeywords") = webpgMetaKeywords
        Else
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Meta Keywords should only be alphanumeric (some special characters are allowed)<br />"
        End If    

        webpgMetaDescription = Trim(.Item("webpgMetaDescription"))
        If webpgMetaDescription = "" Then
            theRS("webpgMetaDescription") = webpgMetaDescription
        ElseIf hdFieldValidate(11, webpgMetaDescription, 0, "") Then
            theRS("webpgMetaDescription") = webpgMetaDescription
        Else
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Meta Description should only be alphanumeric (some special characters are allowed)<br />"
        End If    

        theRS("webpgContent") = Trim(.Item("webpgContent"))

        If Session("adminRole") = 1 Then 
            '## add file to the website for the driver
            webpgFileName = Trim(.Item("webpgFileName"))
            OldWebpgFileName = Trim(.Item("OldWebpgFileName"))
            
            If webpgFileName = "" Then
                Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Missing Page Filename; Cannot be Empty.<br />"
            Else
                If webpgID = 0 Then
                    If hdCreateDriverPage(webpgFileName) = 1 Then
                        theRS("webpgFileName") = webpgFileName
                    Else
                        theRS("webpgFileName") = ""
                        Session("hdAdminErrorMsg") = "Error creating file: " & webpgFileName & "<br />" & Session("hdAdminErrorMsg")
                    End If
                '## check if renaming the page file.
                ElseIf OldWebpgFileName <> webpgFileName Then 
                    If hdCreateDriverPage(webpgFileName) = 1 Then
                        theRS("webpgFileName") = webpgFileName
                        If Not hdDeleteDriverPage(OldWebpgFileName) Then
                            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Error deleting old page: " & OldWebpgFileName & "<br />"
                        End If
                    Else
                        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Error creating new file for rename to: " & webpgFileName & "<br />"
                    End If
                End If
            End If
        End If  '## superuser only

    End With

    theRS.Update
                    
    webpgID = theRS("webpgID")
    
    theRS.Close()
    Set theRS = Nothing 

    If Session("hdAdminErrorMsg") = "" Then
        Session("hdAdminErrorMsg") = "Updates completed for """ & Request.Form("webpgName") & """"
    Else
        theRedirect = "edit.asp?id=" & webpgID
    End If
    
End If  '## no errors

response.redirect(theRedirect)

%>