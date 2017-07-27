<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

'##check if the form was submitted and update the user info.

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

'## Response.CacheControl = "No-cache"

    newsID = Request.Form("newsID")
   
    Set theRS = Server.CreateObject("ADODB.Recordset")
    theRS.ActiveConnection = hdDSN
    
    '## ADD
    If newsID = 0 Then
        theRS.Open "hdNews", , adOpenKeyset, adLockOptimistic, adCmdTable
        theRS.AddNew
    Else
        '## now update the project data.
	    qSQL = "SELECT * FROM hdNews WHERE newsID = " & newsID
	    theRS.Open qSQL, , adOpenStatic, adLockOptimistic
    End If

    With Request.Form
    
        theRS("catID") = CInt(.Item("catID"))
        theRS("newsIsFile") = .Item("newsIsFile")

        arFields = Array("newsDate", "newsTitle", "newsShortDesc", "newsMetaKeywords", "newsMetaDescription")
        
        For Each qField in arFields
            theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
        Next   
        
        theRS("newsDetails") = Trim(.Item("newsDetails"))

    End With

    theRS.Update
                    
    newsID = theRS("newsID")
    
    theRS.Close()
    Set theRS = Nothing 

    If Session("hdAdminErrorMsg") = "" Then 
        Session("hdAdminErrorMsg") = "Updates completed for """ & Request.Form("newsTitle") & """"
        theRedirect = "index.asp"
    Else
        theRedirect = "edit.asp?id=" & newsID
    End If 
    
End If  '## no errors

response.redirect("index.asp")

%>