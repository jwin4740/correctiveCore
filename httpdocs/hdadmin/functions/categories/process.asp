<!--#include virtual="/includes/global.asp" -->

<%
If Not hdSecureCMSAdminPageSuperUser Then response.redirect(hdAdminPath)

'##check if the form was submitted and update the user info.

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

'## Response.CacheControl = "No-cache"

    catID = Request.Form("catID")

    Set theRS = Server.CreateObject("ADODB.Recordset")
    theRS.ActiveConnection = hdDSN
    
    '## ADD
    If catID = 0 Then
        theRS.Open "hdCategories", , adOpenKeyset, adLockOptimistic, adCmdTable
        theRS.AddNew
    Else
        '## now update the project data.
	    qSQL = "SELECT * FROM hdCategories WHERE catID = " & catID
	    theRS.Open qSQL, , adOpenStatic, adLockOptimistic
    End If

    With Request.Form
    
        theRS("catTypeID") = CInt(.Item("catTypeID"))
        theRS("catImage") = .Item("catImage")
        theRS("catIsActive") = .Item("catIsActive")
        
        arFields = Array("catName", "catDescription")
        
        For Each qField in arFields
            theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
        Next   
        
    End With

    theRS.Update
                    
    catID = theRS("catID")
    
    theRS.Close()
    Set theRS = Nothing 

    If Session("hdAdminErrorMsg") = "" Then 
        Session("hdAdminErrorMsg") = "Updates completed for """ & Request.Form("catName") & """"
        theRedirect = "index.asp"
    Else
        theRedirect = "edit.asp?id=" & catID
    End If 
    
End If  '## no errors

response.redirect("index.asp")

%>