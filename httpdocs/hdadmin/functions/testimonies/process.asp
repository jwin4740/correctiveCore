<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtTestimony)

'##check if the form was submitted and update the user info.
       
If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

    testiID = Request.Form("testiID")

    Set theRS = Server.CreateObject("ADODB.Recordset")
    theRS.ActiveConnection = hdDSN
    
    '## ADD
    If testiID = 0 Then
        theRS.Open "hdTestimony", , adOpenKeyset, adLockOptimistic, adCmdTable
        theRS.AddNew
    Else
        '## now update the project data.
        qSQL = "SELECT * FROM hdTestimony WHERE testiID = " & testiID
        theRS.Open qSQL, , adOpenStatic, adLockOptimistic
    End If

    With Request.Form
    
        theRS("catID") = CInt(.Item("catID"))
        theRS("testiIsActive") = .Item("testiIsActive")

        arFields = Array("testiTitle", "testiAuthor", "testiQuote")
        
        For Each qField in arFields
            theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
        Next
        
    End With

    theRS.Update
                    
    testiID = theRS("testiID")
    
    theRS.Close()
    Set theRS = Nothing 

    If Session("hdAdminErrorMsg") = "" Then 
        Session("hdAdminErrorMsg") = "Updates completed for """ & Request.Form("testiTitle") & """"
        theRedirect = "index.asp"
    Else
        theRedirect = "edit.asp?id=" & testiID
    End If

End If  '## no errors

response.redirect(theRedirect)

%>