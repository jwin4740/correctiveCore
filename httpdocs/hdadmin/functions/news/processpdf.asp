<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

theRedirect = "index.asp"

'##check if the form was submitted and update the user info.
If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

    '## Response.CacheControl = "No-cache"

    '## set global upload object
    Server.ScriptTimeout = 5*60*60
    Set Upload = Server.CreateObject("Persits.Upload")
    Upload.ProgressID = Request.QueryString("PID")

    On Error Resume Next
    Count = Upload.SaveVirtual(hdUploadTempDir)
    If Err.Number <> 0 Then
	    Response.Write "<h4>Error:</h4><div>" & Err.Description & "</div>"
	    Err.Clear
    Else
	    On Error Goto 0
    	
        newsID = Upload.Form("newsID")

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

        With Upload.Form
        
            theRS("catID") = CInt(.Item("catID"))
            theRS("newsIsFile") = .Item("newsIsFile")

            arFields = Array("newsDate", "newsTitle", "newsShortDesc")
            
            For Each qField in arFields
                theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
            Next
            
            Set File = Upload.Files("newsDetails")
            If Not File Is Nothing Then
                If LCase(File.Ext) = ".pdf" Then
                    newsPDFFileName = File.FileName
                    File.CopyVirtual hdPDFDir & newsPDFFileName
                    theRS("newsDetails") = newsPDFFileName
                Else
                    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Only PDF files are allowed for upload."
                End If
                File.Delete
            End If               
            
        End With

        theRS.Update
                        
        newsID = theRS("newsID")
        
        theRS.Close()
        Set theRS = Nothing 

        If Session("hdAdminErrorMsg") = "" Then 
            Session("hdAdminErrorMsg") = "Updates completed for """ & Upload.Form("newsTitle") & """"
        Else
            theRedirect = "editpdf.asp?id=" & newsID
        End If

    End If  '## no upload errors        

End If  '## no errors

response.redirect(theRedirect)

%>