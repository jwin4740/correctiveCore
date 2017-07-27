<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

bContinue = True

newsID = CInt(request.querystring("id"))

Set hdRs = Server.CreateObject("ADODB.Recordset")
hdRs.ActiveConnection = hdDSN
hdRs.Source = "SELECT * FROM hdNews WHERE newsID = " & newsID
hdRs.Open()

newsIsFile = hdRs("newsIsFile")
newsDetails = hdRs("newsDetails")

hdRs.Close()
Set hdRS = Nothing

'## check if it's a PDF and delete it if isfile
If newsIsFile Then

    hdFullPathAndFile = Server.MapPath("/") & hdPDFDir & newsDetails
    bContinue = hdDeleteFile(hdFullPathAndFile)

End If

If bContinue Then
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    qSQL = "DELETE FROM hdNews WHERE newsID = " & newsID

    objConn.Execute(qSQL)

    objConn.Close
    Set objConn = Nothing
    
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Article deleted successfully.<br />"
Else
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Error Deleting file.<br />"
End If
    
response.redirect("index.asp")

%>