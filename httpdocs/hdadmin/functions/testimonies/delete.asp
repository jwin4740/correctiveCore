<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtTestimony)

testiID = CInt(request.querystring("id"))

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open hdDSN

qSQL = "DELETE FROM hdTestimony WHERE testiID = " & testiID

objConn.Execute(qSQL)

objConn.Close
Set objConn = Nothing
  
response.redirect("index.asp")

%>