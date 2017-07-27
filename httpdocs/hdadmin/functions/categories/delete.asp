<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

If Not hdSecureCMSAdminPageSuperUser Then response.redirect(hdAdminPath)

catID = CInt(request.querystring("id"))

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open hdDSN

qSQL = "DELETE FROM hdCategories WHERE catID = " & catID

objConn.Execute(qSQL)

objConn.Close
Set objConn = Nothing

Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Category deleted successfully.<br />"
    
response.redirect("index.asp")

%>