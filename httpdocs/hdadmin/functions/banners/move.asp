<%Response.CacheControl = "No-cache"%>
<!--#include virtual="/includes/global.asp" -->
<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtBanners)

theID = Request("ID")
toswap = Request("toswap")
thissort = Request("thissort")

If theID > 0 Then

    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    '## set the sort order of the slot being moved into first
    qSQL = "UPDATE hdBanners SET bannerSortOrder = " & thissort & " WHERE bannerSortOrder = " & toswap
    objConn.Execute(qSQL)

    '## now set the new sort order for the current banner

    qSQL = "UPDATE hdBanners SET bannerSortOrder = " & toswap & " WHERE bannerID = " & theID
    objConn.Execute(qSQL)

    objConn.Close
    Set objConn = Nothing
    
End If

Response.Redirect("index.asp")

%>