<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtBanners)

bContinue = False

bannerID = CInt(Request("id"))
bannerImage = Request("imagename")
bannerSortOrder = Request("sortorder")

'## check if it's a PDF and delete it if isfile
If bannerID > 0 Then

    hdFullPathAndFile = Server.MapPath("/") & hdBannerDir & bannerImage
    bContinue = hdDeleteFile(hdFullPathAndFile)

End If

If bContinue Then
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    qSQL = "DELETE FROM hdBanners WHERE bannerID = " & bannerID
    objConn.Execute(qSQL)

    '## now reorder the sort order

    qSQL = "SELECT bannerID, bannerSortOrder FROM hdBanners WHERE bannerSortOrder >= " & bannerSortOrder & " ORDER BY bannerSortOrder"
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open qSQL, hdDSN

    While Not rs.EOF
        newSortOrder = rs("bannerSortOrder") - 1
        qSQL = "UPDATE hdBanners SET bannerSortOrder = " & newSortOrder & " WHERE bannerID = " & rs("bannerID")
        objConn.Execute(qSQL)
        rs.MoveNext
    Wend

    rs.Close()
    Set rs = Nothing

    objConn.Close
    Set objConn = Nothing
    
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Banner deleted successfully.<br />"
Else
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Error Deleting Banner file.<br />"
End If
    
response.redirect("index.asp")

%>