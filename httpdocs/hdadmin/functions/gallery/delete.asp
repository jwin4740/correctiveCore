<!--#include virtual="/includes/global.asp" -->

<%

Response.CacheControl = "No-cache"

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtGallery)

bContinue = True

galID = CInt(request.querystring("id"))

Set hdRs = Server.CreateObject("ADODB.Recordset")
hdRs.ActiveConnection = hdDSN
hdRs.Source = "SELECT * FROM hdGallery WHERE galID = " & galID
hdRs.Open()

galFileName = hdRs("galFileName")
galThumbFile = hdRs("galThumbFile")

hdRs.Close()
Set hdRS = Nothing

'## check if it's a PDF and delete it if isfile
If galFileName <> "" Then

    hdFullPathAndFile = Server.MapPath("/") & hdGalleryDir & galFileName
    bContinue = hdDeleteFile(hdFullPathAndFile)

    If bContinue Then
        hdFullPathAndFile = Server.MapPath("/") & hdGalleryDir & galThumbFile
        bContinue = hdDeleteFile(hdFullPathAndFile)
    End If
        
End If

If bContinue Then
    Set objConn = Server.CreateObject("ADODB.Connection")
    objConn.Open hdDSN

    qSQL = "DELETE FROM hdGallery WHERE galID = " & galID

    objConn.Execute(qSQL)

    objConn.Close
    Set objConn = Nothing
    
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Gallery Image deleted successfully.<br />"
Else
    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Error Deleting Gallery Image.<br />"
End If
    
response.redirect("index.asp")

%>