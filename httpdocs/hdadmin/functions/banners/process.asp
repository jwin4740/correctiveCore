<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtBanners)

'##check if the form was submitted and update the user info.

theRedirect = "index.asp"
       
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
    	
        bannerID = Upload.Form("bannerID")

        Set theRS = Server.CreateObject("ADODB.Recordset")
        theRS.ActiveConnection = hdDSN
        
        '## ADD
        If bannerID = 0 Then
            theRS.Open "hdBanners", , adOpenKeyset, adLockOptimistic, adCmdTable
            theRS.AddNew
        Else
            '## now update the project data.
	        qSQL = "SELECT * FROM hdBanners WHERE bannerID = " & bannerID
	        theRS.Open qSQL, , adOpenStatic, adLockOptimistic
        End If

        With Upload.Form
        
            '## save INT's
            arFields = Array("catID", "webpgID", "bannerSortOrder")
            For Each qField in arFields
                theRS(qField) = CInt(.Item(qField))
            Next
            
            '## save strings
            arFields = Array("bannerName", "bannerImageAltTag", "bannerURL", "bannerDescription")
            For Each qField in arFields
                theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
            Next
            
            '## upload the file
            Set File = Upload.Files("bannerImage")
            If Not File Is Nothing Then
                If LCase(File.Ext) = ".jpg" Or LCase(File.Ext) = ".gif" Then
                    bannerImageFileName = File.FileName
                    File.CopyVirtual hdBannerDir & bannerImageFileName
                    Set Jpeg = Server.CreateObject("Persits.Jpeg")
                    Jpeg.Open(File.Path)         
                    Jpeg.Width = hdBannerW
                    Jpeg.Height = hdBannerH
                    Jpeg.Save Server.MapPath("/") & hdBannerDir & bannerImageFileName
                    Set Jpeg = Nothing
                    
                    theRS("bannerImage") = bannerImageFileName
                Else
                    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Only JPG or GIF files are allowed for banner uploads."
                End If
                File.Delete
            End If               
            
        End With

        theRS.Update
                        
        bannerID = theRS("bannerID")
        
        theRS.Close()
        Set theRS = Nothing 

        If Session("hdAdminErrorMsg") = "" Then 
            Session("hdAdminErrorMsg") = "Updates completed for """ & Upload.Form("bannerName") & """"
        Else
            theRedirect = "edit.asp?id=" & bannerID
        End If

    End If  '## no upload errors        

End If  '## no errors

response.redirect(theRedirect)

%>