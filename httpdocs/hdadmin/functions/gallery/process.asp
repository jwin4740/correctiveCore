<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtGallery)

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
    	
        galID = Upload.Form("galID")

        Set theRS = Server.CreateObject("ADODB.Recordset")
        theRS.ActiveConnection = hdDSN
        
        '## ADD
        If galID = 0 Then
            theRS.Open "hdGallery", , adOpenKeyset, adLockOptimistic, adCmdTable
            theRS.AddNew
        Else
            '## now update the project data.
	        qSQL = "SELECT * FROM hdGallery WHERE galID = " & galID
	        theRS.Open qSQL, , adOpenStatic, adLockOptimistic
        End If

        With Upload.Form
        
            '## save INT's
            arFields = Array("catID", "galSortOrder")
            For Each qField in arFields
                theRS(qField) = CInt(.Item(qField))
            Next
            
            '## save strings
            arFields = Array("galName", "galDescription")
            For Each qField in arFields
                theRS(qField) = Server.HTMLEncode(Trim(.Item(qField)))
            Next
            
            '## upload the file
            Set File = Upload.Files("galFileName")
            If Not File Is Nothing Then
                If LCase(File.Ext) = ".jpg" Or LCase(File.Ext) = ".gif" Then
                    galleryImageFileName = File.FileName
                    File.CopyVirtual hdGalleryDir & galleryImageFileName
                    Set Jpeg = Server.CreateObject("Persits.Jpeg")
                    Jpeg.Open(File.Path)         
                    Jpeg.Width = hdGalleryW
                    Jpeg.Height = Jpeg.OriginalHeight * hdGalleryW / Jpeg.OriginalWidth ' Resize, preserve aspect ratio
                    Jpeg.Save Server.MapPath("/") & hdGalleryDir & galleryImageFileName
                    '## now save thumb
                    Jpeg.Open( File.Path )         
                    Jpeg.Width = hdGalleryThumbW
                    Jpeg.Height = hdGalleryThumbH
                    galleryThumbFileName = "sm_" & galleryImageFileName
                    Jpeg.Save Server.MapPath("/") & hdGalleryDir & galleryThumbFileName
                    Set Jpeg = Nothing
                    
                    theRS("galFileName") = galleryImageFileName
                    theRS("galThumbFile") = galleryThumbFileName
                Else
                    Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Only JPG or GIF files are allowed for Gallery uploads."
                End If
                File.Delete
            End If               
            
        End With

        theRS.Update
                        
        galID = theRS("galID")
        
        theRS.Close()
        Set theRS = Nothing 

        If Session("hdAdminErrorMsg") = "" Then 
            Session("hdAdminErrorMsg") = "Updates completed for """ & Upload.Form("galName") & """"
        Else
            theRedirect = "edit.asp?id=" & galID
        End If

    End If  '## no upload errors        

End If  '## no errors

response.redirect(theRedirect)

%>