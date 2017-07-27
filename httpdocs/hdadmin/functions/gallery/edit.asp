<!--#include virtual="/includes/global.asp" -->

<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtGallery)

galID = CInt(Request.QueryString("id"))

hdAdminPageTitle = "Manage Gallery Image"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

galName = ""
galFileName = ""
galThumbFile = ""
galDescription = ""
galIsActive = True
galSortOrder = 0
hdTitleBannerText = "Adding Gallery Image."
catID = 0

If galID > 0 Then    

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdGallery WHERE galID = " & galID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../gallery/")
    
    galName = hdRs("galName")
    galFileName = hdRs("galFileName")
    galThumbFile = hdRs("galThumbFile")
    galDescription = hdRs("galDescription")
    galIsActive = hdRs("galIsActive")
    galSortOrder = hdRs("galSortOrder")
    catID = hdRs("catID")
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "Editing Gallery Image: " & galName
    
End If  '## galID > 0

'## used in the aspuploadjs
hdThisFormName = "hdEditForm"
hdThisFileFieldName = "galFileName"

%>
<!-- #include virtual="/includes/lightbox.asp" -->
<!-- #include file="../../includes/aspuploadinit.asp" -->
<!-- #include file="../../includes/aspuploadjs.asp" -->
<script type="text/javascript" language="javascript">
    function DeletePage(theID)
    {
        if (confirm("Delete This Image?"))
        {
	        location.href("delete.asp?id="+theID);
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" enctype="multipart/form-data" name="hdEditForm" action="process.asp?<%=PID%>" >
    <input type="hidden" name="galID" id="galIDID" value="<%=galID%>" />
    <input type="hidden" name="galSortOrder" id="galSortOrderID" value="<%=galSortOrder%>" />
    <input type="hidden" name="galIsActive" id="galIsActiveID" value="<%=galIsActive%>" />
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td class="titleborder"><span id="pagetitle"><%=hdTitleBannerText%></span></td>
      </tr>
    </table>
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td id="hdEditData1" class="border" width="95%" valign="top" style="display: block;">
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                    <td>
                        Category - choose category for this image:</td>
                </tr>
                <tr>
                  <td><%Call hdGetCategoriesDDL(hdGALLERYcat, catID, "")%></td>
                </tr>
    
                <tr>
                    <td>
                        Image Name - a title/name of the image (can be left blank):
                </tr>
                <tr>
                  <td><input type="text" size="70" name="galName" id="galNameID" value="<%=galName%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Short Description - details or information about the image - up to 255 characters:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="galDescription" id="galDescriptionID" value="<%=galDescription%>" maxlength="255" class="style" /></td>
                </tr>
                <%If galID > 0 Then %>
                <tr>
                    <td>
                        View Existing File:
                    </td>
                </tr>
                <tr>
                  <td>
                    <a href="<%=hdGalleryDir%><%=galFileName%>" rel="lightbox[roadtrip]" title="<%=galDescription%>"><img src="<%=hdGalleryDir%><%=galThumbFile%>" width="<%=hdGalleryThumbW%>" height="<%=hdGalleryThumbH%>" border="0" alt="<%=galDescription%>" /></a>
                  </td>
                </tr>
                <%End If %>                
                <tr>
                    <td>
                        Upload New Image:
                    </td>
                </tr>
                <tr>
                  <td><input type="file" size="55" name="galFileName" id="galFileNameID" class="style" /></td>
                </tr>
            </table>
        </td>
      </tr>
      <tr>
        <td id="hdEditDataButtons" class="border" width="95%" valign="top" style="display: block;">
           <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td colspan="2">
                    <img src="<%=hdAdminPath%>images/save.jpg" onclick="javascript:document.hdEditForm.submit();return ShowProgress();" width="126" height="33" hspace="1" vspace="1" alt="Save" style="cursor:hand;" />
                    <img src="<%=hdAdminPath%>images/cancel.jpg" onclick="javascript:location.href='index.asp'" width="126" height="33" hspace="1" vspace="1" alt="Cancel" style="cursor:hand;" />
                    <%If galID > 0 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeletePage(<%=galID%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
                    <%End If %>
                  </td>
                </tr>
            </table>        
        </td>
      </tr>
    </table>
    </form>
    </td>
  </tr>
</table>

<%
End Sub '## PageContent
%>