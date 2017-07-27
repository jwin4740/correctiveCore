<!--#include virtual="/includes/global.asp" -->

<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtBanners)

bannerID = CInt(Request.QueryString("id"))

hdAdminPageTitle = "Manage Banners"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

If bannerID = 0 Then 

    bannerName = ""
    bannerImage = ""
    bannerImageAltTag = ""
    bannerURL = ""
    bannerDescription = ""
    bannerSortOrder = GetLastImageNumber()
    webpgID = 0
    catID = 0
    hdTitleBannerText = "Adding Banner."

Else   
    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdBanners WHERE bannerID = " & bannerID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../banners/")
    
    bannerName = hdRs("bannerName")
    bannerImage = hdRs("bannerImage")
    bannerImageAltTag = hdRs("bannerImageAltTag")
    bannerURL = hdRs("bannerURL")
    bannerDescription = hdRs("bannerDescription")
    bannerSortOrder = CInt(hdRs("bannerSortOrder"))
    webpgID = hdRs("webpgID")
    catID = hdRs("catID")
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "Editing Banner: " & bannerName
    
End If  '## bannerID = 0

'## used in the aspuploadjs
hdThisFormName = "hdEditForm"
hdThisFileFieldName = "bannerImage"

%>
<!-- #include file="../../includes/aspuploadinit.asp" -->
<!-- #include file="../../includes/aspuploadjs.asp" -->
<script type="text/javascript" language="javascript">
function DeleteBanner(id, imagename, sortorder)
{
	if (confirm("Delete Banner ID: "+id+"?"))
	{
		location.href("delete.asp?id="+id+"&imagename="+imagename+"&sortorder="+sortorder);
	}
}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" enctype="multipart/form-data" name="hdEditForm" action="process.asp?<%=PID%>" >
    <input type="hidden" name="bannerID" id="bannerIDID" value="<%=bannerID%>" />
    <input type="hidden" name="catID" id="catIDID" value="<%=catID%>" />
    <input type="hidden" name="bannerSortOrder" id="bannerSortOrderID" value="<%=bannerSortOrder%>" />
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
                    <td>Current Position in the Banner List: <b>#<%=bannerSortOrder%></b></td>
                </tr>
                <tr>
                    <td>Banner Name:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="bannerName" id="bannerNameID" value="<%=bannerName%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>Banner Description:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="bannerDescription" id="bannerDescriptionID" value="<%=bannerDescription%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>Banner URL - Link to website for this banner:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="bannerURL" id="bannerURLID" value="<%=bannerURL%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Web Page - If this banner is for a specific page, then choose the page name from the list:</td>
                </tr>
                <tr>
                  <td><%Call hdGetMainPagesIDDDL(webpgID, "")%></td>
                </tr>                                
                <tr>
                    <td>
                        Image &quot;Alt&quot; Tag - used for search engines and when hoving over imate:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="bannerImageAltTag" id="bannerImageAltTagID" value="<%=bannerImageAltTag%>" maxlength="255" class="style" /></td>
                </tr>
                <%If bannerID > 0 Then %>
                <tr>
                    <td>
                        Current Banner Image:
                    </td>
                </tr>
                <tr>
                  <td><img src="<%=hdBannerDir%><%=bannerImage%>" alt="<%=bannerImageAltTag%>" title="<%=bannerImageAltTag%>" width="<%=hdBannerW %>" height="<%=hdBannerH %>" /></td>
                </tr>
                <%End If %>                
                <tr>
                    <td>
                        Upload New Banner Image (<%=hdBannerW %> x <%=hdBannerH %> pixels):
                    </td>
                </tr>
                <tr>
                  <td><input type="file" size="55" name="bannerImage" id="bannerImageID" value="<%=bannerImage%>" class="style" /></td>
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
                    <%If bannerID > 0 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeleteBanner(<%=bannerID%>, '<%=bannerImage%>', <%=bannerSortOrder%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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

Function GetLastImageNumber()
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.ActiveConnection = hdDSN
    rs.Source = "SELECT TOP 1 bannerSortOrder FROM hdBanners ORDER BY bannerSortOrder DESC"
    rs.Open()

    If rs.EOF then
        LastImageNumber = 1
    Else
        LastImageNumber = rs("bannerSortOrder") + 1
    End If
    
    rs.close()
    Set rs = nothing
    
    GetLastImageNumber = LastImageNumber    
End Function
%>