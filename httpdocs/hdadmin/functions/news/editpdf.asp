<!--#include virtual="/includes/global.asp" -->

<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

newsID = CInt(Request.QueryString("id"))

hdAdminPageTitle = "Manage &quot;PDF Link&quot; News Article"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

newsTitle = ""
newsDate = FormatDateTime(Now(), 2)
newsShortDesc = ""
newsIsFile = True
newsDetails = ""
newsIcon = ""
newsMetaKeywords = ""
newsMetaDescription = ""
newsIsActive = True
hdTitleBannerText = "Adding &quot;PDF Link&quot; News."
catID = 0

If newsID > 0 Then    

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdNews WHERE newsID = " & newsID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../news/")
    
    newsTitle = hdRs("newsTitle")
    newsDate = hdRs("newsDate")
    newsShortDesc = hdRs("newsShortDesc")
    newsIsFile = hdRs("newsIsFile")
    newsDetails = hdRs("newsDetails")
    newsIcon = hdRs("newsIcon")
    newsMetaKeywords = hdRs("newsMetaKeywords")
    newsMetaDescription = hdRs("newsMetaDescription")
    newsIsActive = hdRs("newsIsActive")
    catID = hdRs("catID")
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "Editing &quot;PDF Link&quot;: " & newsTitle
    
End If  '## newsID > 0

'## better be a file to edit on this type of news article
If Not newsIsFile Then Response.Redirect("../news/")

'## used in the aspuploadjs
hdThisFormName = "hdEditForm"
hdThisFileFieldName = "newsDetails"

%>
<!-- #include file="../../includes/aspuploadinit.asp" -->
<!-- #include file="../../includes/aspuploadjs.asp" -->
<script type="text/javascript" language="javascript">
    function DeletePage(theID)
    {
        if (confirm("Delete this Article?"))
        {
	        location.href("delete.asp?id="+theID);
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" enctype="multipart/form-data" name="hdEditForm" action="processpdf.asp?<%=PID%>" >
    <input type="hidden" name="newsID" id="newsIDID" value="<%=newsID%>" />
    <input type="hidden" name="newsIsFile" id="newsIsFileID" value="<%=newsIsFile%>" />
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
                        Category - choose which category you want your article to display:</td>
                </tr>
                <tr>
                  <td><%Call hdGetCategoriesDDL(hdNEWScat, catID, "")%></td>
                </tr>
                <tr>
                    <td>
                        Article Date:</td>
                </tr>    
                <tr>
                  <td><input type="text" size="35" name="newsDate" id="newsDateID" value="<%=newsDate%>" maxlength="35" class="style" /></td>
                </tr>    
                <tr>
                    <td>
                        Title Of Your Article - Title the article as if it where a search term:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="newsTitle" id="newsTitleID" value="<%=newsTitle%>" maxlength="180" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Short Description - paste a few lines of your article here - up to 180 characters:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="newsShortDesc" id="newsShortDescID" value="<%=newsShortDesc%>" maxlength="180" class="style" /></td>
                </tr>
                <%If newsID > 0 Then %>
                <tr>
                    <td>
                        Link to Existing PDF News Article File:
                    </td>
                </tr>
                <tr>
                  <td><a href="<%=hdPDFDir%><%=newsDetails%>" target="_blank" title="<%=newsTitle%>"><%=newsDetails%></a></td>
                </tr>
                <%End If %>                
                <tr>
                    <td>
                        Upload New PDF File:
                    </td>
                </tr>
                <tr>
                  <td><input type="file" size="55" name="newsDetails" id="newsDetailsID" value="<%=newsDetails%>" class="style" /></td>
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
                    <%If newsID > 0 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeletePage(<%=newsID%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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