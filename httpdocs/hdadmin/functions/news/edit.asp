<!--#include virtual="/includes/global.asp" -->
<!--#include virtual="/includes/fckeditor/fckeditor.asp" -->

<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

hdAdminPageTitle = "News"

hdAddBodyOnLoad = "onload=""initEditPage();"""

newsID = CInt(Request.QueryString("id"))

%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

newsTitle = ""
newsDate = FormatDateTime(Now(), 2)
newsShortDesc = ""
newsIsFile = False
newsDetails = ""
newsIcon = ""
newsMetaKeywords = ""
newsMetaDescription = ""
newsIsActive = True
hdTitleBannerText = "Adding News."
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

    hdTitleBannerText = "Editing: " & newsTitle
    
End If  '## newsID > 0

%>

<script type="text/javascript" language="javascript">
  
    function hdOpen(sID) 
    {
        for (i=1;i<4;i++)
        {
            if(i == sID)
            {
                var sdisp = "block";
                var sbgcolor = "#F1F1F1";
                var sftwgt = "bold";
            }
        else
            {
                var sdisp = "none";
                var sbgcolor = "white";
                var sftwgt = "";    
            }
        document.getElementById('hdEditData'+i).style.display = sdisp;
        document.getElementById('hdEditTab'+i).style.backgroundColor = sbgcolor;
        document.getElementById('hdEditTab'+i).style.fontWeight = sftwgt;
        }
    }
    
    function initEditPage() 
    {
        for (i=1;i<4;i++)
        {
            document.getElementById('hdEditData'+i).style.display = "none";
        }
        hdOpen(1);
    }

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
    <form method="post" action="process.asp" id="hdEditFormID" name="hdEditForm">
    <input type="hidden" name="newsID" id="newsIDID" value="<%=newsID%>" />
    <input type="hidden" name="newsIsFile" id="newsIsFileID" value="<%=newsIsFile%>" />
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td id="hdEditTab1" nowrap class="border" style="width: 150px; text-align:center; cursor:hand;" onclick="hdOpen(1);">Details</td>
        <td id="hdEditTab2" nowrap class="border" style="width: 150px; text-align:center; cursor:hand;" onclick="hdOpen(2);">Content</td>
        <td id="hdEditTab3" nowrap class="border" style="width: 150px; text-align:center; cursor:hand;" onclick="hdOpen(3);">SEO</td>
      </tr>
    </table>
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
            </table>
        </td>
      </tr>
      <tr>
        <td id="hdEditData2" class="border" width="95%" valign="top" style="display: none;">
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                    <td colspan="2">
 		                <%
                        Dim oFCKeditor
                        Set oFCKeditor = New FCKeditor
                        oFCKeditor.BasePath = "/includes/fckeditor/"
                        oFCKeditor.Height = "400"
                        oFCKeditor.Width = "95%"
                        oFCKeditor.Value = newsDetails
                        oFCKeditor.Create "newsDetails"
		                %> 
                    </td>
                </tr>                                
            </table>
        </td>
      </tr>
      <tr>
        <td id="hdEditData3" class="border" width="95%" valign="top" style="display: none;">
           <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td>Unique Page Keywords: </td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2"><input type="text" size="70" name="newsMetaKeywords" id="newsMetaKeywordsID" value="<%=newsMetaKeywords%>" maxlength="255" class="style" />                      </td>
                </tr>
                <tr>
                  <td>Unique Page Description: </td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2">
                        <textarea name="newsMetaDescription" id="newsMetaDescriptionID" cols="55" rows="7" class="style"><%=newsMetaDescription%></textarea>
                    </td>
                </tr>
            </table>        
        </td>
      </tr>
      <tr>
        <td id="hdEditDataButtons" class="border" width="95%" valign="top" style="display: block;">
           <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td colspan="2">
                    <img src="<%=hdAdminPath%>images/save.jpg" onclick="javascript:document.hdEditForm.submit();" width="126" height="33" hspace="1" vspace="1" alt="Save" style="cursor:hand;" />
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