<!--#include virtual="/includes/global.asp" -->
<!--#include virtual="/includes/fckeditor/fckeditor.asp" -->

<%

hdAddBodyOnLoad = "onload=""initEditPage();"""

webpgID = CInt(Request.QueryString("id"))

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtPages)

If webpgID = 0 And Session("adminRole") <> 1 Then
    Session("hdAdminErrorMsg") = "Not authorized to add pages."
    Response.Redirect("index.asp")
End If

hdAdminPageTitle = "Edit Webpage"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

webpgName = ""
webpgTitle = ""
webpgFileName = ""
webpgMetaKeywords = ""
webpgMetaDescription = ""
webpgContent = ""
webpgIsActive = True
hdTitleBannerText = "Adding New Page."

If webpgID > 0 Then    

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdWebPage WHERE webpgID = " & webpgID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../pages/")
    
    webpgName = hdRs("webpgName")
    webpgTitle = hdRs("webpgTitle")
    webpgFileName = hdRs("webpgFileName")
    webpgMetaKeywords = hdRs("webpgMetaKeywords")
    webpgMetaDescription = hdRs("webpgMetaDescription")
    webpgContent = hdRs("webpgContent")
    webpgIsActive = hdRs("webpgIsActive")    
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "You are currently editing: " & webpgName
    
End If  '## webpgID > 0

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

    function DeletePage(theID,thePg)
    {
        if (confirm("Delete this Page?  This will also delete the file: "+thePg))
        {
	        location.href("delete.asp?id="+theID+"&pg="+thePg);
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" action="process.asp" id="hdEditFormID" name="hdEditForm">
    <input type="hidden" name="webpgID" id="webpgIDID" value="<%=webpgID%>" />
    <input type="hidden" name="OldWebpgFileName" id="OldWebpgFileName" value="<%=webpgFileName%>" />
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
                        Page Name:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="webpgName" id="webpgNameID" value="<%=webpgName%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Title:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="webpgTitle" id="webpgTitleID" value="<%=webpgTitle%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        File Name:
                    </td>
                </tr>
                <tr>
                    <td><input type="text" size="70" name="webpgFileName" id="webpgFileNameID" value="<%=webpgFileName%>" maxlength="255" class="style" <%If Session("adminRole") <> 1 Then%>disabled="disabled"<%End If%> /></td>
                </tr>                               
                <tr>
                    <td>
                        <%If Session("adminRole") = 1 And webpgFileName <> "" Then %>
                        File Status:&nbsp;<%=hdGetFileStatus(webpgFileName,webpgID)%>
                        <%End If %>
                    </td>
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
                        oFCKeditor.Width = "99%"
                        oFCKeditor.Value = webpgContent
                        oFCKeditor.Create "webpgContent"
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
                    <td colspan="2"><input type="text" size="70" name="webpgMetaKeywords" id="webpgMetaKeywords" value="<%=webpgMetaKeywords%>" maxlength="255" class="style" />                      </td>
                </tr>
                <tr>
                  <td>Unique Page Description: </td>
                  <td>&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="2">
                        <textarea name="webpgMetaDescription" id="webpgMetaDescriptionID" cols="55" rows="7" class="style"><%=webpgMetaDescription%></textarea>
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
                    <%If Session("adminRole") = 1 And webpgID > 1 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeletePage(<%=webpgID%>,'<%=webpgFileName%>')" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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