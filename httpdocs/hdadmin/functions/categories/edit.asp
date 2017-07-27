<!--#include virtual="/includes/global.asp" -->

<%

If Not hdSecureCMSAdminPageSuperUser Then response.redirect(hdAdminPath)

'#############
'## NEED TO VERIFY IF USER CAN MANAGE NEWS.

catID = CInt(Request.QueryString("id"))

hdAdminPageTitle = "Category"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

cattypeID = 0
catName = ""
catImage = ""
catDescription = ""
catIsActive = True
cattypeDBTableName = ""
catWarningRecords = 0
hdTitleBannerText = "Adding Category"

If catID > 0 Then    

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT hdCategories.*, hdCategoryType.cattypeDBTableName FROM hdCategories LEFT JOIN hdCategoryType ON " & _
        "hdCategories.cattypeID = hdCategoryType.cattypeID WHERE hdCategories.catID = " & catID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../categories/")
    
    cattypeID = hdRs("cattypeID")
    catName = hdRs("catName")
    catImage = hdRs("catImage")
    catDescription = hdRs("catDescription")
    catIsActive = hdRs("catIsActive")
    
    cattypeDBTableName = hdRs("cattypeDBTableName")
    
    hdRs.Close
    Set hdRs = Nothing
    
    If cattypeDBTableName <> "" Then

        Set hdRs = Server.CreateObject("ADODB.Recordset")
        hdRs.ActiveConnection = hdDSN
        hdRs.Source = "SELECT Count(catID) AS CountOfcatID FROM " & cattypeDBTableName & " WHERE catID = " & catID
        hdRs.Open()
        catWarningRecords = hdRs("CountOfcatID")
        hdRs.Close
        Set hdRs = Nothing        
    End If

    hdTitleBannerText = "Editing: " & catName
    
End If  '## catID > 0

%>
<script type="text/javascript" language="javascript">
    function DeletePage(theID)
    {
        if (confirm("Delete this Category?  Remember, if any records reference this category \n you'll break the relationship."))
        {
	        location.href("delete.asp?id="+theID);
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" action="process.asp" id="hdEditFormID" name="hdEditForm">
    <input type="hidden" name="catID" id="catIDID" value="<%=catID%>" />
    <input type="hidden" name="catImage" id="catImageID" value="<%=catImage%>" />
    <input type="hidden" name="catIsActive" id="catIsActiveID" value="<%=catIsActive%>" />
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
                        Category Type: 
                    </td>
                </tr>
                <tr>
                  <td>
                    <%Call hdGetCategoryTypesDDL(cattypeID, "")%>
                    <%If catWarningRecords > 0 Then %>
                        <%=catWarningRecords%> items assigned to this category.
                    <%End If%>
                  </td>
                </tr>
    
                <tr>
                    <td>
                        Category Name:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="catName" id="catNameID" value="<%=catName%>" maxlength="50" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Short Description:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="catDescription" id="catDescriptionID" value="<%=catDescription%>" maxlength="180" class="style" /></td>
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
                    <%If catID > 0 AND catWarningRecords = 0 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeletePage(<%=catID%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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