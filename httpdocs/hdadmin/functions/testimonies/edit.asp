<!--#include virtual="/includes/global.asp" -->

<%

Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtTestimony)

testiID = CInt(Request.QueryString("id"))

hdThisRecType = "Testimony"

hdAdminPageTitle = "Manage " & hdThisRecType
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

testiDate = ""
testiTitle = ""
testiAuthor = ""
testiQuote = ""
testiIsActive = True
hdTitleBannerText = "Adding " & hdThisRecType
catID = 0

If testiID > 0 Then    

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdTestimony WHERE testiID = " & testiID
    hdRs.Open()
    
    If hdRs.EOF Then Response.Redirect("../testimonies/")
    
    testiDate = hdRs("testiDate")
    testiTitle = hdRs("testiTitle")
    testiAuthor = hdRs("testiAuthor")
    testiQuote = hdRs("testiQuote")
    testiIsActive = hdRs("testiIsActive")
    catID = hdRs("catID")
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "Editing " & hdThisRecType & ": " & testiTitle
    
End If  '## testiID > 0

%>

<script type="text/javascript" language="javascript">
    function DeletePage(theID)
    {
        if (confirm("Delete this <%=hdThisRecType%>?"))
        {
	        location.href("delete.asp?id="+theID);
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" name="hdEditForm" action="process.asp" >
    <input type="hidden" name="testiID" id="testiIDID" value="<%=testiID%>" />
    <input type="hidden" name="testiIsActive" id="testiIsActiveID" value="<%=testiIsActive%>" />
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
                        Category - choose from available categories for your <%=hdThisRecType%>:</td>
                </tr>
                <tr>
                  <td><%Call hdGetCategoriesDDL(hdTESTIMONYcat, catID, "")%></td>
                </tr>
    
                <tr>
                    <td>
                        <%=hdThisRecType%> Title - as if it where a search term:</td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="testiTitle" id="testiTitleID" value="<%=testiTitle%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Author - who gave the Testimony:
                    </td>
                </tr>
                <tr>
                  <td><input type="text" size="70" name="testiAuthor" id="testiAuthorID" value="<%=testiAuthor%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        The Quote - Enter the Testimony here:
                    </td>
                </tr>
                <tr>
                  <td>
                  <textarea name="testiQuote" rows="9" cols="55" id="testiQuoteID" class="style"><%=testiQuote%></textarea>
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
                    <%If testiID > 0 Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeletePage(<%=testiID%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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