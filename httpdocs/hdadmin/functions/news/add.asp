<!--#include virtual="/includes/global.asp" -->

<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

hdAdminPageTitle = "Add News Article"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()
%>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <form method="post" action="x.asp" id="hdEditFormID" name="hdEditForm">
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td class="titleborder"><span id="pagetitle">Step 1, select news article type.</span></td>
      </tr>
    </table>    
    <table border="0" cellspacing="5" cellpadding="5">
      <tr><td>&nbsp;</td></tr>
      <tr>
        <td>
        <p><input style="width:175px;cursor:hand;" type="button" name="goArticle" id="goArticleID" value="News Article" onclick="javascript:location.href='edit.asp?id=0'" /></p>
        <p><input style="width:175px;cursor:hand;" type="button" name="goPDF" id="goPDFID" value="Link to PDF" onclick="javascript:location.href='editpdf.asp?id=0'" /></p>
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