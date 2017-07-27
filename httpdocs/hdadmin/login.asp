<!--#include file="../includes/global.asp" -->
<%
hdAdminPageTitle = "Halstead Design CMS Login"

hdAddBodyOnLoad = "onload=""document.forms.hdloginf.hdusername.focus();"""

'## quick logout
If Request.QueryString("logout") = "True" Then
    Session.Abandon
    Session("hdAdminIsLoggedIn") = ""
    Session("hdAdminErrorMsg") = "Logout completed."
End If
%>

<!-- #include file="hdadmindriver.asp" -->

<%
Public Sub PageContent()
%>

<form method="post" action="hdadminloginprocess.asp" name="hdloginf" id="hdloginfID">
<table width="600" border="0" cellspacing="0" cellpadding="5">
    <tr>
      <td width="77">&nbsp;</td>
      <td width="503">Username</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td><input name="hdusername" type="text" id="hdusernameid" size="20" maxlength="30" class="style" /></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Password</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td><input name="hdpassword" type="password" id="hdpasswordid" size="20" maxlength="30" class="style" /></td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td><input type="submit" name="hdsubmit" id="hdsubmitid" value="Login" /></td>
    </tr>
</table>
</form>

<%
End Sub '## PageContent
%>