<!--#include virtual="/includes/global.asp" -->
<%

adminID = CInt(Request.QueryString("id"))

Call hdSecureCMSAdminPage

hdAdminPageTitle = "Manage Users"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

adminPersonName = ""
adminEmail = ""
adminUsername = ""
adminPassword = ""
adminRole = 2
adminIsEnabled = True
hdTitleBannerText = "Adding New Page."
showPassword = ""

If adminID > 0 Then    
    
    hdSQL = "SELECT * FROM hdAdmin WHERE adminRole >= " & Session("adminRole") & " AND adminID = " & adminID
    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = hdSQL
    hdRs.Open()
    
    If hdRs.EOF Then 
        '## send an email about missing or denied user access.
        hdSubject = "User Edit Security Alert"
        hdBody = Session("adminPersonName") & " attempted to issue SQL: " & hdSQL & vbCrLf & vbCrLf & _
            "adminUsername: " & Session("adminUsername") & vbCrLf & _
            "adminRole: " & Session("adminRole")
            
        Call hdSendEmail(CStr("cms@" & Session("setSiteURL")), hdMasterUsersEmailAddr, hdSubject, hdBody, False)
        Response.Redirect("index.asp")
    End If
    
    adminPersonName = hdRs("adminPersonName")
    adminEmail = hdRs("adminEmail")
    adminUsername = hdRs("adminUsername")
    adminPassword = hdRs("adminPassword")
    adminRole = hdRs("adminRole")
    adminIsEnabled = hdRs("adminIsEnabled")
    
    hdRs.Close
    Set hdRs = Nothing

    hdTitleBannerText = "Editing User: " & adminPersonName
    
    If Request.Querystring("sp") = "show" Then showPassword = adminPassword
    
End If  '## adminID > 0

Set hdrsRoles = Server.CreateObject("ADODB.Recordset")
hdrsRoles.ActiveConnection = hdDSN
hdrsRoles.Source = "SELECT * FROM hdAdminRole WHERE admrlsID >= " & Session("adminRole") & " ORDER BY admrlsID"
hdrsRoles.Open()

%>

<script type="text/javascript" language="javascript">
    function DeleteUser(theID)
    {
        if (confirm("Delete this User?"))
        {
	        location.href="delete.asp?id="+theID;
        }
    }
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="top">
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td class="titleborder"><span id="pagetitle"><%=hdTitleBannerText%></span></td>
      </tr>
    </table>    
    <form method="post" action="process.asp" id="hdEditUserID" name="hdEditUser">
    <input type="hidden" name="adminID" id="adminIDID" value="<%=adminID%>" />
    <input type="hidden" name="oldadminUsername" id="oldadminUsernameID" value="<%=adminUsername%>" />
    <input type="hidden" name="oldadminEmail" id="oldadminEmailID" value="<%=adminEmail%>" />
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td id="hdEditData1" class="border" width="95%" valign="top">
            <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                    <td>
                        Name:</td>
                  <td><input type="text" size="45" name="adminPersonName" id="adminPersonNameID" value="<%=adminPersonName%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Username:</td>
                  <td>
                        <input type="text" size="45" name="adminUsername" id="adminUsernameID" value="<%=adminUsername%>" maxlength="50" class="style" /><br />
                        * 5 to 20 characters
                  </td>
                </tr>
                <tr>
                    <td>
                        Password:</td>
                  <td>
                    <input type="text" size="45" name="adminPassword" id="adminPasswordID" value="<%=showPassword%>" maxlength="20" class="style" /><br />
                    * 5 to 20 characters<br />
                    <%If adminID <> 0 Then %>* Only enter a password to change it.&nbsp;
                        <%If showPassword = "" Then%>
                            <a href="edit.asp?id=<%=adminID%>&sp=show">Show Password</a>
                        <%Else %>
                            <a href="edit.asp?id=<%=adminID%>">Hide Password</a>
                        <%End If%>
                    <%End If%>
                  </td>
                </tr>                                
                <tr>
                    <td>
                        Email:
                    </td>
                  <td><input type="text" size="45" name="adminEmail" id="adminEmailID" value="<%=adminEmail%>" maxlength="255" class="style" /></td>
                </tr>
                <tr>
                    <td>
                        Role:
                    </td>
                  <td>
                    <select name="adminRole" id="adminRoleID" size="1">
                    <%
                    While Not hdrsRoles.EOF
                    %>
                        <option value="<%=hdrsRoles("admrlsID")%>" <%If adminRole = hdrsRoles("admrlsID") Then Response.Write("selected")%>><%=hdrsRoles("admrlsName")%> (<%=hdrsRoles("admrlsDesc")%>)</option>
                    <%
                        hdrsRoles.MoveNext
                    Wend
                    %>
                    </select>
                  </td>
                </tr>
                <tr>
                    <td>
                        Enabled:
                    </td>
                    <td><input type="radio" name="adminIsEnabled" value="True" <%If adminIsEnabled Then response.write("checked")%> />Yes&nbsp;&nbsp;<input type="radio" name="adminIsEnabled" value="False" <%If Not adminIsEnabled Then response.write("checked")%> />No <%If adminIsEnabled = False Then%>(<span style="color:Red">Disabled</span>)<%End If%></td>
                </tr>                               
            </table>
        </td>
      </tr>
       <tr>
        <td class="border" width="95%" valign="top">
           <table width="100%" border="0" cellspacing="0" cellpadding="5">
                <tr>
                  <td colspan="2">
                    <img src="<%=hdAdminPath%>images/save.jpg" onclick="javascript:document.hdEditUser.submit();" width="126" height="33" hspace="1" vspace="1" alt="Save" style="cursor:hand;" />
                    <img src="<%=hdAdminPath%>images/cancel.jpg" onclick="javascript:location.href='index.asp'" width="126" height="33" hspace="1" vspace="1" alt="Cancel" style="cursor:hand;" />
                    <%If adminID > 0 and adminID <> Session("adminID") Then %>
                    <img src="<%=hdAdminPath%>images/delete.jpg" onclick="javascript:DeleteUser(<%=adminID%>)" width="126" height="33" hspace="1" vspace="1" alt="Delete" style="cursor:hand;" />
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