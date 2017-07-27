<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

hdAdminPageTitle = "User Management"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT hdAdmin.*, hdAdminRole.admrlsName FROM hdAdmin INNER JOIN hdAdminRole ON " & _
        "hdAdmin.adminRole = hdAdminRole.admrlsID WHERE adminRole >= " & _
        Session("adminRole") & " ORDER BY adminRole, adminPersonName, adminUsername"
    hdRs.Open()

%>

<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td valign="top" class="border">
    <table width="100%" border="0" cellspacing="0" cellpadding="5">
      <%
      While Not hdRs.EOF
      %>
      <tr>
        <td width="20"><a href="edit.asp?id=<%=hdRs("adminID")%>" title="Edit <%=hdRs("adminPersonName")%>">
            <img src="<%=hdAdminPath%>images/page.jpg" width="20" height="24" border="0" alt="Edit <%=hdRs("adminUsername")%>" /></a>
        </td>
        <td width="100%"><a href="edit.asp?id=<%=hdRs("adminID")%>" title="Edit <%=hdRs("adminPersonName")%>">
            <%=hdRs("admrlsName")%> :: <%=hdRs("adminPersonName")%><%If hdRs("adminIsEnabled") = False Then%>&nbsp;<span style="color:Red">Disabled</span><%End If%></a>
        </td>
        <td width="29"><a href="edit.asp?id=<%=hdRs("adminID")%>" title="Edit <%=hdRs("adminPersonName")%>">Edit</a></td>
      </tr>
      <%
        hdRs.MoveNext
      Wend
      %>
      <tr>
         <td><a href="edit.asp?id=0" title="Add New User"><img src="<%=hdAdminPath%>images/add.jpg" width="20" height="20" border="0" alt="Add New User" /></a></td><td colspan="2"><a href="edit.asp?id=0" title="Add New User">Add New User</a></td>
      </tr>
    </table>
    </td>
  </tr>
</table>

<%
    hdRs.Close
    Set hdRs = Nothing

End Sub '## PageContent
%>