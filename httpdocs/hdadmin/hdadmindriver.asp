<%Response.CacheControl = "No-cache"%>
<html>
<head>
    <title>HD Website Content Administration System</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="<%=hdAdminPath%>css/styles.asp" rel="stylesheet" type="text/css" />
</head>
<body <%=hdAddBodyOnLoad%>>
    <div id="head">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td width="36" height="36">
                    <img src="<%=hdAdminPath%>images/page/logo.jpg" alt="" />
                </td>
                <td width="443"><span id="loggedinas"><%If Session("hdAdminIsLoggedIn") = "Yes" Then %>
                    Logged in as <%=Session("adminPersonName")%></span>
                </td>
                <td width="546" align="right" style="padding:5px;"><a href="<%=hdAdminPath%>" class="tnav">Dashboard</a> 
				&nbsp;|&nbsp; <a href="<%=hdAdminPath%>functions/users/edit.asp?id=<%=session("adminID")%>" class="tnav">My Settings</a> &nbsp;|&nbsp; <a href="<%=hdAdminPath%>functions/users/" class="tnav">Users</a> &nbsp;|&nbsp; <a href="<%=hdAdminPath%>login.asp?logout=True" class="tnav">Logout</a>
				<%End If%>
				</td>
            </tr>
        </table>
    </div>
    <div id="page"><%=hdAdminPageTitle%>
    <%If Session("hdAdminErrorMsg") <> "" Then %>
        <span id="alert">~ <%=Session("hdAdminErrorMsg")%></span>
    <%
        Session("hdAdminErrorMsg") = ""
    End If 
    %>    
    </div>
    <table width="100%" border="0" cellspacing="8" cellpadding="0">
        <tr>
            <td width="203" valign="top">
                <!--#include file="includes/leftnav.asp" -->
            </td>
            <td width="90%" valign="top">
                <div id="contentarea">
                    <% Call PageContent %>
                </div>
            </td>
<%If Session("hdAdminIsLoggedIn") = "Yes" Then%>
            <td width="267" valign="top">
              
            </td>
<%End If    '## logged in %>
        </tr>
    </table>
</body>
</html>
