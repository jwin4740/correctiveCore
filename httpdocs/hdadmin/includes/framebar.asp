<%@  enablesessionstate="False" language="VBScript" %>
<% Response.Expires = -1 %>
<!-- #include virtual="/includes/variables.asp" -->
<html>
<head>
    <title>Uploading files</title>
    <style type='text/css'>
        td
        {
            font-family: arial;
            font-size: 9pt;
        }
    </style>
</head>
<% If Request("b") = "IE" Then %>
<!-- Internet Explorer -->
<body bgcolor="#C0C0B0">
    <iframe src="<%=hdAdminPath%>includes/bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>"
        title="Upload Progress" noresize scrolling="no" frameborder="0" framespacing="10"
        width="369" height="65"></iframe>
    <table border="0" width="100%" cellpadding="2" cellspacing="0">
        <tr>
            <td align="center">
                To cancel uploading, press your browser's <b>STOP</b> button.
            </td>
        </tr>
    </table>
</body>
<%Else%>
<!-- Netscape Navigator etc ... -->
<frameset rows="65%, 35%" cols="100%" border="0" framespacing="0" frameborder="NO">
<frame SRC="<%=hdAdminPath%>includes/bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" noresize scrolling="NO" frameborder="NO" name="sp_body">
</frameset>
<%End If%>
</html>
