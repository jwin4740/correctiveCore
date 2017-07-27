<%@ EnableSessionState=False LANGUAGE=VBScript %>
<!-- #include virtual="/includes/variables.asp" -->
<%
  Response.Expires = -1
  PID = Request("PID")
  TimeO = Request("to")

  Set UploadProgress = Server.CreateObject("Persits.UploadProgress")

  format = "%TUploading files...%t%B3%T%R left (at %S/sec) %r%U/%V(%P)%l%t"

  bar_content = UploadProgress.FormatProgress(PID, TimeO, "#00007F", format)

  If "" = bar_content Then
%>
<html>
<head>
    <title>Upload Finished</title>

<script language="JavaScript">
function CloseMe()
{
	window.parent.close();
	return true;
}
</script>

</head>
<body onload="CloseMe()">
</body>
</html>
<% Else %>
<html>
<head>

    <meta http-equiv="Refresh" content="1;URL=<%=Request.ServerVariables("URL") & "?to=" & TimeO & "&PID=" & PID %>" />

    <title>Uploading Files...</title>
    <style type='text/css'>
        td
        {
            font-family: arial;
            font-size: 9pt;
        }
        td.spread
        {
            font-size: 6pt;
            line-height: 6pt;
        }
        td.brick
        {
            font-size: 6pt;
            height: 12px;
        }
    </style>
</head>
<body bgcolor="#C0C0B0" topmargin="0">
    <%= bar_content %>
</body>
</html>
<% End If %>
