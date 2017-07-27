<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtPages)

hdAdminPageTitle = "Primary Pages Management - Current Pages"
%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case "n"
        hsSQLORDERBY = " ORDER BY webpgName"
    Case "t"
        hsSQLORDERBY = " ORDER BY webpgTitle"
    Case "f"
        hsSQLORDERBY = " ORDER BY webpgFileName"        
    Case Else
        hsSQLORDERBY = " ORDER BY webpgIsHomePage, webpgIsActive DESC, webpgName"
End Select

SortQueryStringEnd = "&s=" & hdRsSortOrder

'## unique QS's

QueryStringEnd = SortQueryStringEnd & LocalQueryStringEnd

hsSQL = "SELECT * FROM hdWebPage " & hsSQLWHERE & hsSQLORDERBY
    
%>
<!-- #include virtual="/includes/db/getpagingrs.asp" -->
<!-- #include file="../../hdadmindriver.asp" -->
<%
Public Sub PageContent()
%>

<table width="100%" border="0" cellspacing="1" cellpadding="0">
  <tr>
    <td valign="top" class="border">
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td class="titleborder">
            <span id="pagetitle">Current Pages in Primary Directory</span>
        </td>
      </tr>
    </table> 
<%If hdRsBcontinue Then %>  
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="5%"></td>
            <td width="55%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=n<%=LocalQueryStringEnd%>" title="sort by page name">Name</a></td>
            <td width="55%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=f<%=LocalQueryStringEnd%>" title="sort by file name">File</a></td>
            <td width="5%"></td>
          </tr>
          <%
          For X = 1 to hdRs.PageSize
            If X MOD 2 = 0 Then
                myStyleColor = "#ffffff"
            Else
                myStyleColor = "#F2F2F2"
            End If
          %>
          <tr bgcolor="<%=myStyleColor%>">
            <%hdLineTitle = "Edit " & hdRs("webpgTitle") %>
            <td width="5%"><a href="edit.asp?id=<%=hdRs("webpgID")%>" title="<%=hdLineTitle%>"><img src="<%=hdAdminPath%>images/page.jpg" width="20" height="24" border="0" alt="<%=hdLineTitle%>" /></a></td>
            <td width="75%"><a href="edit.asp?id=<%=hdRs("webpgID")%>" title="<%=hdLineTitle%>"><%=hdRs("webpgName")%></a></td>
            <td width="15%"><a href="edit.asp?id=<%=hdRs("webpgID")%>" title="<%=hdLineTitle%>"><%=hdRs("webpgFileName")%></a></td>
            <td width="5%"><a href="edit.asp?id=<%=hdRs("webpgID")%>" title="<%=hdLineTitle%>">Edit</a></td>
          </tr>
          <%
            hdRs.MoveNext
            If hdRs.EOF Then Exit For
          Next        
          %>
          <tr>
            <td colspan="4" align="center"><!-- #include virtual="/includes/db/rspaging.asp" --></td>
          </tr>
        </table>
<%End If  '## hdRsBcontinue %>        
    </td>
  </tr>
</table>

<!-- #include virtual="/includes/db/rsclose.asp" -->  

<%
End Sub '## PageContent
%>