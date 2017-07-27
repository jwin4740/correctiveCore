<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtTestimony)

hdAdminPageTitle = "Testimonies"

%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case "d"
        hsSQLORDERBY = " ORDER BY testiDate DESC"
    Case "t"
        hsSQLORDERBY = " ORDER BY testiTitle"
    Case "a"
        hsSQLORDERBY = " ORDER BY testiAuthor"        
    Case Else
        hsSQLORDERBY = " ORDER BY testiDate DESC"
End Select

SortQueryStringEnd = "&s=" & hdRsSortOrder

'## unique QS's
If Request("selcat") = "" Then 
    mySelectedCatID = 0
Else
    mySelectedCatID = CInt(Request("selcat"))
    LocalQueryStringEnd = "&selcat=" & mySelectedCatID
End If

If mySelectedCatID Then
    hsSQLWHERE = " WHERE catID = " & mySelectedCatID
Else
    hsSQLWHERE = ""
End If

QueryStringEnd = SortQueryStringEnd & LocalQueryStringEnd

hsSQL = "SELECT * FROM hdTestimony " & hsSQLWHERE & hsSQLORDERBY
%>
<!-- #include virtual="/includes/db/getpagingrs.asp" -->
<!-- #include file="../../hdadmindriver.asp" -->
<%
Public Sub PageContent()
%>
    <script type="text/javascript" language="javascript">
        function ddl_change_type() { location.href = "index.asp?selcat=" + document.hdEditForm.catID.options[document.hdEditForm.catID.options.selectedIndex].value; }  
    </script>
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td valign="top" class="border">
        <table width="100%" border="0" cellspacing="1" cellpadding="0">
        <form method="post" action="" name="hdEditForm">
          <tr>
            <td class="titleborder">
                <span id="pagetitle">Current Testimonies for Category 
                <%Call hdGetCategoriesDDL(hdTESTIMONYcat, mySelectedCatID, " onchange=""ddl_change_type();""")%>
                </span>
            </td>
          </tr>
        </form>
        </table>        
<%If hdRsBcontinue Then %>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="20">&nbsp;</td>
            <td width="10%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=d<%=LocalQueryStringEnd%>" title="sort by Date">Date</a></td>
            <td width="75%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=t<%=LocalQueryStringEnd%>" title="sort by Title">Title</a></td>
            <td width="15%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=a<%=LocalQueryStringEnd%>" title="sort by Author">Author</a></td>
            <td width="29"></td>
          </tr>
          <%
          For X = 1 to hdRs.PageSize
            If X MOD 2 = 0 Then
                myStyleColor = "#ffffff"
            Else
                myStyleColor = "#F2F2F2"
            End If

            hdEditPageHREF = "edit.asp?id=" & hdRs("testiID")
            hdEditPageTitle = "Edit " & hdRs("testiTitle")

          %>
          <tr bgcolor="<%=myStyleColor%>">
            <td width="20"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><img src="<%=hdAdminPath%>images/page.jpg" width="20" height="24" border="0" alt="<%=hdEditPageTitle%>" /></a></td>
            <td width="10%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=FormatDateTime(hdRs("testiDate"),2)%></a></td>
            <td width="75%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=hdRs("testiTitle")%></a></td>
            <td width="15%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=hdRs("testiAuthor")%></a></td>
            <td width="29"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>">Edit</a></td>
          </tr>
          <%
            hdRs.MoveNext
            If hdRs.EOF Then Exit For
          Next        
          %>
          <tr>
            <td colspan="5" align="center"><!-- #include virtual="/includes/db/rspaging.asp" --></td>
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