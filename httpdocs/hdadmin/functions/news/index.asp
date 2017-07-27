<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtNews)

hdAdminPageTitle = "News"

%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case "d"
        hsSQLORDERBY = " ORDER BY hdNews.newsDate DESC"
    Case "t"
        hsSQLORDERBY = " ORDER BY hdNews.newsTitle"
    Case "c"
        hsSQLORDERBY = " ORDER BY hdCategories.catName"
    Case Else
        hsSQLORDERBY = " ORDER BY newsDate DESC"
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
    hsSQLWHERE = " WHERE hdNews.catID = " & mySelectedCatID
Else
    hsSQLWHERE = ""
End If

QueryStringEnd = SortQueryStringEnd & LocalQueryStringEnd

hsSQL = "SELECT hdNews.*, hdCategories.catName FROM hdNews " & _
    "LEFT JOIN hdCategories ON hdNews.catID = hdCategories.catID " & hsSQLWHERE & hsSQLORDERBY
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
                <span id="pagetitle">Current News Articles in Category 
                <%Call hdGetCategoriesDDL(hdNEWScat, mySelectedCatID, " onchange=""ddl_change_type();""")%>
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
            <td width="75%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=t<%=LocalQueryStringEnd%>" title="sort by article name">Article</a></td>
            <td width="15%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=c<%=LocalQueryStringEnd%>" title="sort by category">Category</a></td>
            <td width="29"></td>
          </tr>
          <%
          For X = 1 to hdRs.PageSize
            If X MOD 2 = 0 Then
                myStyleColor = "#ffffff"
            Else
                myStyleColor = "#F2F2F2"
            End If
            If hdRs("newsIsFile") Then
                hdEditPageName = "editpdf.asp"
            Else
                hdEditPageName = "edit.asp"
            End if
          %>
          <tr bgcolor="<%=myStyleColor%>">
            <td width="20"><a href="<%=hdEditPageName%>?id=<%=hdRs("newsID")%>" title="Edit <%=hdRs("newsTitle")%>"><img src="<%=hdAdminPath%>images/page.jpg" width="20" height="24" border="0" alt="Edit <%=hdRs("newsTitle")%>" /></a></td>
            <td width="10%"><a href="<%=hdEditPageName%>?id=<%=hdRs("newsID")%>" title="Edit <%=hdRs("newsTitle")%>"><%=FormatDateTime(hdRs("newsDate"),2)%></a></td>
            <td width="75%"><a href="<%=hdEditPageName%>?id=<%=hdRs("newsID")%>" title="Edit <%=hdRs("newsTitle")%>"><%=hdRs("newsTitle")%> ~ <%=Left(hdRs("newsShortDesc"),75)%><%If Len(hdRs("newsShortDesc")) > 75 Then%>...<%End If%></a></td>
            <td width="15%"><a href="<%=hdEditPageName%>?id=<%=hdRs("newsID")%>" title="Edit <%=hdRs("newsTitle")%>"><%=hdRs("catName")%></a></td>
            <td width="29"><a href="<%=hdEditPageName%>?id=<%=hdRs("newsID")%>" title="Edit <%=hdRs("newsTitle")%>">Edit</a></td>
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