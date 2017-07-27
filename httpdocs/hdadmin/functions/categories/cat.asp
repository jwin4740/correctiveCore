<!--#include virtual="/includes/global.asp" -->
<%
If Not hdSecureCMSAdminPageSuperUser Then response.redirect(hdAdminPath)

hdAdminPageTitle = "Categories"

hsSQL = "SELECT hdNews.*, hdCategories.catName FROM hdNews LEFT JOIN hdCategories ON hdNews.catID = hdCategories.catID ORDER BY newsDate DESC"
%>
<!-- #include file="../../includes/getpagingrs.asp" -->
<!-- #include file="../../hdadmindriver.asp" -->
<%
Public Sub PageContent()

    If hdRsBcontinue Then
    %>

    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td valign="top" class="border">
        <table width="100%" border="0" cellspacing="1" cellpadding="0">
          <tr>
            <td class="titleborder">
                <span id="pagetitle">Current News Articles in Category 
                <%Call hdGetCategoriesDDL(hdNEWScat, catID)%>
                </span>
            </td>
          </tr>
        </table>        
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="20">&nbsp;</td>
            <td width="10%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?cp=<%=hdRsCurrentPage%>&s=a" title="sort by Date">Date</a></td>
            <td width="75%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?cp=<%=hdRsCurrentPage%>&s=a" title="sort by article name">Article</a></td>
            <td width="15%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?cp=<%=hdRsCurrentPage%>&s=c" title="sort by category">Category</a></td>
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
            <td colspan="5" align="center"><!-- #include file="../../includes/rspaging.asp" --></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>

    <%
    End If  '## hdRsBcontinue
%>
<!-- #include file="../../includes/rsclose.asp" -->  
<%
End Sub '## PageContent
%>