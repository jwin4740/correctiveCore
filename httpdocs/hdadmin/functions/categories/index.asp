<!--#include virtual="/includes/global.asp" -->
<%
If Not hdSecureCMSAdminPageSuperUser Then response.redirect(hdAdminPath)

hdAdminPageTitle = "Category Management"

%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case "c"
        hsSQLORDERBY = " ORDER BY hdCategories.catName"
    Case "t"
        hsSQLORDERBY = " ORDER BY hdCategoryType.cattypeName, hdCategories.catName"
    Case Else
        hsSQLORDERBY = " ORDER BY hdCategories.catName"
End Select

SortQueryStringEnd = "&s=" & hdRsSortOrder

'## unique QS's
If Request("selct") = "" Then 
    mySelectedCatTypeID = 0
Else
    mySelectedCatTypeID = CInt(Request("selct"))
    LocalQueryStringEnd = "&selct=" & mySelectedCatTypeID
End If

If mySelectedCatTypeID Then
    hsSQLWHERE = " WHERE hdCategoryType.cattypeID = " & mySelectedCatTypeID
Else
    hsSQLWHERE = ""
End If

QueryStringEnd = SortQueryStringEnd & LocalQueryStringEnd

hsSQL = "SELECT hdCategories.catID, hdCategories.catName, hdCategories.catDescription, hdCategoryType.cattypeID, hdCategoryType.cattypeName " & _
    "FROM hdCategories LEFT JOIN hdCategoryType ON hdCategories.cattypeID = hdCategoryType.cattypeID" & hsSQLWHERE & hsSQLORDERBY
    
    
%>
<!-- #include virtual="/includes/db/getpagingrs.asp" -->
<!-- #include file="../../hdadmindriver.asp" -->
<%
Public Sub PageContent()
%>
    <script type="text/javascript" language="javascript">
        function ddl_change_type() { location.href = "index.asp?selct=" + document.hdEditForm.cattypeID.options[document.hdEditForm.cattypeID.options.selectedIndex].value; }  
    </script>
    <table width="100%" border="0" cellspacing="1" cellpadding="0">
      <tr>
        <td valign="top" class="border">
        <table width="100%" border="0" cellspacing="1" cellpadding="0">
          <form method="post" action="" name="hdEditForm">
          <tr>
            <td class="titleborder">
                <span id="pagetitle">Filter Type:  
                <%Call hdGetCategoryTypesDDL(mySelectedCatTypeID, " onchange=""ddl_change_type();""")%>
                </span>
            </td>
          </tr>
          </form>
        </table>        
<%If hdRsBcontinue Then %>  
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="55%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=c<%=LocalQueryStringEnd%>" title="sort by category">Category</a></td>
            <td width="35%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=t<%=LocalQueryStringEnd%>" title="sort by category Type">Type</a></td>
            <td width="10%"></td>
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
            <%hdLineTitle = "Edit " & hdRS("cattypeName") & " Category &quot;" & hdRs("catName") & "&quot;" %>
            <td width="55%"><a href="edit.asp?id=<%=hdRs("catID")%>" title="<%=hdLineTitle%>"><%=hdRs("catName")%> (<%=hdRs("catDescription")%>)</a></td>
            <td width="35%"><a href="edit.asp?id=<%=hdRs("catID")%>" title="<%=hdLineTitle%>"><%=hdRs("cattypeName")%></a></td>
            <td width="10%"><a href="edit.asp?id=<%=hdRs("catID")%>" title="<%=hdLineTitle%>">Edit</a></td>
          </tr>
          <%
            hdRs.MoveNext
            If hdRs.EOF Then Exit For
          Next        
          %>
          <tr>
            <td colspan="3" align="center"><!-- #include virtual="/includes/db/rspaging.asp" --></td>
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