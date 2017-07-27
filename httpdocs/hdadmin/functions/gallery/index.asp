<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtGallery)

hdAdminPageTitle = "Photo Gallery"

'## testing
'## hdRsDefaultPageSize = 2
%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case "id"
        hsSQLORDERBY = " ORDER BY hdGallery.galID"
    Case "i"
        hsSQLORDERBY = " ORDER BY hdGallery.galID DESC"
    Case "n"
        hsSQLORDERBY = " ORDER BY hdGallery.galName"
    Case "c"
        hsSQLORDERBY = " ORDER BY hdGallery.catID"
    Case Else
        hsSQLORDERBY = " ORDER BY hdGallery.galID DESC"
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
    hsSQLWHERE = " WHERE hdGallery.catID = " & mySelectedCatID
Else
    hsSQLWHERE = ""
End If

QueryStringEnd = SortQueryStringEnd & LocalQueryStringEnd

hsSQL = "SELECT hdGallery.*, hdCategories.catName FROM hdGallery " & _
    "LEFT JOIN hdCategories ON hdGallery.catID = hdCategories.catID " & hsSQLWHERE & hsSQLORDERBY
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
                <span id="pagetitle">Current Gallery Images in Category 
                <%Call hdGetCategoriesDDL(hdGALLERYcat, mySelectedCatID, " onchange=""ddl_change_type();""")%>
                </span>
            </td>
          </tr>
        </form>
        </table>         
<%If hdRsBcontinue Then %>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="5%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=id<%=LocalQueryStringEnd%>" title="sort by ID first to last">ID</a></td>
            <td width="30%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=i<%=LocalQueryStringEnd%>" title="sort by Image newest to oldest">Image</a></td>
            <td width="45%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=n<%=LocalQueryStringEnd%>" title="sort by Name">Name</a></td>
            <td width="15%"><a class="nav" href="<%=Request.ServerVariables("SCRIPT_NAME")%>?s=c<%=LocalQueryStringEnd%>" title="sort by category">Category</a></td>
            <td width="5%">&nbsp;</td>
          </tr>
          <%
          For X = 1 to hdRs.PageSize
            If X MOD 2 = 0 Then
                myStyleColor = "#ffffff"
            Else
                myStyleColor = "#F2F2F2"
            End If

            hdEditPageHREF = "edit.asp?id=" & hdRs("galID")
            hdEditPageTitle = "Edit " & hdRs("galName")
                    
            galID = hdRs("galID")
            galName = hdRs("galName")
            galThumbFile = hdRs("galThumbFile")
            galDescription = hdRs("galDescription")
            catName = hdRs("catName")
            
          %>
          <tr bgcolor="<%=myStyleColor%>">
            <td width="5%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=galID%></a></td>
            <td width="30%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><img src="<%=hdGalleryDir%><%=galThumbFile%>" width="<%=hdGalleryThumbW%>" height="<%=hdGalleryThumbH%>" border="0" alt="<%=galDescription%>" /></a></td>
            <td width="45%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=galName%></a></td>
            <td width="15%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=catName%></a></td>
            <td width="5%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>">Edit</a></td>
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