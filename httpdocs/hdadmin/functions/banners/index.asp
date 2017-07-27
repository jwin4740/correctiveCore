<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

If Session("adminRole") <> 1 Then hdCheckFeaturePermission(setMgtBanners)

hdAdminPageTitle = "Banners"

'## testing
'## hdRsDefaultPageSize = 2

%>
<!-- #include virtual="/includes/db/getpagingqs.asp" -->
<%

'## check this data unique selections

'## unique sort
Select Case hdRsSortOrder
    Case Else
        hsSQLORDERBY = " ORDER BY bannerSortOrder"
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

hsSQL = "SELECT * FROM hdBanners " & hsSQLWHERE & hsSQLORDERBY
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
                <span id="pagetitle">Banners Listed in Sort Order</span>
            </td>
          </tr>
        </table>        
<%If hdRsBcontinue Then %>
        <table width="100%" border="0" cellspacing="0" cellpadding="5">
          <tr>
            <td width="45%">Banner</td>
            <td width="35%">Name</td>
            <td width="15%">Sort Order</td>
            <td width="5%">&nbsp;</td>
          </tr>
          <%
          For X = 1 to hdRs.PageSize
            If X MOD 2 = 0 Then
                myStyleColor = "#ffffff"
            Else
                myStyleColor = "#F2F2F2"
            End If

            hdEditPageHREF = "edit.asp?id=" & hdRs("bannerID")
            hdEditPageTitle = "Edit " & hdRs("bannerName")
                    
            bannerID = hdRs("bannerID")
            bannerName = hdRs("bannerName")
            bannerDescription = hdRs("bannerDescription")
            bannerSortOrder = hdRs("bannerSortOrder")
            bannerImage = hdRs("bannerImage")
            bannerImageAltTag = hdRs("bannerImageAltTag")
            
            hdRs.MoveNext
            
            PrevSort = bannerSortOrder - 1
            NextSort = bannerSortOrder + 1
            
            If hdRs.BOF Then PrevSort = 0
            
            If hdRs.EOF Then NextSort = 0            

          %>
          <tr bgcolor="<%=myStyleColor%>">
            <td width="45%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><img src="<%=hdBannerDir%><%=bannerImage%>" width="<%=(hdBannerW/2)%>" height="<%=(hdBannerH/2)%>" border="0" alt="<%=bannerImageAltTag%>" /></a></td>
            <td width="35%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>"><%=bannerName%></a></td>
            <td width="15%" align="center">
                <%If PrevSort > 0 Then %><a href="move.asp?ID=<%=bannerID%>&toswap=<%=PrevSort%>&thissort=<%=bannerSortOrder%>"><img border="0" alt="move up" title="move up" src="<%=hdAdminPath%>images/up.gif" /></a><%Else %>&nbsp;<%End If%>
                <%=bannerSortOrder%>
                <%If NextSort > 0 Then %><a href="move.asp?ID=<%=bannerID%>&toswap=<%=NextSort%>&thissort=<%=bannerSortOrder%>"><img border="0" alt="move down" title="move down" src="<%=hdAdminPath%>images/down.gif" /></a><%Else %>&nbsp;<%End If%>
            </td>
            <td width="5%"><a href="<%=hdEditPageHREF%>" title="<%=hdEditPageTitle%>">Edit</a></td>
          </tr>
          <%
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