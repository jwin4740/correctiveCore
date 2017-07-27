<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

hdAdminPageTitle = "Website Default Settings"
%>

<!-- #include file="../../hdadmindriver.asp" -->

<%
Public Sub PageContent()

    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdSiteSettings WHERE setID = 1"
    hdRs.Open()
%>
<form method="post" action="sitesettingsprocess.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="5">
    <tr class="border">
        <td>
            Website URL http://
        </td>
        <td>
            <input type="text" size="50" name="setSiteURL" id="setSiteURLID" value="<%=hdRs("setSiteURL")%>" maxlength="100" class="style" />
        </td>
    </tr>
    <tr class="border">
        <td>
            Website Title:
        </td>
        <td>
            <input type="text" size="50" name="setDefaultTitle" id="setDefaultTitleID" value="<%=hdRs("setDefaultTitle")%>" maxlength="255" class="style" />
        </td>
    </tr>    
    <tr class="border">
        <td valign="top">
            META Keywords
        </td>
        <td>
            <textarea name="setDefaultMetaKeywords" id="setDefaultMetaKeywordsID" cols="45" rows="4" class="style"><%=hdRs("setDefaultMetaKeywords")%></textarea>
        </td>
    </tr>
    <tr class="border">
        <td valign="top">
            META Description
        </td>
        <td>
            <textarea name="setDefaultMetaDescription" id="setDefaultMetaDescriptionID" cols="45" rows="4" class="style"><%=hdRS("setDefaultMetaDescription")%></textarea>
        </td>
    </tr>
    <tr class="border">
        <td colspan="2">
            CMS Version: <b><%=hdRs("setCMSVersion")%></b>
        </td>
    </tr>
<%
    If Session("adminRole") = 1 Then    '## superuser settings
%>  
    <tr class="border">
        <td valign="top">
            Is Site Active:
        </td>
        <td>
            <input type="checkbox" name="setIsActive" id="setIsActiveID" value="True"
            <%If hdRs("setIsActive") Then Response.Write("checked")%>
             />            
        </td>
    </tr>
    <tr class="border">
        <td colspan="2">
            Site Administrator Functions.  Allowed to ...
        </td>
    </tr>

    <tr class="border">
        <td valign="top">
            Manage Pages:
        </td>
        <td>
            <input type="checkbox" name="setMgtPages" id="setMgtPagesID" value="True"
            <%If hdRs("setMgtPages") Then Response.Write("checked")%>
             />            
        </td>
    </tr>
    <tr class="border">
        <td valign="top">
            Banners:
        </td>
        <td>
            <input type="checkbox" name="setMgtBanners" id="setMgtBannersID" value="True"
            <%If hdRs("setMgtBanners") Then Response.Write("checked")%>
             />            
        </td>
    </tr> 
    <tr class="border">
        <td valign="top">
            Blogs:
        </td>
        <td>
            <input type="checkbox" name="setMgtBlog" id="setMgtBlogID" value="True"
            <%If hdRs("setMgtBlog") Then Response.Write("checked")%>
             />            
        </td>
    </tr>    
    <tr class="border">
        <td valign="top">
            Calendar:
        </td>
        <td>
            <input type="checkbox" name="setMgtCalendar" id="setMgtCalendarID" value="True"
            <%If hdRs("setMgtCalendar") Then Response.Write("checked")%>
             />            
        </td>
    </tr>
    <tr class="border">
        <td valign="top">
            Contacts:
        </td>
        <td>
            <input type="checkbox" name="setMgtContact" id="setMgtContactID" value="True"
            <%If hdRs("setMgtContact") Then Response.Write("checked")%>
             />            
        </td>
    </tr>  
    <tr class="border">
        <td valign="top">
            NEWS:
        </td>
        <td>
            <input type="checkbox" name="setMgtNews" id="setMgtNewsID" value="True"
            <%If hdRs("setMgtNews") Then Response.Write("checked")%>
             />            
        </td>
    </tr> 
    <tr class="border">
        <td valign="top">
            Photo Gallery:
        </td>
        <td>
            <input type="checkbox" name="setMgtGallery" id="setMgtGalleryID" value="True"
            <%If hdRs("setMgtGallery") Then Response.Write("checked")%>
             />            
        </td>
    </tr>  
 
    <tr class="border">
        <td valign="top">
            Projects:
        </td>
        <td>
            <input type="checkbox" name="setMgtProjects" id="Checkbox2" value="True"
            <%If hdRs("setMgtProjects") Then Response.Write("checked")%>
             />            
        </td>
    </tr> 
            
    <tr class="border">
        <td valign="top">
            Quick Mail:
        </td>
        <td>
            <input type="checkbox" name="setMgtMailer" id="setMgtMailerID" value="True"
            <%If hdRs("setMgtMailer") Then Response.Write("checked")%>
             />            
        </td>
    </tr> 
    <tr class="border">
        <td valign="top">
            Testimonies:
        </td>
        <td>
            <input type="checkbox" name="setMgtTestimony" id="setMgtTestimonyID" value="True"
            <%If hdRs("setMgtTestimony") Then Response.Write("checked")%>
             />            
        </td>
    </tr> 
<%
    End If  '## superuser
%>  
    <tr>
        <td>
            &nbsp;
        </td>
        <td>
            <input type="submit" name="submit" value="Update" />
        </td>
    </tr>
</table>
</form>
<%
End Sub '## PageContent
%>