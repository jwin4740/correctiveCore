<!--#include virtual="/includes/global.asp" -->
<%
Call hdSecureCMSAdminPage

hdAdminPageTitle = "HD Admin Overview ~ Dashboard"
%>

<!-- #include file="hdadmindriver.asp" -->

<%
Public Sub PageContent()
%>

<table width="678" border="0" cellspacing="0" cellpadding="2">
    <tr>
        <td width="27">
            <img src="<%=hdAdminPath%>images/page_2.gif" width="23" height="29">
        </td>
        <td width="631">
            <span id="title">Primary Pages (total:<%=hdGetFeatureRecordCount("hdWebPage","webpgID")%>)</span><br>
            Manage your main web pages within your primary directory
            <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/pages/" title="List Content Page">Edit</a>
            <%If Session("adminRole") = 1 Then %>
            |&nbsp;<a href="<%=hdAdminPath%>functions/pages/edit.asp?id=0" title="Add New Content Page">Add New</a>
            <%End If%>
        </td>
    </tr>
    <%If Session("adminRole") = 1 Then %>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/news.gif" width="23" height="24">
        </td>
        <td>
            <span id="title">Categories (total:<%=hdGetFeatureRecordCount("hdCategories","catID")%>)</span><br>
            Manage CMS Categories.<br />
            &nbsp;<a href="<%=hdAdminPath%>functions/categories/" title="List News Articles">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/categories/edit.asp?id=0" title="Add News Article">Add Category</a>
        </td>
    </tr>
    <%End If%>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/banner.jpg" width="23" height="21">
        </td>
        <td>
            <span id="title">Banner Images (total:<%=hdGetFeatureRecordCount("hdBanners","bannerID")%>)</span><br>
            You can use banner as ads or page decoration.<br />
            &nbsp;<a href="<%=hdAdminPath%>functions/banners/" title="Banner List">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/banners/edit.asp?id=0" title="Add New Banner">Add Bbanner</a>            
        </td>
    </tr>    
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/news.gif" width="23" height="24">
        </td>
        <td>
            <span id="title">News Articles (total:<%=hdGetFeatureRecordCount("hdNews","newsID")%>)</span><br>
            Streaming news and new posts are a good way to keep your site active.<br />
            &nbsp;<a href="<%=hdAdminPath%>functions/news/" title="List News Articles">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/news/edit.asp?id=0" title="Add News Article">Add News Article</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/news/editpdf.asp?id=0" title="Add PDF Page">Add &quot;Link to PDF&quot; Article</a>
        </td>
    </tr>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/gallery.jpg" width="23" height="21">
        </td>
        <td>
            <span id="title">Photo gallery (total:<%=hdGetFeatureRecordCount("hdGallery","galID")%>)</span><br>
            Manage your images, share them with the world or keep them to yourself.<br />
            &nbsp;<a href="<%=hdAdminPath%>functions/gallery/" title="Photo Gallery">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/gallery/edit.asp?id=0" title="Add Photos">Add Photos</a>

        </td>
    </tr>    
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/testis.jpg" width="27" height="29">
        </td>
        <td>
            <span id="title">Testimonials (total:<%=hdGetFeatureRecordCount("hdTestimony","testiID")%>)</span><br>
            Add personal touches with customer testimonies.<br />
            &nbsp;<a href="<%=hdAdminPath%>functions/testimonies/" title="List Testimonies">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/testimonies/edit.asp?id=0" title="Add Testimony">Add Testimony</a>
            
        </td>
    </tr>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td height="15" colspan="2"><br /><span id="title">Coming Soon ...</span></td>
    </tr>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>    
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/tweet.gif" width="27" height="29">
        </td>
        <td>
            <span id="title">Blog</span><br>
            Post data about you or your company daily
        </td>
    </tr>    
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/calendar.gif" width="23" height="21">
        </td>
        <td>
            <span id="title">Calendar Admin</span><br>
            Your calendar can be used on the backend or displayed live on your website.
        </td>
    </tr>
    <tr>
        <td height="15" colspan="2" background="<%=hdAdminPath%>images/gline2.jpg"><img src="<%=hdAdminPath%>images/gline2.jpg" width="1" height="12"></td>
    </tr>
    <tr>
        <td>
            <img src="<%=hdAdminPath%>images/mail.gif" width="23" height="18">
        </td>
        <td>
            <span id="title">Quick Mail</span><br>
            Send an html formated email to clients
        </td>
    </tr>
</table>

<%
End Sub '## PageContent
%>