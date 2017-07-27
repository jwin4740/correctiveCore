<%If Session("hdAdminIsLoggedIn") = "Yes" Then %>
      <table width="200" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/sitesettings/" class="nav" title="Configure you default site setting and meta data values.">Site Settings</a></td>
        </tr>
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/pages/" class="nav" title="Manage Content Pages">Primary Pages</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/pages/" title="List Content Page">Edit</a>
            <%If Session("adminRole") = 1 Then %>
            |&nbsp;<a href="<%=hdAdminPath%>functions/pages/edit.asp?id=0" title="Add New Content Page">Add New</a>
            <%End If %>
          
          </td>
        </tr>
        <%If Session("adminRole") = 1 Then %>
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/categories/" class="nav">Categories</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/categories/" title="List Categories">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/categories/edit.asp?id=0" title="Add New Category">Add New</a>          
          </td>
        </tr>
        <%End If%>
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/banners/" class="nav" title="Manage Banners">Banners</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/banners/" title="Banner List">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/banners/edit.asp?id=0" title="Add New Banners">Add New</a>          
          </td>
        </tr>        
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/news/" class="nav">News Articles</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/news/" title="List Articles">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/news/add.asp" title="Add News Article">Add New</a>
          </td>
        </tr>
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/gallery/" class="nav">Photo Gallery</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/gallery/" title="Photo Gallery">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/gallery/edit.asp?id=0" title="Add New Photos">Add New</a>
          </td>
        </tr>        
        <tr>
          <td class="border"><a href="<%=hdAdminPath%>functions/testimonies/" class="nav">Testimonies</a>
          <br />
            &nbsp;<a href="<%=hdAdminPath%>functions/testimonies/" title="List Testimonies">Edit</a>
            |&nbsp;<a href="<%=hdAdminPath%>functions/testimonies/edit.asp?id=0" title="Add Testimony">Add New</a>
          </td>
        </tr> 
    </table>
    <br /><br />
    <table width="200" border="0" cellspacing="1" cellpadding="0"> 
        <tr>
          <td class="border"><center>COMING SOON</center></td>
        </tr>          
        <tr>
          <td class="border"><a href="#" class="nav">Blog</a></td>
        </tr>
        <tr>
          <td class="border"><a href="#" class="nav">Calendar</a></td>
        </tr>
        <tr>
          <td class="border"><a href="#" class="nav">Contact List</a></td>
        </tr>        
        <tr>
          <td class="border"><a href="#" class="nav">Quick Mail</a></td>
        </tr>        <tr>
          <td class="border"><a href="#" class="nav">Google Analytics</a></td>
        </tr>
    </table>
	  <br /><br />
<%End If %>
	  <table width="200" border="0" cellspacing="1" cellpadding="0">
        <tr>
          <td class="border">CMS By Halstead Design</td>
        </tr>
        <tr>
          <td class="border"><a title="Lake Norman Website Design" href="http://www.halsteaddesign.net" target="_blank" class="nav">www.halsteaddesign.net</a></td>
        </tr>
      </table>
