<div id="divFooter" class="footerArea">

<div class="container">

<div class="divPanel">

<div class="row-fluid">
<div class="span3" id="footerArea1">

<%
        myNews = getRsNews(1, "footer 1")
        
        If IsArray(myNews) Then
            For idx = 0 to UBound(myNews,2)
                If myNews(0,idx) Then
                    hdActionPage = hdPDFDir & myNews(2,idx)
                    hdActionTarget = " target=""_blank"""
                Else
                    hdActionPage = "viewdetails.asp?id=" & myNews(1,idx)
                    hdActionTarget = ""
                End If
                hdActionTitle = "" & myNews(4,idx)        
            %>
<h3><%=myNews(4,idx)%></h3>
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>

</div>
<div class="span3" id="footerArea2">

<%
        myNews = getRsNews(1, "footer 2")
        
        If IsArray(myNews) Then
            For idx = 0 to UBound(myNews,2)
                If myNews(0,idx) Then
                    hdActionPage = hdPDFDir & myNews(2,idx)
                    hdActionTarget = " target=""_blank"""
                Else
                    hdActionPage = "viewdetails.asp?id=" & myNews(1,idx)
                    hdActionTarget = ""
                End If
                hdActionTitle = "" & myNews(4,idx)        
            %>
<h3><%=myNews(4,idx)%></h3>
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>

</div>
<div class="span3" id="footerArea3">

<%
        myNews = getRsNews(1, "footer 3")
        
        If IsArray(myNews) Then
            For idx = 0 to UBound(myNews,2)
                If myNews(0,idx) Then
                    hdActionPage = hdPDFDir & myNews(2,idx)
                    hdActionTarget = " target=""_blank"""
                Else
                    hdActionPage = "viewdetails.asp?id=" & myNews(1,idx)
                    hdActionTarget = ""
                End If
                hdActionTitle = "" & myNews(4,idx)        
            %>
<h3><%=myNews(4,idx)%></h3>
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>

</div>
<div class="span3" id="footerArea4">

<h3>Locations</h3>  
   
<ul id="contact-info">
   <li>
<i class="general foundicon-home icon" style="margin-bottom:50px"></i>
<span class="field">Huntersville:</span>
<br />
16405 Northcross Dr. Suite D<br />
Huntersville, NC<br />
28078
</li>
<li>
<i class="general foundicon-mail icon" style="margin-bottom:50px"></i>
<span class="field">Email:</span>
<br />
<a href="mailto:jj@correctivecore.net" title="Email">jj@correctivecore.net</a>
<br />
(980) 434-6770
</li>
</ul>

   

</div>
</div>

<br /><br />
<div class="row-fluid">
<div class="span12">

<p class="social_bookmarks">
<a href="https://www.facebook.com/correctivecore/" target="_blank"><i class="social foundicon-facebook"></i> </a>
<a href="https://twitter.com/jetzer_julie" target="_blank"><i class="social foundicon-twitter"></i> </a>
</p>


<p class="copyright">
<a href="#">Halstead Design</a> . <a href="#">Website Design</a> . <a href="#">SEO</a>
</p>
</div>
</div>
<br />

</div>

</div>

</div>