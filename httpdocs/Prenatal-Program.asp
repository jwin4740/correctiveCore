<!--#include virtual="/includes/global.asp" -->
<!--#include file="hddriver.asp" -->
<%
Public Sub PageContent()
%>
    <% Response.Write(webpgContent) %>

<h2>Program Options</h2>
<ul class="nav nav-tabs">
<li class="active"><a href="#Schedule-Assessment" data-toggle="tab">Schedule Assessment</a></li>
<li><a href="#Prenatal-Full-Support-Program" data-toggle="tab">Prenatal Full Support Program</a></li>
<li><a href="#Prenatal-Self-Starter-with-Support" data-toggle="tab">Prenatal Self Starter with Support</a></li>
<li><a href="#Prenatal-Self-Starter-On-line-Program" data-toggle="tab">Prenatal Self Starter On-line Program</a></li>
<li><a href="#Prenatal-Workshops" data-toggle="tab">Prenatal Workshops</a></li>
</ul>
<div class="tab-content">
<div class="tab-pane fade in active" id="Schedule-Assessment">
<%
        myNews = getRsNews(1, "Schedule Assessment")
        
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
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>
</div>
<div class="tab-pane fade" id="Prenatal-Full-Support-Program">
<%
        myNews = getRsNews(1, "Prenatal Full Support Program")
        
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
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>
</div>
<div class="tab-pane fade" id="Prenatal-Self-Starter-with-Support">
<%
        myNews = getRsNews(1, "Prenatal Self Starter with Support")
        
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
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>
</div>
<div class="tab-pane fade" id="Prenatal-Self-Starter-On-line-Program">
<%
        myNews = getRsNews(1, "Prenatal Self Starter On-line Program")
        
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
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>
</div>
<div class="tab-pane fade" id="Prenatal-Workshops">
<%
        myNews = getRsNews(1, "Prenatal Workshops")
        
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
<%=myNews(2,idx)%>
<%  
            Next
        End If
        %>
</div>
</div>
<hr>
<p style="text-align: center;"><img src="/assets/prenatalcorrectivecore-img.png" width="850" height="300" alt="" /></p>
<hr>
<p><span style="font-size: x-small;">YOU SHOULD CONSULT YOUR PHYSICIAN OR OTHER HEALTH CARE PRACTITIONER BEFORE STARTING THIS OR ANY OTHER EXERCISE PROGRAM. THIS IS PARTICULARLY TRUE IF YOU OR YOUR FAMILY HAVE A HISTORY OF HIGH BLOOD PRESSURE OR HEART DISEASE, OR IF YOU HAVE EVER EXPERIENCED DISCOMFORT WHILE EXERCISING. NOTHING STATED OR POSTED ON CORRECTIVE CORE &amp; MUSCULOSKELETAL HEALTH, LLC ALONG WITH THRIVEing, LLC SERVICES ARE NOT INTENDED TO BE, AND MUST NOT BE TAKEN TO BE, THE PRACTICE OF MEDICAL OR PROFESSIONAL ADVICE OR CARE.</span></p>
<p><span style="font-size: x-small;">YOUR USE OF THE CORRECTIVE CORE &amp; MUSCULOSKELETAL HEALTH, LLC ALONG WITH THRIVEing, LLC SERVICE ARE AT YOUR OWN RISK. CORRECTIVE CORE &amp; MUSCULOSKELETAL HEALTH, LLC ALONG WITH THRIVEing, LLC OR THEIR AFFILIATES SHALL NOT BE LIABLE FOR ANY LIABILITY, OF ANY KIND, RESULTING FROM THE USE OF THE CORRECTIVE CORE &amp; MUSCULOSKELETAL HEALTH, LLC ALONG WITH THRIVEing, LLC SERVICES IN OFFICE OR OUTSIDE OF SET OFFICE APPOINTMENTS.</span></p>
    

<%    
End Sub
%>
