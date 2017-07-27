<!--#include virtual="/includes/global.asp" -->
<!--#include file="hddriver-contact.asp" -->


<%
Public Sub PageContent()
%>
    <% Response.Write(webpgContent) %>

<div id="form-wrapper"><form action="request.asp" method="post" >

<script type="text/javascript"     src="http://www.google.com/recaptcha/api/challenge?k=6LfWRc4SAAAAAF0m_h1smQbi2WeQ6OcYj5i60Zis">  </script>  <noscript>     <iframe src="http://www.google.com/recaptcha/api/noscript?k=6LfWRc4SAAAAAF0m_h1smQbi2WeQ6OcYj5i60Zis"         height="300" width="500" frameborder="0"></iframe><br>     <textarea name="recaptcha_challenge_field" rows="3" cols="40">     </textarea>     <input type="hidden" name="recaptcha_response_field"         value="manual_challenge">  </noscript>

Please complete required fields <font color="#FF0000">*</font><br /><br />
<input name="Full Name" type="text" id="Full Name"  value="Full Name" onfocus="if(this.value==this.defaultValue)this.value='';" onblur="if(this.value=='')this.value=this.defaultValue;" alt="Full Name"  /><br />
<input name="Phone" type="text" id="Phone" value="Phone Number" onfocus="if(this.value==this.defaultValue)this.value='';" onblur="if(this.value=='')this.value=this.defaultValue;" alt="Phone Number"  /><br />
<input name="Email" type="text" id="Email"  value="Email Address" onfocus="if(this.value==this.defaultValue)this.value='';" onblur="if(this.value=='')this.value=this.defaultValue;" alt="Email Address"  /><br />
<textarea name="Comments"></textarea><br />
<label>
  <input type="submit" name="Send Email" id="Send Email" value="Submit" />
</label>
<br />
</form></div>

<%
End Sub
%>