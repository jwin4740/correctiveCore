
<%
Randomize
iMaxLimit = 19
Session("hdCaptchaValue") = Int((iMaxLimit * Rnd) + 1) 
If Session("hdCaptchaAdd") = 1 Then
    Session("hdCaptchaAdd") = 0 
Else 
    Session("hdCaptchaAdd") = 1
End If
%>

            Answer: <%=Session("hdCaptchaValue")%> + <%=Session("hdCaptchaAdd")%>
            <br />
            <input type="text" class="input-block-level" name="myCaptcha" value="" maxlength="2" placeholder="Answer" />

