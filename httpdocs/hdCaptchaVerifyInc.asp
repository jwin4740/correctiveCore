<%
    'Check Captcha.
    bCaptcha = True
    if Request.Form("myCaptcha") = "" then
        Response.Write("Please complete the Captcha request. <br />")
        bCaptcha = False
    else
        If IsNumeric(Request.Form("myCaptcha")) Then
            If CInt(Request.Form("myCaptcha")) = CInt(Session("hdCaptchaValue") + Session("hdCaptchaAdd")) Then
                '## correct captcha
            Else
                 Response.Write("Invalid Captcha entry.  Please try again. <br />")
                 bCaptcha = False
            End If
        Else
            Response.Write("Invalid Captcha entry.  Please add the numbers and try again. <br />")
            bCaptcha = False
        End If
    end if

    If Not bCaptcha Then
        Response.Write("Use your back button in your browser to complete the form")
        Response.End 
    End If
 %>