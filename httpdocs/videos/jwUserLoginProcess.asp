<!--#include virtual="/includes/global.asp" -->
<%

Session("jwUserLoginErrorMsg") = ""
Session("jwUserLoginIsLoggedIn") = ""

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

    jwUsername = Request.Form("jwUsername")
    jwPassword = Request.Form("jwPassword")
  
    '## if no problems with the username and password then insert new user into database
    If Session("jwUserLoginErrorMsg") = "" Then

        Set hdRs = Server.CreateObject("ADODB.Recordset")
     
        hdRs.ActiveConnection = hdDSN ' hdDSN must be the database credentials
          
        hdRs.Source = "SELECT * FROM jwUsers WHERE userUsername = '" & jwUsername & "' AND userPassword = '" & jwPassword & "'"
        hdRs.Open

        If hdRs.EOF Then
            Session("jwUserLoginIsLoggedIn") = ""
            Session("jwUserLoginErrorMsg") = Session("jwUserLoginErrorMsg") & "Login failed: Username/Password incorrect "
               Response.Redirect("login.asp") 
        Else
            '## set all the details for the logged in user
            Session("jwUserLoginIsLoggedIn") = "Yes"
               
            Session("userFirstName") = hdRs("userFirstName")
            Session("userUsername") = hdRs("userUsername")
            Session("userEmail") = hdRs("userEmail")
            Session("userID") = hdRs("userID")
            jwProgram = hdRs("userProgram")
            If jwProgram = "prenatal" Then
                Session("jwUserProgram") = jwProgram
                Response.Redirect("Prenatal-Videos.asp") 
            ElseIf jwProgram = "foundation" Then
                Session("jwUserProgram") = jwProgram
                Response.Redirect("Foundation-Videos.asp")
            Else 
                Response.Redirect("login.asp") 
            End If
        End If
       
        hdRs.Close
        Set hdRs = Nothing
    
    End If
    Response.Redirect("login.asp")    
End If



%>

