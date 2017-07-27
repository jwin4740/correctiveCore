<!--#include virtual="/includes/global.asp" -->
<%
 Response.Redirect("login.asp") 
' Session("jwUserLoginErrorMsg") = ""
Session("jwUserRegistrationErrorMsg") = ""

If (Request.ServerVariables("CONTENT_LENGTH") <> 0) Then

DIM programCapture
programCapture = Request.Form("programChoice")

If programCapture = "" Then
Session("jwUserRegistrationErrorMsg") = "Please select a program to enroll in"
Response.Redirect("register83018421293AHam.asp") 
End If


' userToken = Request.Form("jwRegistrationToken")
' Set hdRs = Server.CreateObject("ADODB.Recordset")
     
' hdRs.ActiveConnection = hdDSN ' hdDSN must be the database credentials
          
' hdRs.Source = "SELECT * FROM jwTokens WHERE Token = '" & userToken & "' AND Used = No"
' hdRs.Open

If hdRs.EOF Then
    Session("jwUserRegistrationErrorMsg") = Session("jwUserRegistrationErrorMsg") & "account creation failed: invalid token"
    Response.Redirect("register83018421293AHam.asp") 
End If
       
hdRs.Close
Set hdRs = Nothing

If InStr(userToken, "p") <> 0 Then
    programCapture = "prenatal"
ElseIf InStr(userToken, "f") <> 0 Then
    programCapture = "foundation"
Else 
Session("jwUserRegistrationErrorMsg") = Session("jwUserRegistrationErrorMsg") & "account creation failed: didn't grab suffix"
Response.Redirect("register83018421293AHam.asp") 

End If


jwRegistrationUsername = Request.Form("jwRegistrationUsername")
jwRegistrationPassword = Request.Form("jwRegistrationPassword")
jwRegistrationConfirmPassword = Request.Form("jwRegistrationConfirmPassword")
jwFirstName = Request.Form("jwFirstName")
jwLastName = Request.Form("jwLastName")
jwDOB = Request.Form("jwDOB")
jwAddress = Request.Form("jwAddress")
jwCity = Request.Form("jwCity")
jwState = Request.Form("jwState")
jwZip = Request.Form("jwZip")
jwEmail = Request.Form("jwFirstEmail")
jwCellNumber = Request.Form("jwCellNumber")


' If jwRegistrationPassword <> jwRegistrationConfirmPassword Then
'   Session("jwUserRegistrationErrorMsg") = Session("jwUserRegistrationErrorMsg") & "account creation failed: passwords do not match. please try again"
'   Response.Redirect("register83018421293AHam.asp") 
' End If 

' If len(jwRegistrationPassword) < 6 Then
'   Session("jwUserRegistrationErrorMsg") = Session("jwUserRegistrationErrorMsg") & "account creation failed: password must be longer than 5 characters"
'   Response.Redirect("register83018421293AHam.asp") 
' End If 

'     '## if no problems user registration fields insert into database
      
'     If Session("jwUserRegistrationErrorMsg") = "" Then
 
'         Set jwConn = Server.CreateObject("ADODB.Connection")
'         jwConn.Open hdDSN
'         qSQL = "INSERT INTO jwUsers (userUsername, userPassword, userProgram, userFirstName, userLastName, userDOB, userAddress, userCity, userState, userZip, userEmail, userCell) 
'         VALUES ('" & jwRegistrationUsername & "', '" & jwRegistrationPassword & "', '" & programCapture & "', '" & jwFirstName & "', '" & jwLastName & "', '" & jwDOB & "', '" & jwAddress & "', '" & jwCity & "', '" & jwState & "', '" & jwZip & "', '" & jwEmail & "', '" & jwCellNumber & "')"
'         ' uSQL = "UPDATE jwTokens SET Used = Yes WHERE Token = '" & userToken & "'"
'         jwConn.Execute(qSQL)
'         ' jwConn.Execute(uSQL)
'         Response.Redirect("login.asp")  
'         jwConn.Close
'         Set jwConn = Nothing

'     End If  
  
End If



%>

