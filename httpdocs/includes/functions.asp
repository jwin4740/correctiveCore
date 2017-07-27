<%

'## COMMON FUNCTIONS SECTION

'###################################
function hdIsValidEmailAddress(addr)

	dim list, item
	dim i, c

	hdIsValidEmailAddress = true

	'Exclude any address with '..'.
	if InStr(addr, "..") > 0 then
		hdIsValidEmailAddress = false
		exit function
	end if

	'Split email address into the user and domain names.
	list = Split(addr, "@")
	if UBound(list) <> 1 then
		hdIsValidEmailAddress = false
		exit function
	end if

	'Check both names.
	for each item in list

		'Make sure the name is not zero length.
		if Len(item) <=  0 then
			hdIsValidEmailAddress = false
			exit function
		end if

		'Make sure only valid characters appear in the name.
		for i = 1 to Len(item)
			c = Lcase(Mid(item, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz&_-.", c) <= 0 and not IsNumeric(c) then
				hdIsValidEmailAddress = false
				exit function
			end if
		next

		'Make sure the name does not start or end with invalid characters.
		if Left(item, 1) = "." or Right(item, 1) = "." then
			hdIsValidEmailAddress = false
			exit function
		end if

	next

	'Check for a '.' character in the domain name.
	if InStr(list(1), ".") <= 0 then
		hdIsValidEmailAddress = false
		exit function
	end if

end function


'############################################################
Function hdSendEmail(fromAddr, toAddr, subject, body, isHTML)

	Set Mail = Server.CreateObject("Persits.MailSender")
	
		Mail.Host = "127.0.0.1" ' Specify your websites IP Address
		Mail.From = fromAddr ' Specify sender's address

		'## provide multiple TO addresses comma delimited.
		recipients = Split(toAddr, ",")
		for each name in recipients
			Mail.AddAddress name
		next

		Mail.AddReplyTo fromAddr
		Mail.Subject = subject
		Mail.IsHTML = isHTML
		Mail.Body = body

		On Error Resume Next

		Mail.Send 'this actually sends the email

		If Err <> 0 Then
			hdSendEmail = False
		Else
			hdSendEmail = True
		End If
	
	Set Mail = Nothing

End Function


'###########################################################################
'## ValidationType	= what type of validation we're doing
'## FieldToCheck	= what string we're checking
'## MinLen			= if greater then 0 we'll check the string length.
Function hdFieldValidate(ValidationType, FieldToCheck, MinLen, SpecialCheck)

	If FieldToCheck = "" Then
		allValid = False	
	Else
		Select Case ValidationType
			Case 1	'## USA zip or phone. Some phone validation use ()-. chars
				checkOK = "0123456789-"
				
			Case 2	'## phone
				checkOK = "0123456789-().+"
			
			Case 3	'## alpha
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
				
			Case 4	'## alphanumeric
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"

			Case 5	'## some special characters
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_ !&()*$#@{}[]:;'\<>?"

			Case 6	'## credit card characters
				checkOK = "0123456789- "

			Case 7	'## string
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_"

			Case 8	'## mysql password
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!$*"

			Case 9	'## valid admin password characters
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789!@#$%^&*()-_{}[]|~"

			Case 10	'## valid name characters
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 -'"

			Case 11	'## valid meta keyword/desc characters
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 -,_+$."

			Case 12	'## valid domain name characters
				checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-."
																																						
			Case Else
				checkOK = SpecialCheck
				
		End Select

		allValid = True
		For i = 1 To Len(FieldToCheck)
			ch = Mid(FieldToCheck, i, 1)
			For j = 1 To Len(checkOK)
				if ch = Mid(checkOK, j, 1) Then Exit For
			Next
			If j > Len(checkOK) Then
				allValid = False
				Exit For
			End If
		Next
		
		If MinLen > 0 Then
			If Len(FieldToCheck) < MinLen Then allValid = False
		End If
			
	End If
	
	hdFieldValidate = allValid
	
End Function

Function GetApplicationPath()
        GetApplicationPath = Mid(Request.ServerVariables("APPL_MD_PATH"), Len(Request.ServerVariables("INSTANCE_META_PATH")) + 6) & "/"
End Function

'#########################
Sub hdSecureCMSAdminPage()

    If Session("hdAdminIsLoggedIn") <> "Yes" Then Response.Redirect(hdAdminPath & "login.asp")

End Sub '## hdSecureCMSAdminPage()


'#######################################
Function hdSecureCMSAdminPageSuperUser()
    
    bContinue = False
    
    Call hdSecureCMSAdminPage()
    
    If Session("adminRole") = 1 Then
        bContinue = True
    Else
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Insufficient privileges to execute requested function.<br />"
    End If 
    
    hdSecureCMSAdminPageSuperUser = bContinue 
    
End Function '## hdSecureCMSAdminPageSuperUser()


'#######################
Sub LoadHDSiteSettings()
    
    On Error Resume Next

    arFields = Array("setSiteURL", "setDefaultTitle", "setDefaultMetaKeywords", "setDefaultMetaDescription", "setIsActive")

    '## hit the database up for the default site settings.
    
    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.ActiveConnection = hdDSN
    hdRs.Source = "SELECT * FROM hdSiteSettings WHERE setID = 1"
    hdRs.Open()

    If hdRs.EOF Then
        Response.Write("HD Error 91: error reading site settings")
        Response.End
    End If
    
    '## loop thru each var from the db and set the session vars.
    For Each qField in arFields
        Session(qField) = hdRs(qField)
    Next    
    
    hdRs.Close()
    Set hdRs = Nothing

End Sub '## LoadHDSiteSettings()


'#######################################
Function hdCreateDriverPage(newFileName)
    
    'On Error Resume Next
    
    '## return codes
    '## 1 = success
    '## 2 = template file not found
    '## 3 = new file already exists
    
    hdCreateDriverPage = 1
 
    'Map the file name to the physical path on the server.

    pathToNewFile = Server.MapPath("/") & "/" & newFileName
    pathToTemplate = Server.MapPath("/") & hdAdminPath & hdPathToDriverTemplate
  
    set fs = CreateObject("Scripting.FileSystemObject")

    '## make sure the template file is there
    If fs.FileExists(pathToTemplate) then
        '## make sure the new file isn't there
        If fs.FileExists(pathToNewFile) Then
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "File " & newFileName & " already exists.<br />"
            hdCreateDriverPage = 3
        Else
            set inputfile = fs.OpenTextFile(pathToTemplate, ForReading)
            set outputfile = fs.CreateTextFile(pathToNewFile, ForWriting)
            
            do while not inputfile.AtEndOfStream
                outputfile.WriteLine(inputfile.ReadLine)
            loop

            outputfile.Close()
            inputfile.Close()
        End If
    Else
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Template file not found.  Please contact your sysadmin.<br />"
        hdCreateDriverPage = 2
    End If

End Function '## hdCreateDriverPage()


'##############################################
Function hdGetFileStatus(hdFileToCheck,hdTheID)
    
    'On Error Resume Next
    hdMyReturn = ""
    
    hdPathToFile = Server.MapPath("/") & "/" & hdFileToCheck
    
    set fs = CreateObject("Scripting.FileSystemObject")

    '## make sure the template file is there
    If fs.FileExists(hdPathToFile) then
        '## make sure the new file isn't there
        hdMyReturn = "<span style=""color:green;"">File Exists.</span>&nbsp;&nbsp;&nbsp;<a href=""/" & hdFileToCheck & """ target=""_blank"">" & _
            "Click here to Preview Page.</a>"
    Else
        hdMyReturn = "<span style=""color:red;"">File Missing!</span>&nbsp;<a href=""createdriverpage.asp?id=" & webpgID & "&pg=" & hdFileToCheck & """>" & _
            "Click here to Create File.</a>"
    End If

    hdGetFileStatus = hdMyReturn
    
End Function '## hdGetFileStatus()


'##########################################
Function hdDeleteDriverPage(hdTheDriverFile)
    
    hdPathToFile = Server.MapPath("/") & "/" & hdTheDriverFile

    hdDeleteDriverPage = hdDeleteFile(hdPathToFile)
    
End Function '## hdDeleteDriverPage()


'##########################################
Function hdDeleteFile(hdFileToDelete)
    
    'On Error Resume Next
    bContinue = False
      
    set fs = CreateObject("Scripting.FileSystemObject")

    '## make sure the template file is there
    If fs.FileExists(hdFileToDelete) then
        fs.DeleteFile(hdFileToDelete)
        If Not fs.FileExists(hdFileToDelete) then 
            bContinue = True
        End If
    Else
        bContinue = True
    End If

    hdDeleteFile = bContinue
    
End Function '## hdDeleteFile()


'########################################
Function hdCheckIfAdminEmailExists(theEmailAddr)
    
    On Error Resume Next
   
    bContinue = True
    
    Set hdRsEmCk = Server.CreateObject("ADODB.Recordset")
    hdRsEmCk.ActiveConnection = hdDSN
    hdRsEmCk.Source = "SELECT * FROM hdAdmin WHERE adminEmail = '" & theEmailAddr & "'"
    hdRsEmCk.Open()

    If hdRsEmCk.EOF Then bContinue = False
   
    hdRsEmCk.Close()
    Set hdRsEmCk = Nothing
    
    hdCheckIfAdminEmailExists = bContinue

End Function '## hdCheckIfAdminEmailExists()


'########################################
Function hdCheckIfAdminUsernameExists(theUserName)
    
    On Error Resume Next
   
    bContinue = True
    
    Set hdRsEmCk = Server.CreateObject("ADODB.Recordset")
    hdRsEmCk.ActiveConnection = hdDSN
    hdRsEmCk.Source = "SELECT * FROM hdAdmin WHERE adminUsername = '" & theUserName & "'"
    hdRsEmCk.Open()

    If hdRsEmCk.EOF Then bContinue = False
   
    hdRsEmCk.Close()
    Set hdRsEmCk = Nothing
    
    hdCheckIfAdminUsernameExists = bContinue

End Function '## hdCheckIfAdminUsernameExists()


'############################################
Function hdGenerateAdminUsername(theUserName)
    
    On Error Resume Next
   
    leftUsername = Trim(Left(theUsername, 18))
    
    '## if there are this many users with the same name, then bail. 
    '## we got a problem
    For tLoop = 1 to 99
        tempUsername = leftUsername & CStr(tLoop)
        If Not hdCheckIfAdminUsernameExists(tempUsername) Then Exit For
    Next
    
    If tLoop >= 99 Then
        Repsonse.Write("Username Error.  Please contact your System Administrator.")
        Response.End
    End If
    
    hdGenerateAdminUsername = tempUsername

End Function '## hdGenerateAdminUsername()


'####################################################
Function hdVerifyAdminUsernameUnique(hdAdminUsername)
    
    On Error Resume Next

    If hdFieldValidate(7, hdAdminUsername, 5, "") Then
        If hdCheckIfAdminUsernameExists(hdAdminUsername) Then
            hdAdminUsername = hdGenerateAdminUsername(hdAdminUsername)
            Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Username Existed. " & hdAdminUsername & " was assigned; please update.<br />"
        End If            
    Else
        hdAdminUsername = hdGenerateAdminUsername("AdminTemp")
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Your username must be 5-20 alphanumeric characters only. " & hdAdminUsername & " was assigned; please update.<br />"
    End If
    
    hdVerifyAdminUsernameUnique = hdAdminUsername

End Function '## hdVerifyAdminUsernameUnique()


'#####################################################################
Sub hdGetCategoriesDDL(theCatTypeID, mySelectedCatID, myOnChangeEvent)
    
    '##On Error Resume Next
   
    Set hdRsGetCat = Server.CreateObject("ADODB.Recordset")
    hdRsGetCat.ActiveConnection = hdDSN
    hdRsGetCat.Source = "SELECT catID,catName FROM hdCategories WHERE cattypeID = " & theCatTypeID & " ORDER BY catName"
    hdRsGetCat.Open()

    Response.Write("<select name=""catID"" id=""catIDID"" size=""1""" & myOnChangeEvent & ">" & vbCrLf)
    Response.Write("<option value=""0"">Category not selected</option>" & vbCrLf)
    While Not hdRsGetCat.EOF
        If mySelectedCatID = hdRsGetCat("catID") Then
            thisSelected = " selected"
        Else
            thisSelected  = ""
        End If

        Response.Write("<option value=""" & hdRsGetCat("catID") & """" & thisSelected & ">" & hdRsGetCat("catName") & "</option>" & vbCrLf)
    
        hdRsGetCat.MoveNext
    Wend
    Response.Write("</select>" & vbCrLf)
    
    hdRsGetCat.Close()
    Set hdRsGetCat = Nothing

End Sub '## hdGetCategoriesDDL()


'##############################################################
Sub hdGetCategoryTypesDDL(mySelectedCatTypeID, myOnChangeEvent)
    
    '##On Error Resume Next
   
    Set hdRsGetCat = Server.CreateObject("ADODB.Recordset")
    hdRsGetCat.ActiveConnection = hdDSN
    hdRsGetCat.Source = "SELECT * FROM hdCategoryType ORDER BY cattypeName"
    hdRsGetCat.Open()

    Response.Write("<select name=""cattypeID"" id=""cattypeIDID"" size=""1""" & myOnChangeEvent & ">" & vbCrLf)
    Response.Write("<option value=""0"">Category Type not selected</option>" & vbCrLf)
    While Not hdRsGetCat.EOF
        If mySelectedcattypeID = hdRsGetCat("cattypeID") Then
            thisSelected = " selected"
        Else
            thisSelected  = ""
        End If

        Response.Write("<option value=""" & hdRsGetCat("cattypeID") & """" & thisSelected & ">" & hdRsGetCat("cattypeName") & "</option>" & vbCrLf)
    
        hdRsGetCat.MoveNext
    Wend
    Response.Write("</select>" & vbCrLf)
    
    hdRsGetCat.Close()
    Set hdRsGetCat = Nothing

End Sub '## hdGetCategoryTypesDDL()


'#######################################
Sub hdCheckFeaturePermission(theFeature)
    
    On Error Resume Next
   
    bContinue = False
    
    Set hdRsFC = Server.CreateObject("ADODB.Recordset")
    hdRsFC.ActiveConnection = hdDSN
    hdRsFC.Source = "SELECT " & theFeature & " FROM hdSiteSettings WHERE setID = 1"
    hdRsFC.Open()

    If hdRsFC.EOF Then
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Database Error!  Contact Administrator.<br />"
    Else
        bContinue = hdRsFC(theFeature)
        If Not bContinue Then Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "Feature Disabled.  Contact Halstead Designs to enable.<br />"
    End If
    
    hdRsFC.Close()
    Set hdRsFC = Nothing
    
    '## if failed, bail.
    If Not bContinue Then Response.Redirect(hdAdminPath)
    
End Sub '## hdCheckFeaturePermission()


'##############################################################
Sub hdGetMainPagesIDDDL(mySelectedwebpgID, myOnChangeEvent)
    
    '##On Error Resume Next
   
    Set hdRsGetPgs = Server.CreateObject("ADODB.Recordset")
    hdRsGetPgs.ActiveConnection = hdDSN
    hdRsGetPgs.Source = "SELECT webpgID, webpgName FROM hdWebPage ORDER BY webpgIsHomePage, webpgIsActive DESC, webpgName"
    hdRsGetPgs.Open()

    Response.Write("<select name=""webpgID"" id=""webpgIDID"" size=""1""" & myOnChangeEvent & ">" & vbCrLf)
    Response.Write("<option value=""0"">Webpage Not Selected</option>" & vbCrLf)
    While Not hdRsGetPgs.EOF
        If mySelectedwebpgID = hdRsGetPgs("webpgID") Then
            thisSelected = " selected"
        Else
            thisSelected  = ""
        End If

        Response.Write("<option value=""" & hdRsGetPgs("webpgID") & """" & thisSelected & ">" & hdRsGetPgs("webpgName") & "</option>" & vbCrLf)
    
        hdRsGetPgs.MoveNext
    Wend
    Response.Write("</select>" & vbCrLf)
    
    hdRsGetPgs.Close()
    Set hdRsGetPgs = Nothing

End Sub '## hdGetCategoryTypesDDL()


'#######################################################
Function hdGetFeatureRecordCount(theTableName,theIDField)
    
    On Error Resume Next
    
    Set hdRsEmCk = Server.CreateObject("ADODB.Recordset")
    hdRsEmCk.ActiveConnection = hdDSN
    hdRsEmCk.Source = "SELECT COUNT(" & theIDField & ") As TotalRecs FROM " & theTableName
    hdRsEmCk.Open()

    If hdRsEmCk.EOF Then 
        hdGetFeatureRecordCount = 0
    Else
        hdGetFeatureRecordCount = hdRsEmCk("TotalRecs")
    End If
   
    hdRsEmCk.Close()
    Set hdRsEmCk = Nothing

End Function '## hdCheckIfAdminUsernameExists()


'#######################################
'##  usually called hddriver to return
'##  the page related banner.  If banners
'##  are linked to webpages, then return
'##  the banner for "thePageID"
'#######################################
Function hdGetThisBannerImage(thePageID)

    thisBannerImage = hdDefaultBannerImage

    If thePageID > 0 Then
        Set rsBnr = Server.CreateObject("ADODB.Recordset")
        rsBnr.ActiveConnection = hdDSN
        rsBnr.Source = "SELECT bannerImage FROM hdBanners WHERE webpgID = " & thePageID
        rsBnr.Open()
    
        If Not rsBnr.EOF Then
            thisBannerImage = rsBnr("bannerImage")
        End If
        
        rsBnr.Close()
        Set rsBnr = Nothing
    
    End If
    
    hdGetThisBannerImage = thisBannerImage

End Function


'#######################################
Function hdGetThisBannerAltTag(thePageID)

    thisAltTag = ""

    If thePageID > 0 Then
        Set rsBnr = Server.CreateObject("ADODB.Recordset")
        rsBnr.ActiveConnection = hdDSN
        rsBnr.Source = "SELECT bannerImageAltTag FROM hdBanners WHERE webpgID = " & thePageID
        rsBnr.Open()
    
        If Not rsBnr.EOF Then
            thisAltTag = rsBnr("bannerImageAltTag")
        End If
        
        rsBnr.Close()
        Set rsBnr = Nothing
    
    End If
    
    hdGetThisBannerAltTag = thisAltTag

End Function

'###################################################
Sub jwPrenatalVideoAccessPage()

    If Session("jwUserProgram") <> "prenatal" Then Response.Redirect(jwVideoPath & "login.asp")

End Sub 

Sub jwFoundationVideoAccessPage()

    If Session("jwUserProgram") <> "foundation" Then Response.Redirect(jwVideoPath & "login.asp")

End Sub 


%>
