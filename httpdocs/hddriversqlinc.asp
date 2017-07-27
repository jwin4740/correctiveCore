<%

'## CODE USED FOR DRIVER PAGES TO READ DATA FROM DB IF REQURIED.
'On Error Resume Next

bContinue = True

'## check the calling page to see if it's a CMS page and load the record set.

hdScriptName = Trim(Request.ServerVariables("SCRIPT_NAME"))
hdScriptName = Mid(hdScriptName, (InstrRev(hdScriptName, "/", Len(hdScriptName), 1)+1))

Set rsPage = Server.CreateObject("ADODB.Recordset")
rsPage.ActiveConnection = hdDSN
rsPage.Source = "SELECT * FROM hdWebPage WHERE webpgFileName = '" & hdScriptName & "'"
rsPage.CursorType = adOpenForwardOnly
rsPage.CursorLocation = adUseServer
rsPage.LockType = adLockReadOnly
rsPage.Open()

If Err <> 0 Then bContinue = False

If bContinue Then
    If Not rsPage.EOF Then 
        '## set needs vars.
		webpgID = rsPage("webpgID")
        webpgTitle = rsPage("webpgTitle")
        webpgName = rsPage("webpgName")
        webpgContent = rsPage("webpgContent")
        webpgMetaKeywords = rsPage("webpgMetaKeywords")
        webpgMetaDescription = rsPage("webpgMetaDescription")
		webpgFileName = rsPage("webpgFileName")
    End If

    rsPage.Close()
    Set rsPage = Nothing
    
    If webpgMetaKeywords = "" Then webpgMetaKeywords = Session("setDefaultMetaKeywords")
    If webpgMetaDescription = "" Then webpgMetaDescription = Session("setDefaultMetaDescription")
End If  

%>