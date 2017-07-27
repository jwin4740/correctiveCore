<%

If hsSQL <> "" Then 
    hdRsCurrentPage = Request.QueryString("cp")
    '## check which page was selected - defaults to 1st page
    If hdRsCurrentPage = "" Or Not(IsNumeric(hdRsCurrentPage)) Then 
	    hdRsCurrentPage = 1
    Else
	    hdRsCurrentPage = CInt(hdRsCurrentPage)
    End If

    hdRsSortOrder	= Request.QueryString("s")
    hdRsBcontinue = True

    Set hdConn = Server.CreateObject("ADODB.Connection")
    hdConn.Open hdDSN
    Set hdRs = Server.CreateObject("ADODB.Recordset")
    hdRs.CursorLocation = adUseClient
    hdRs.CursorType = adOpenStatic
    hdRs.CacheSize = hdRsDefaultPageSize
    hdRs.Open hsSQL, hdConn,,,adCmdText
    hdRs.PageSize = hdRsDefaultPageSize
    hdRsTotalPages = hdRs.PageCount
    If hdRsCurrentPage > hdRsTotalPages Or hdRsCurrentPage < 1 Then hdRsCurrentPage = 1
    If Not(hdRs.EOF) Then hdRs.AbsolutePage = hdRsCurrentPage
    If hdRs.BOF And hdRs.EOF Then 
        Session("hdAdminErrorMsg") = Session("hdAdminErrorMsg") & "No Records Found."
        hdRsBcontinue = False
    Else
        hdRsTotalSummaryRecords = hdRs.RecordCount
    End If
End If

%>