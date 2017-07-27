<%
'## get recordset paging querystrings

    hdRsCurrentPage = Request.QueryString("cp")
    '## check which page was selected - defaults to 1st page
    If hdRsCurrentPage = "" Or Not(IsNumeric(hdRsCurrentPage)) Then 
	    hdRsCurrentPage = 1
    Else
	    hdRsCurrentPage = CInt(hdRsCurrentPage)
    End If

    hdRsSortOrder = Request.QueryString("s")

%>