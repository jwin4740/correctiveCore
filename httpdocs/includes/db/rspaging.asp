<%	

If hdRsBcontinue Then

	If hdRsDefaultPageSize > hdRsTotalSummaryRecords Then
		hdRsCurrentRecordTo = hdRsTotalSummaryRecords
		hdRsCurrentRecordFrom = 1
	Else
		hdRsCurrentRecordTo = hdRsCurrentPage * hdRsDefaultPageSize 
		hdRsCurrentRecordFrom = hdRsCurrentRecordTo - hdRsDefaultPageSize + 1		
	End If
	
	If hdRsCurrentRecordTo > hdRsTotalSummaryRecords Then hdRsCurrentRecordTo = hdRsTotalSummaryRecords 
%>

<p>
<%=hdRSRecordTypesTitle%>&nbsp;<%=hdRsCurrentRecordFrom%> to <%=hdRsCurrentRecordTo%> of <%=hdRsTotalSummaryRecords%>
<br />
<%
	If hdRsTotalPages > 1 Then	'##display paging if more than one page worth of data.
		
		If hdRsCurrentPage > 1 Then 

			Response.Write("<a href=""" & Request.ServerVariables("SCRIPT_NAME") & _
				"?cp=" & hdRsCurrentPage-1 & QueryStringEnd & """>" & _
				"&lt;&lt; Previous </a>&nbsp;")

		End If

		For NAVLOOP = 1 TO hdRsTotalPages
			If NAVLOOP MOD hdRSPagingLineBreakCount = 0 Then Response.Write("<br>") '## break appart the number for records returned 
			If NAVLOOP = hdRsCurrentPage Then 
		
				Response.Write("&nbsp;[<b>" & NAVLOOP & "</b>]&nbsp;")
				
			Else

				Response.Write("&nbsp;<a href=""" & Request.ServerVariables("SCRIPT_NAME") & _
				"?cp=" & NAVLOOP & QueryStringEnd & """>" & NAVLOOP & "</a>&nbsp;")
			End If
		Next
		
		If hdRsCurrentPage < hdRsTotalPages Then 
			If (hdRsCurrentRecordTo + hdRsDefaultPageSize) > hdRsTotalSummaryRecords Then
				hdRsNextPages = hdRsTotalSummaryRecords MOD hdRsDefaultPageSize
			Else
				hdRsNextPages = hdRsDefaultPageSize
			End If

			Response.Write("&nbsp;<a href=""" & Request.ServerVariables("SCRIPT_NAME") & _
				"?cp=" & hdRsCurrentPage+1 & QueryStringEnd & """>" & _
				"Next &nbsp;&gt;&gt;</a>")
		End If
		
    End If  '## hdRsTotalPages > 1
%>
</p>
<%
End If  '## hdRsBcontinue
%>
