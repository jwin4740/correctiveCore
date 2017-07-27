<%	
Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
PID = "PID=" & UploadProgress.CreateProgressID()
barref = hdAdminPath & "includes/framebar.asp?to=10&" & PID
%>
