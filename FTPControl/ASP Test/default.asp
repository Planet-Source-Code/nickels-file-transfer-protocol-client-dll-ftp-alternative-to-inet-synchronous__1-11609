<HTML>
<HEAD>
</HEAD>
<BODY>
<%
	
	ListContentFile = "C:\list.dat"
	
	Function ReadTextFile(FilePath)
		Dim fso, f
		Const ForReading = 1
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.fileexists(FilePath) then
			Set f = fso.OpenTextFile(FilePath, ForReading)
			ReadTextFile = f.ReadAll
			f.Close
		End if
	End Function



	dim ftpObj
	
	set ftpObj = CreateObject("FTPClient.Connection")
	
	ftpObj.Connect "127.0.0.1",21,"anonymous","kltest"
	
	ftpObj.ListContents ListContentFile
	
	Response.write replace(ftpObj.LogData,vbcrlf,"<BR>") & "<BR><BR><BR>"
	
	Response.Write replace(ReadTextFile(ListContentFile),vbcrlf,"<BR>")


	
	ftpObj.Disconnect
	
	set ftpobj = nothing



%>
</BODY>
</HTML>