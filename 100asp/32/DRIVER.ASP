<Html>
<head>
<title>
磁盘驱动器浏览器
</title>
</head>
<body>
<%
Function Tran(Driver)
	Select Case Driver
		Case 0: Tran="设备无法识别"
		Case 1: Tran="软盘驱动器"
		Case 2: Tran="硬盘驱动器"
		Case 3: Tran="网络硬盘驱动器"
		Case 4: Tran="光盘驱动器"
		Case 5: Tran="RAM虚拟磁盘"
	End Select
End Function
set fs=Server.CreateObject("Scripting.FileSystemObject")
%>
<table border=1 width="80%">
<tr>
<td>盘符</td>
<td>类型</td>
<td>卷标</td>
<td>总计大小</td>
<td>可用空间</td>
<td>文件系统</td>
</tr>
<%
on error resume next
For each Drive in fs.Drives
	Response.Write "<Tr>"
	Response.Write "<Td>" & Drive.DriveLetter & " </td>"
	Response.write "<Td> " & Tran(Drive.DriveType) & " </td>"
	Response.write "<Td> " & Drive.VolumeName & " </td>"
	Response.write "<Td> " & cstr(Drive.TotalSize) & " </td>"
	Response.write "<Td> " & cstr(Drive.Availablespace) & " </td>"
	Response.write "<Td> " & (Drive.FileSystem) & " </td>"
	Response.Write "</Tr>"
Next

set fs=nothing

%>
</table>
</body>
</html>

