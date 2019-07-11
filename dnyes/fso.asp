<%
Dim fso, f1, ts, s
dim file,file2
file=Server.MapPath("testfile.txt")
file2=Server.MapPath("testfile2.txt")

ON Error Resume Next
  Const ForReading = 1
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f1 = fso.CreateTextFile(file, True)
  ' 写一行。
  if err.number =0 then
	Response.Write "<li>FSO创建文件完成 <br>"
  elseif err.number = 70 then
	response.write "<li>文件"&file&"已存在并被锁定!"
	response.end
  elseif err.number = 424 then
	response.write "<li>空间不支持FSO"
	response.end
  elseif err.number <> 0 then
	response.write "<li>其他未知的错误,错误编号="&err.number
	response.end
  end if
  f1.WriteLine "Hello World"
  f1.WriteBlankLines(1)
  f1.Close
  ' 读取文件的内容。
  if err.number = 0 then
	Response.Write "<li>FSO写入内容完成<br>"
  else
	response.write "<li>其他未知的错误,错误编号="&err.number
	response.end
  end if
  Set ts = fso.OpenTextFile(file, ForReading)
  s = ts.ReadLine
  Response.Write "<li>FSO读取文件完成:文件内容为 '" & s & "'"
  ts.Close
  set fso=nothing
dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
objFSO.movefile ""&file&"",""&file2&""
if err.number =0 then
	Response.Write "<li>FSO改名完成"
elseif err.number = 53 then
	response.write "<li>"& file &"不存在!"
	response.end
elseif err.number = 58 then
	response.write "<li>"& file2 &"已存在并被锁定!"
	response.end
elseif err.number = 70 then
	response.write "<li>没有开放FSO修改权限!"
	response.write "<li>没有开放FSO删除权限!"
	response.end
elseif err.number <> 0 then
	response.write "<li>其他未知的错误,错误编号="&err.number
	response.end
end if
objFSO.DeleteFile (file2),true
if err.number =0 then
	Response.Write "<li>FSO删除完成"
end if
set objFSO=nothing
if err.number = 53 then
	response.write "<li>"& file2 &"不存在!"
	response.end
elseif err.number = 58 then
	response.write "<li>"& file2 & "已存在!"
	response.end
elseif err.number = 70 then
	response.write "<li>没有开放FSO删除权限!"
	response.end
elseif err.number <> 0 then
	response.write "<li>其他未知的错误,错误编号="&err.number
	response.end
end if
%>