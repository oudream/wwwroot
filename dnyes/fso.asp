<%
Dim fso, f1, ts, s
dim file,file2
file=Server.MapPath("testfile.txt")
file2=Server.MapPath("testfile2.txt")

ON Error Resume Next
  Const ForReading = 1
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set f1 = fso.CreateTextFile(file, True)
  ' дһ�С�
  if err.number =0 then
	Response.Write "<li>FSO�����ļ���� <br>"
  elseif err.number = 70 then
	response.write "<li>�ļ�"&file&"�Ѵ��ڲ�������!"
	response.end
  elseif err.number = 424 then
	response.write "<li>�ռ䲻֧��FSO"
	response.end
  elseif err.number <> 0 then
	response.write "<li>����δ֪�Ĵ���,������="&err.number
	response.end
  end if
  f1.WriteLine "Hello World"
  f1.WriteBlankLines(1)
  f1.Close
  ' ��ȡ�ļ������ݡ�
  if err.number = 0 then
	Response.Write "<li>FSOд���������<br>"
  else
	response.write "<li>����δ֪�Ĵ���,������="&err.number
	response.end
  end if
  Set ts = fso.OpenTextFile(file, ForReading)
  s = ts.ReadLine
  Response.Write "<li>FSO��ȡ�ļ����:�ļ�����Ϊ '" & s & "'"
  ts.Close
  set fso=nothing
dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
objFSO.movefile ""&file&"",""&file2&""
if err.number =0 then
	Response.Write "<li>FSO�������"
elseif err.number = 53 then
	response.write "<li>"& file &"������!"
	response.end
elseif err.number = 58 then
	response.write "<li>"& file2 &"�Ѵ��ڲ�������!"
	response.end
elseif err.number = 70 then
	response.write "<li>û�п���FSO�޸�Ȩ��!"
	response.write "<li>û�п���FSOɾ��Ȩ��!"
	response.end
elseif err.number <> 0 then
	response.write "<li>����δ֪�Ĵ���,������="&err.number
	response.end
end if
objFSO.DeleteFile (file2),true
if err.number =0 then
	Response.Write "<li>FSOɾ�����"
end if
set objFSO=nothing
if err.number = 53 then
	response.write "<li>"& file2 &"������!"
	response.end
elseif err.number = 58 then
	response.write "<li>"& file2 & "�Ѵ���!"
	response.end
elseif err.number = 70 then
	response.write "<li>û�п���FSOɾ��Ȩ��!"
	response.end
elseif err.number <> 0 then
	response.write "<li>����δ֪�Ĵ���,������="&err.number
	response.end
end if
%>