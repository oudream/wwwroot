<%
''Option Explicit 
Server.ScriptTimeout = 999     '要上传较大的文件就要开启此项

Dim OKAr
Dim OKsize

'***************设置允许上传的文件类型********************

'OKAr = split("jpg,gif,rar,zip,txt,doc,xls",",")
OKAr = split("jpg,gif",",")

'***********************************

'***************设置允许上传的文件大小********************

OKsize = 153600

'***********************************

Function Bytes2bStr(vin)
if lenb(vin) =0 then
	Bytes2bStr = ""
	exit function
end if
''二进制转换为字符串
 Dim BytesStream,StringReturn
 Set BytesStream = Server.CreateObject("ADODB.Stream")
 BytesStream.Type = 2 
 BytesStream.Open
 BytesStream.WriteText vin
 BytesStream.Position = 0
 BytesStream.Charset = "gb2312"
 BytesStream.Position = 2
 StringReturn = BytesStream.ReadText
 BytesStream.close
 Set BytesStream = Nothing
 Bytes2bStr = StringReturn
End Function

Function Myrequest(fldname)
	''取表单数据
	''支持对同名表单域的读取
	dim i
	dim fldHead
	dim tmpvalue
	for i = 0 to loopcnt-1
		fldHead = fldInfo(i,0)
		if instr(lcase(fldHead),lcase(fldname))>0 then
			''表单在数组中
			''判断该表单域内容
			tmpvalue = FldInfo(i,1)
			if instr(fldHead,"filename=""")<1 then
				Tmpvalue = Bytes2bStr(tmpvalue)
				if myrequest <> "" then 
					myrequest = myrequest & "," &tmpvalue
				else
					MyRequest = tmpvalue
				end if				
			else
				myrequest = tmpvalue
			end if				
		end if
	next
end function

function GetFileName(fldName)
	''都取原上传文件文件名
	dim i
	dim fldHead
	dim fnpos
	for i = 0 to loopcnt-1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 then
			fnpos = instr(fldHead,"filename=""")
			if fnpos < 1 then exit for
			fldHead = mid(fldHead,fnpos+10)
			''表单内容
			GetFileName = mid(fldHead,1,instr(fldHead,"""")-1)
			GetfileName = mid(GetFileName,instrRev(GetFileName,"\")+1)
		end if
	next
end function

function GetContentType(fldName)
	''读取上传文件的类型
	''限定读取文件域的内容
	dim i
	dim fldHead,cpos
	for i = 0 to loopcnt - 1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 and instr(fldHead,"filename=""") >0 then
			''读取contentType
			cpos = instr(fldHead,"content-type: ")
			GetContentType = mid(fldHead,cpos+14)
		end if
	next
end function

Sub SaveToFile(fd,path,fname)
''保存文件''参数说明:''fd:byte（）类型数据，文件内容''path:保存路径后面必须带"/"''fname:文件名
	dim Fstream
	Set FStream = Server.CreateObject("adodb.stream")
	fstream.mode = 3
	fstream.type = 1
	fstream.open
	fstream.position = 0
	fstream.Write fd
	fstream.savetofile Server.Mappath(path&fname),2
	fstream.close
	set fstream = nothing
end sub

Function GetFileTypeName(Fldname)
If instr(Fldname,".") > 0 Then
GetFileTypeName = right(Fldname,3)
Else
Response.Write "文件名非法，请修改后再上传"
Response.End()
End If
End Function

Function IsvalidFile(TypeName)   '限制上传文件类型
IsvalidFile = False
Dim GName
For Each GName in OKAr
If TypeName = GName Then
IsvalidFile = True
Exit For
End If
Next
End Function

%>