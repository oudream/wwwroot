<%
''Option Explicit 
Server.ScriptTimeout = 999     'Ҫ�ϴ��ϴ���ļ���Ҫ��������

Dim myconn,connstr
connstr = "Provider=Microsoft.Jet.Oledb.4.0;data source="&Server.MapPath("Upload.mdb")
set myconn = Server.CreateObject("adodb.connection")
myconn.open connstr

Dim OKAr
Dim OKsize

Set rs = myconn.execute("select * from Config")
'***************���������ϴ����ļ�����********************

OKAr = split(rs(0),",")

'***********************************

'***************���������ϴ����ļ���С********************

OKsize = Clng(rs(1))

'***********************************

Function Bytes2bStr(vin)
if lenb(vin) =0 then
	Bytes2bStr = ""
	exit function
end if
''������ת��Ϊ�ַ���
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
	''ȡ������
	''֧�ֶ�ͬ������Ķ�ȡ
	dim i
	dim fldHead
	dim tmpvalue
	for i = 0 to loopcnt-1
		fldHead = fldInfo(i,0)
		if instr(lcase(fldHead),lcase(fldname))>0 then
			''����������
			''�жϸñ�������
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
	''��ȡԭ�ϴ��ļ��ļ���
	dim i
	dim fldHead
	dim fnpos
	for i = 0 to loopcnt-1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 then
			fnpos = instr(fldHead,"filename=""")
			if fnpos < 1 then exit for
			fldHead = mid(fldHead,fnpos+10)
			''������
			GetFileName = mid(fldHead,1,instr(fldHead,"""")-1)
			GetfileName = mid(GetFileName,instrRev(GetFileName,"\")+1)
		end if
	next
end function
function GetContentType(fldName)
	''��ȡ�ϴ��ļ�������
	''�޶���ȡ�ļ��������
	dim i
	dim fldHead,cpos
	for i = 0 to loopcnt - 1
		fldHead = lcase(fldInfo(i,0))
		if instr(fldHead,lcase(fldName)) > 0 and instr(fldHead,"filename=""") >0 then
			''��ȡcontentType
			cpos = instr(fldHead,"content-type: ")
			GetContentType = mid(fldHead,cpos+14)
		end if
	next
end function
Sub SaveToFile(fd,path,fname)
''�����ļ�''����˵��:''fd:byte�����������ݣ��ļ�����''path:����·����������"/"''fname:�ļ���
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
Response.Write "�ļ����Ƿ������޸ĺ����ϴ�"
Response.End()
End If
End Function

Function IsvalidFile(TypeName)   '�����ϴ��ļ�����
IsvalidFile = False
Dim GName
For Each GName in OKAr
If TypeName = GName Then
IsvalidFile = True
Exit For
End If
Next
End Function

Sub CheckLogin()
If Session("Admin") <> "OK" Then
Logined = False
Response.Write ("δͨ�������֤����<a href=Index.asp>����</a>")
Response.Redirect ("Index.asp")
Response.End()
End If
End Sub

%>