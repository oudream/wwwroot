<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>ͼƬ�ϴ�</title>

<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to add pic
function Addpic(imagePath){	
	window.opener.frames.message.focus();								
	window.opener.frames.message.document.execCommand('InsertImage', false, imagePath);
window.close();
}

// -->
</script>
</head>
<body>

<%
'�ϴ�������0���������1��chinaaspupload
dim upload_type
upload_type=upload

dim upNum
dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfiletype
dim upload,file,formName,formPath,iCount,filename,fileExt
upNum=session("upNum")
response.write "<body leftmargin=0 topmargin=0><br><p align=center><span style=""font-family: ����; font-size: 9pt""><FONT color=red>"&upNum&"</font></span>"
select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "��ϵͳδ���Ų������"
	response.end
end select

sub upload_0()
	set upload=new upload_5xSoft ''�����ϴ�����
	formPath=upload.form("filepath")

	''��Ŀ¼���(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	response.write "<body leftmargin=5 topmargin=3>"
	iCount=0

	for each formName in upload.file	''�г������ϴ��˵��ļ�
		set file=upload.file(formName)  ''����һ���ļ�����

		if file.filesize<100 then
			response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if

		if file.filesize>204000 then
			response.write "<span style=""font-family: ����; font-size: 9pt"">�ļ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if

		fileExt=lcase(right(file.filename,4))

		if fileext<>".jpg" and fileext<>".png" and fileext<>".gif" and fileext<>".bmp" then
			response.write "<span style=""font-family: ����; font-size: 9pt"">���ļ���ʽ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if

		randomize
		ranNum=int(90000*rnd)+10000
		filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

		if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
			file.SaveAs Server.mappath(""& FileUploadPath &""&filename)   ''�����ļ�
			iCount=iCount+1
		end if

		set file=nothing
	next
	set upload=nothing
	Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
	session("filename")=filename
	session("upNum")=upNum+1
	response.write "<span style=""font-family: ����; font-size: 9pt"">�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]<br><br>[ <a href='"&filename&"' target='_blank'>Ԥ����ͼƬ</a> ]<br><br>[ <a href=# onclick=""Addpic('"&filename&"')"">���������ӵ��༭����</a> ]<br><br>��������ĵ��в�����ͼƬ�������������ҳ���<font color=red>ͼƬ����</font>��ѡ��<font color=red>��</font>��</span>"
	response.end
end sub

sub upload_1()
	set FileUp=server.createobject("ChinaASP.UpLoad") ''�����ϴ�����
	filepath=server.MapPath(""& FileUploadPath &"")
	response.write "<body leftmargin=5 topmargin=3>"
	iCount=0
	for each f in fileup.Files ''�г������ϴ��˵��ļ�

		if f.filesize<100 then
			response.write "<span style=""font-family: ����; font-size: 9pt"">����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if

		if f.filesize>204000 then
			response.write "<span style=""font-family: ����; font-size: 9pt"">�ļ���С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if

		fileExt=lcase(right(f.filename,4))

		if fileext<>".jpg" and fileext<>".png" and fileext<>".gif" and fileext<>".bmp" then
			response.write "<span style=""font-family: ����; font-size: 9pt"">���ļ���ʽ�������ϴ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</span>"
			response.end
		end if
		
		uploadsuc=false
		Forumupload=split(Forum_upload,",")
		
		for i=0 to ubound(Forumupload)
			if fileEXT="."&trim(Forumupload(i)) then
				uploadsuc=true
				exit for
			else
				uploadsuc=false
			end if
		next

		randomize
		ranNum=int(90000*rnd)+10000
		filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt
		
		if f.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
			f.saveas filePath & "\"&filename   ''�����ļ�
			iCount=iCount+1
		end if
		set f=nothing
	next
	
	set FileUp=nothing
	Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
		session("filename")=filename
		session("upNum")=upNum+1
		response.write "<span style=""font-family: ����; font-size: 9pt"">�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]<br><br>[ <a href='"& FileUploadPath & filename&"' target='_blank'>Ԥ����ͼƬ</a> ]<br><br>[ <a href=# onclick=""Addpic('"FileUploadPath & filename&"')"">���������ӵ��༭����</a> ]<br><br>��������ĵ��в�����ͼƬ�������������ҳ���<font color=red>ͼƬ����</font>��ѡ��<font color=red>��</font>��</span>"
		response.end
end sub

'--------������ת�����ļ���--------
function MakedownName(filename)
	dim fname,Forumupload,u
	fname = now()
	fname = replace(fname,"-","")
	fname = replace(fname," ","") 
	fname = replace(fname,":","")
	fname = replace(fname,"PM","")
	fname = replace(fname,"AM","")
	fname = replace(fname,"����","")
	fname = replace(fname,"����","")
	fname = int(fname) + int((10-1+1)*Rnd + 1)
	Forumupload=split(Forum_upload,",")
	for u=0 to ubound(Forumupload)
		if instr(filename,Forumupload(u))>0 then
			uploadfiletype=Forumupload(u)
			MakedownName=fname & "." & Forumupload(u)
			exit for
		end if
	next
end function
%>
<div align="center">
	<br>
	<br>
	<a href="javascript:window.close()">�رմ���</a>
	<br>
</body>
</html>