<!--#include FILE="conn.asp"-->
<!--#include FILE="upload.inc"-->
<!-- #include file="inc/const.asp" -->
<html>
<head>
<title>�ļ��ϴ�</title>
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>>

<script>
parent.document.forms[0].Submit.disabled=false;
parent.document.forms[0].Submit2.disabled=false;
</script>
<table width="100%" border=0 cellspacing=0 cellpadding=0>
<tr><td class=tablebody1 valign=top height=40>
<%
response.buffer=true
'�ϴ�������0���������1��lyfupload��2��aspupload��3��chinaaspupload
dim upload_type
upload_type=0

dim upload,file,formName,formPath,iCount,filename,fileExt
dim upNum
dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfiletype
upNum=request.cookies("aspsky")("upNum")
'response.write request.cookies("aspsky")("upNum")
'response.end
if not isnumeric(upnum) then upnum=1

if Cint(GroupSetting(7))=0 then
	response.write "��û���ڱ���̳�ϴ��ļ���Ȩ��"
	response.end
end if
if not founduser then
 	response.write "ֻ�е�½�û������ϴ���"
	response.end
end if
if cint(upNum)>Cint(GroupSetting(40)) then
 	response.write "һ��ֻ���ϴ�"&GroupSetting(40)&"���ļ���"
	response.end
end if
response.write "<FONT color=red>"&upNum&"</font>"
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
'��Ŀ¼���(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 

iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
	set file=upload.file(formName)  ''����һ���ļ�����
	if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ����ļ���[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
	end if
 	
	if file.filesize>cint(GroupSetting(44))*1000 then
 	response.write "<font size=2>�ļ���С���������� "&Forum_Setting(33)&"K��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
	end if

	fileExt=lcase(right(file.filename,4))
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
	if fileEXT=".asp" and fileEXT=".asa" and fileEXT=".aspx" then
		uploadsuc=false
	end if
	if uploadsuc=false then
 	response.write "�ļ���ʽ����ȷ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	response.end
	end if

	randomize
	ranNum=int(90000*rnd)+10000
	filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

	if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
	file.SaveAs Server.mappath(FileName)   ''�����ļ�
	for i=0 to ubound(Forumupload)
		if fileEXT="."&trim(Forumupload(i)) then
	 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&Forumupload(i)&"]"&FileName&"[/upload]'</script>"
		exit for
		end if
	next
	iCount=iCount+1
	end if
	set file=nothing
next
set upload=nothing

Htmend iCount&" ���ļ��ϴ�����!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
	response.cookies("aspsky")("upNum")=upNum+1
	Response.Cookies("aspsky").path=cookiepath
	response.write "�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	response.end
end sub

'lyfupload
sub upload_1()
	dim obj,Forum_Setting(64),filename
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
    	obj.maxsize=Cint(GroupSetting(44))*1000
	'����
    	obj.extname=Forum_upload
	'������
	Forum_Setting(64)=Server.MapPath(obj.request("filepath"))
	'response.write obj.extname
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",Forum_Setting(64), true,filename)

	if ss= "3" then
   		Response.Write ("<font size=2>�ļ����ظ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss= "0" then
   		Response.Write ("<font size=2>�ļ��ߴ����!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss = "1" then
 		Response.Write ("<font size=2>�ļ�����ָ�������ļ�!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	elseif ss = "" then
 		Response.Write ("<font size=2>�ļ��ϴ�ʧ��!��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]")
		response.end
	else
		if upNum="" then upNum=1
		response.cookies("aspsky")("upNum")=upNum+1
		response.write "<font size=2>�ļ��ϴ��ɹ� [ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&uploadfiletype&"]"&obj.request("filepath") & "/" & filename&"[/upload]'</script>"
		'response.write "<script>parent.document.forms[0].myface.value='" & obj.request("filepath") & "/" & filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
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
</td></tr>
</table>
</body>
</html>