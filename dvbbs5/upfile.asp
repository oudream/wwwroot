<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>�ļ��ϴ�</title>
</head>
<body>
<script>
parent.document.forms[0].Submit.disabled=false;
parent.document.forms[0].Submit2.disabled=false;
</script>
<%
'�ϴ�������0���������1��lyfupload��2��aspupload��3��chinaaspupload
dim upload_type
dim upload,file,formName,formPath,iCount,filename,fileExt
upload_type=0

if session("upface")="done" then
	response.write "<font size=2>���Ѿ��ϴ���ͷ��</font>"
	response.end
end if

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "��ϵͳδ���Ų������"
	response.end
end select

'������ϴ�
sub upload_0()
set upload=new upload_5xSoft ''�����ϴ�����

 formPath=upload.form("filepath")
 ''��Ŀ¼���(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "<font size=2>����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>10000 then
 	response.write "<font size=2>ͼƬ��С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
  file.SaveAs Server.mappath(filename)   ''�����ļ�
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" �ɹ�!<br>"
 response.write "<script>parent.document.forms[0].myface.value='"&FileName&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing

session("upface")="done"

Htmend iCount&" ���ļ��ϴ�����!"
end sub

'lyfupload
sub upload_1()
	dim obj,Forum_Setting(64),filename
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'��С
    	obj.maxsize=Cint(Forum_Setting(33))*1000
	'����
    	obj.extname="gif,jpg,bmp,jpeg"
	'������
	Forum_Setting(64)=Server.MapPath(obj.request("filepath"))
	'response.write Forum_Setting(64)
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",Forum_Setting(64), true,filename)

	if ss= "3" then
   		Response.Write ("�ļ����ظ�!")
		response.end
	elseif ss= "0" then
   		Response.Write ("�ļ��ߴ����!")
		response.end
	elseif ss = "1" then
 		Response.Write ("�ļ�����ָ�������ļ�!")
		response.end
	elseif ss = "" then
 		Response.Write ("�ļ��ϴ�ʧ��!")
		response.end
	else
		Response.Write "<font size=2>ͼƬ�ϴ��ɹ�,��copy�±ߵ�ͼƬ����,�Ա�����</font>" 
		response.write "<script>parent.document.forms[0].myface.value='" & obj.request("filepath") & "/" & filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

sub HtmEnd(Msg)
 response.write "<font size=2>ͼƬ�ϴ��ɹ�,��copy�±ߵ�ͼƬ����,�Ա�����</font>"
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
  Forum_upload="gif,jpg,bmp"
  Forumupload=split(Forum_upload,",")
  for u=0 to ubound(Forumupload)
		if instr(filename,Forumupload(u))>0 then
			MakedownName=fname & "." & Forumupload(u)
			exit for
		end if
  next
end function

%>
</body>
</html>