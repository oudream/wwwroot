<!--#include FILE="conn.asp"-->
<!--#include FILE="upload.inc"-->
<!-- #include file="inc/const.asp" -->
<html>
<head>
<title>文件上传</title>
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
'上传方法，0＝无组件，1＝lyfupload，2＝aspupload，3＝chinaaspupload
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
	response.write "您没有在本论坛上传文件的权限"
	response.end
end if
if not founduser then
 	response.write "只有登陆用户方能上传！"
	response.end
end if
if cint(upNum)>Cint(GroupSetting(40)) then
 	response.write "一次只能上传"&GroupSetting(40)&"个文件！"
	response.end
end if
response.write "<FONT color=red>"&upNum&"</font>"
select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "本系统未开放插件功能"
	response.end
end select

sub upload_0()
set upload=new upload_5xSoft ''建立上传对象

formPath=upload.form("filepath")
'在目录后加(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 

iCount=0
for each formName in upload.file ''列出所有上传了的文件
	set file=upload.file(formName)  ''生成一个文件对象
	if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
	end if
 	
	if file.filesize>cint(GroupSetting(44))*1000 then
 	response.write "<font size=2>文件大小超过了限制 "&Forum_Setting(33)&"K　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
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
 	response.write "文件格式不正确　[ <a href=# onclick=history.go(-1)>重新上传</a> ]"
	response.end
	end if

	randomize
	ranNum=int(90000*rnd)+10000
	filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

	if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
	file.SaveAs Server.mappath(FileName)   ''保存文件
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

Htmend iCount&" 个文件上传结束!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
	response.cookies("aspsky")("upNum")=upNum+1
	Response.Cookies("aspsky").path=cookiepath
	response.write "文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]"
	response.end
end sub

'lyfupload
sub upload_1()
	dim obj,Forum_Setting(64),filename
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
    	obj.maxsize=Cint(GroupSetting(44))*1000
	'类型
    	obj.extname=Forum_upload
	'重命名
	Forum_Setting(64)=Server.MapPath(obj.request("filepath"))
	'response.write obj.extname
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",Forum_Setting(64), true,filename)

	if ss= "3" then
   		Response.Write ("<font size=2>文件名重复!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss= "0" then
   		Response.Write ("<font size=2>文件尺寸过大!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss = "1" then
 		Response.Write ("<font size=2>文件不是指定类型文件!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	elseif ss = "" then
 		Response.Write ("<font size=2>文件上传失败!　[ <a href=# onclick=history.go(-1)>重新上传</a> ]")
		response.end
	else
		if upNum="" then upNum=1
		response.cookies("aspsky")("upNum")=upNum+1
		response.write "<font size=2>文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]</font>"
	 	response.write "<script>parent.frmAnnounce.Content.value+='[upload="&uploadfiletype&"]"&obj.request("filepath") & "/" & filename&"[/upload]'</script>"
		'response.write "<script>parent.document.forms[0].myface.value='" & obj.request("filepath") & "/" & filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

'--------将日期转化成文件名--------
function MakedownName(filename)
  dim fname,Forumupload,u
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
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