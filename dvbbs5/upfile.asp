<!--#include FILE="conn.asp"-->
<!--#include FILE="inc/const.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>文件上传</title>
</head>
<body>
<script>
parent.document.forms[0].Submit.disabled=false;
parent.document.forms[0].Submit2.disabled=false;
</script>
<%
'上传方法，0＝无组件，1＝lyfupload，2＝aspupload，3＝chinaaspupload
dim upload_type
dim upload,file,formName,formPath,iCount,filename,fileExt
upload_type=0

if session("upface")="done" then
	response.write "<font size=2>您已经上传了头像</font>"
	response.end
end if

select case upload_type
case 0
	call upload_0()
case 1
	call upload_1()
case else
	response.write "本系统未开放插件功能"
	response.end
end select

'无组件上传
sub upload_0()
set upload=new upload_5xSoft ''建立上传对象

 formPath=upload.form("filepath")
 ''在目录后加(/)
 if right(formPath,1)<>"/" then formPath=formPath&"/" 


iCount=0
for each formName in upload.file ''列出所有上传了的文件
 set file=upload.file(formName)  ''生成一个文件对象
 if file.filesize<100 then
 	response.write "<font size=2>请先选择你要上传的图片　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if
 	
 if file.filesize>10000 then
 	response.write "<font size=2>图片大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".gif" and fileEXT<>".jpg" then
 	response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
	response.end
 end if 
 
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&file.FileName
 
 if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
  file.SaveAs Server.mappath(filename)   ''保存文件
'  response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"
 response.write "<script>parent.document.forms[0].myface.value='"&FileName&"'</script>"
  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing

session("upface")="done"

Htmend iCount&" 个文件上传结束!"
end sub

'lyfupload
sub upload_1()
	dim obj,Forum_Setting(64),filename
	dim ss,aa1,aa
	Set obj = Server.CreateObject("LyfUpload.UploadFile")
	'大小
    	obj.maxsize=Cint(Forum_Setting(33))*1000
	'类型
    	obj.extname="gif,jpg,bmp,jpeg"
	'重命名
	Forum_Setting(64)=Server.MapPath(obj.request("filepath"))
	'response.write Forum_Setting(64)
	'response.end
	filename=MakedownName(obj.filetype("file1"))
	'response.write filename
	ss=obj.SaveFile("file1",Forum_Setting(64), true,filename)

	if ss= "3" then
   		Response.Write ("文件名重复!")
		response.end
	elseif ss= "0" then
   		Response.Write ("文件尺寸过大!")
		response.end
	elseif ss = "1" then
 		Response.Write ("文件不是指定类型文件!")
		response.end
	elseif ss = "" then
 		Response.Write ("文件上传失败!")
		response.end
	else
		Response.Write "<font size=2>图片上传成功,请copy下边的图片连接,以备后用</font>" 
		response.write "<script>parent.document.forms[0].myface.value='" & obj.request("filepath") & "/" & filename & "'</script>"
		session("upface")="done"
		response.end
	end if
	set obj=nothing
end sub

sub HtmEnd(Msg)
 response.write "<font size=2>图片上传成功,请copy下边的图片连接,以备后用</font>"
 response.end
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