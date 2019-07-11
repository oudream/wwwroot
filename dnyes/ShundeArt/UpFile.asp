<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<!--#include FILE="upload.inc"-->
<html>
<head>
<title>图片上传</title>

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
'上传方法，0＝无组件，1＝chinaaspupload
dim upload_type
upload_type=upload

dim upNum
dim uploadsuc
dim Forumupload
dim ranNum
dim uploadfiletype
dim upload,file,formName,formPath,iCount,filename,fileExt
upNum=session("upNum")
response.write "<body leftmargin=0 topmargin=0><br><p align=center><span style=""font-family: 宋体; font-size: 9pt""><FONT color=red>"&upNum&"</font></span>"
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

	''在目录后加(/)
	if right(formPath,1)<>"/" then formPath=formPath&"/" 
	response.write "<body leftmargin=5 topmargin=3>"
	iCount=0

	for each formName in upload.file	''列出所有上传了的文件
		set file=upload.file(formName)  ''生成一个文件对象

		if file.filesize<100 then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
			response.end
		end if

		if file.filesize>204000 then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">文件大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
			response.end
		end if

		fileExt=lcase(right(file.filename,4))

		if fileext<>".jpg" and fileext<>".png" and fileext<>".gif" and fileext<>".bmp" then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件格式不允许上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
			response.end
		end if

		randomize
		ranNum=int(90000*rnd)+10000
		filename=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&fileExt

		if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
			file.SaveAs Server.mappath(""& FileUploadPath &""&filename)   ''保存文件
			iCount=iCount+1
		end if

		set file=nothing
	next
	set upload=nothing
	Htmend iCount&" 个文件上传结束!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
	session("filename")=filename
	session("upNum")=upNum+1
	response.write "<span style=""font-family: 宋体; font-size: 9pt"">文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]<br><br>[ <a href='"&filename&"' target='_blank'>预览该图片</a> ]<br><br>[ <a href=# onclick=""Addpic('"&filename&"')"">点击这里添加到编辑器中</a> ]<br><br>如果您在文档中插入了图片，请在添加文章页面的<font color=red>图片新闻</font>处选择“<font color=red>是</font>。</span>"
	response.end
end sub

sub upload_1()
	set FileUp=server.createobject("ChinaASP.UpLoad") ''建立上传对象
	filepath=server.MapPath(""& FileUploadPath &"")
	response.write "<body leftmargin=5 topmargin=3>"
	iCount=0
	for each f in fileup.Files ''列出所有上传了的文件

		if f.filesize<100 then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">请先选择你要上传的文件　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
			response.end
		end if

		if f.filesize>204000 then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">文件大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
			response.end
		end if

		fileExt=lcase(right(f.filename,4))

		if fileext<>".jpg" and fileext<>".png" and fileext<>".gif" and fileext<>".bmp" then
			response.write "<span style=""font-family: 宋体; font-size: 9pt"">该文件格式不允许上传　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</span>"
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
		
		if f.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
			f.saveas filePath & "\"&filename   ''保存文件
			iCount=iCount+1
		end if
		set f=nothing
	next
	
	set FileUp=nothing
	Htmend iCount&" 个文件上传结束!"
end sub

sub HtmEnd(Msg)
	if upNum="" then upNum=1
		session("filename")=filename
		session("upNum")=upNum+1
		response.write "<span style=""font-family: 宋体; font-size: 9pt"">文件上传成功 [ <a href=# onclick=history.go(-1)>继续上传</a> ]<br><br>[ <a href='"& FileUploadPath & filename&"' target='_blank'>预览该图片</a> ]<br><br>[ <a href=# onclick=""Addpic('"FileUploadPath & filename&"')"">点击这里添加到编辑器中</a> ]<br><br>如果您在文档中插入了图片，请在添加文章页面的<font color=red>图片新闻</font>处选择“<font color=red>是</font>。</span>"
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
	<a href="javascript:window.close()">关闭窗口</a>
	<br>
</body>
</html>