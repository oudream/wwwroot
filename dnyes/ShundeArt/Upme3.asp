<!--#include file="Inc/config3.asp"-->
<!--#include file="Inc/upload.asp"-->
<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<%
const upload_type=0   '上传方法：0=无惧无组件上传类，1=FSO上传 2=lyfupload，3=aspupload，4=chinaaspupload

dim upload,file,formName,SavePath,filename,fileExt
dim upNum
dim EnableUpload
dim Forumupload
dim ranNum
dim uploadfiletype
dim msg,founderr
msg=""
founderr=false
EnableUpload=false
SavePath = FileUploadPath   '存放上传文件的目录
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '在目录后加(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
	select case upload_type
	  case 0
		call upload_0()  '使用化境无组件上传类
	  case else
		'response.write "本系统未开放插件功能"
		'response.end
	end select
%>
</body>
</html>
<%
sub upload_0()    '使用化境无组件上传类
	set upload=new upload_file    '建立上传对象
	for each formName in upload.file '列出所有上传了的文件
		set file=upload.file(formName)  '生成一个文件对象
		if file.filesize<100 then
 			msg="请先选择你要上传的文件！"
			founderr=true
		end if
		if file.filesize>(MaxFileSize*1024) then
 			msg="文件大小超过了限制，最大只能上传" & CStr(MaxFileSize) & "K的文件！"
			founderr=true
		end if

		fileExt=lcase(file.FileExt)
		Forumupload=split(UpFileType,"|")
		for i=0 to ubound(Forumupload)
			if fileEXT=trim(Forumupload(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" or fileEXT="cer" or fileEXT="cdx" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="这种文件类型不允许上传！\n\n只允许上传这几种文件类型：" & UpFileType
			founderr=true
		end if
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			file.SaveToFile Server.mappath(FileName)   '保存文件
           filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt	
		   set rs=server.createobject("adodb.recordset")
           sql="select * from uploadpic" 
           rs.open sql,conn,1,3
           rs.addnew
           rs("picname")=filename1
           username=request.cookies(Forcast_SN)("username")
           rs("username")=username
           rs.update
           rs.close
           set rs=nothing
			msg="上传文件成功！"
 	  		strJS=strJS & "content=parent.document.ifrm.a.value;"
 	  				FileType=right(fileExt,3)
			select case FileType
			    case "jpg","gif","png","bmp"
					strJS=strJS &"content=content+'" & filename & " ';" & vbcrlf
					
				case "mht","htm"
					strJS=strJS & "content=content+'" & filename & "';" & vbcrlf
					
                case "swf"
					strJS=strJS & "content=content+'" & filename & "';" & vbcrlf
				
	  			end select
	  			strJS=strJS & "parent.document.ifrm.a.value=content;" & vbcrlf
	  			
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "window.location = 'upme3.htm';" & vbcrlf
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>
