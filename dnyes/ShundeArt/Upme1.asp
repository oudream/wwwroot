<!--#include file="Inc/config1.asp"-->
<!--#include file="Inc/upload.asp"-->
<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<%
const upload_type=0   '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload

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
SavePath = FileUploadPath   '����ϴ��ļ���Ŀ¼
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
	select case upload_type
		case 0
			call upload_0()  'ʹ�û���������ϴ���
		case else
			'response.write "��ϵͳδ���Ų������"
			'response.end
	end select

sub upload_0()    'ʹ�û���������ϴ���
	set upload=new upload_file    '�����ϴ�����
	for each formName in upload.file '�г������ϴ��˵��ļ�
		set file=upload.file(formName)  '����һ���ļ�����
		if file.filesize<100 then
			msg="����ѡ����Ҫ�ϴ����ļ���"
			founderr=true
		end if
		if file.filesize>(MaxFileSize*1024) then
			msg="�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSize) & "K���ļ���"
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
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			founderr=true
		end if

		strJS="<SCRIPT language=javascript>" & vbcrlf
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
			file.SaveToFile Server.mappath(FileName)   '�����ļ�
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
			msg="�ϴ��ļ��ɹ���"
			strJS=strJS & "content=parent.document.flash.a.value;"
			FileType=right(fileExt,3)
			select case FileType
				case "jpg","gif","png","bmp"
					strJS=strJS &"content=content+'<img src=" & filename & " align=left>';" & vbcrlf
				case "mht"
					strJS=strJS & "content=content+'<iframe src=" & filename & " width=640 height=400>';" & vbcrlf
				case "swf"
					strJS=strJS & "content=content+'" & filename & "';" & vbcrlf
			end select
			strJS=strJS & "parent.document.flash.a.value=content;" & vbcrlf
			strJS=strJS & "content=parent.document.flash.j.value;"
			FileType=right(fileExt,3)
			select case FileType
				case "jpg","gif","png","bmp"
					strJS=strJS &"content=content+'<img src=" & filename1 & " align=left>';" & vbcrlf
				case "mht"
					strJS=strJS & "content=content+'<iframe src=" & filename1 & " width=640 height=400>';" & vbcrlf
				case "swf"
					strJS=strJS & "content=content+'" & filename1 & "';" & vbcrlf
			end select
			strJS=strJS & "parent.document.flash.j.value=content;" & vbcrlf
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
		strJS=strJS & "window.location = 'upme1.htm';" & vbcrlf
		strJS=strJS & "</script>"
		response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>