<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upload.asp"-->
<!--#include file="ChkUser.asp" -->

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
SavePath = (FileUploadPath)   '����ϴ��ļ���Ŀ¼
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body>
<%
	select case upload_type
	  case 0
		call upload_0()  'ʹ�û���������ϴ���
	  case else
		'response.write "��ϵͳδ���Ų������"
		'response.end
	end select
%>
</body>
</html>
<%
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
                else
                        EnableUpload=true
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
            fillle="����"   
			file.SaveToFile Server.mappath(FileName)   '�����ļ�
           filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt	
		   set rs=server.createobject("adodb.recordset")
           sql="select * from "& db_UploadPic_Table
           rs.open sql,conn,1,3
           rs.addnew
           rs("picname")=filename1
           username=request.cookies(Forcast_SN)("username")
           rs("username")=username
           rs.update
           rs.close
           set rs=nothing
			msg="�ϴ��ļ��ɹ���"
 	  		
			strJS=strJS & "content=parent.message.document.body.innerHTML;"
			FileType=right(fileExt,3)
			
			
					strJS=strJS & "content=content+'����ͼ�����������<br><a href="& filename &"><img border=0 src=images/"&FileType&".gif></a><br>"& fillle &"';" & vbcrlf
					
				
	  			strJS=strJS & "parent.message.document.body.innerHTML=content;" & vbcrlf
	  			
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "history.go(-1);" & vbcrlf
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>
