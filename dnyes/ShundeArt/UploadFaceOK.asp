<!--#include file="Inc/upload.asp"-->
<!--#include file="conn.asp" -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp" -->
<!--#include file="chkuser.asp" -->
<%
const upload_type=0   '�ϴ�������0=������ϴ���
Const UpFileType="png|gif|jpg|bmp"        '������ϴ��ļ�����
Const MaxFileSize=200        '�ϴ��ļ���С����

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

if UserTableType = "Dvbbs" then		''�ж��Ƿ�����DVBBS��������·��ָ�� BBS ͷ��Ŀ¼��
	SavePath = FileUploadPath   '����ϴ��ļ���Ŀ¼
else
	SavePath = FaceUploadPath   '����ϴ��ļ���Ŀ¼
end if

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
		end if
		if EnableUpload=false then
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & UpFileType
			founderr=true
		end if
		
		if founderr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt
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
		end if
		response.write "<script>parent.document.FrmAddLink.photo.value='"& filename &"';"
		response.write "window.location = 'UploadFace.asp';" & vbcrlf
		response.write "alert('" & msg & "');</script>" & vbcrlf
		set file=nothing
	next
	set upload=nothing
end sub
%>