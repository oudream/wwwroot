<!--#include file="dbm_upload.inc" -->
<%

Aguest_allname	= Split(session("guest_allname"), ",", -1, 1)

set upload=new upload_5xSoft ''�����ϴ�����

UName=upload.form("ToPath")

set file=upload.file("LocalFile")  ''����һ���ļ����� 



If session("user_adminlevel")=1 Then

if UName<>"" then
FUpFolderPath_num=0
AUpFolderPath=Split(UName, "\", -1, 1)
for i=0 to ubound(AUpFolderPath)
if AUpFolderPath(i)="sosoon" then AUpFolderPath_num = AUpFolderPath_num + 1
next
else
AUpFolderPath_num=100
end if
if AUpFolderPath_num<1 then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' �����ϴ���Χ��');
window.history.go( -1 );
</SCRIPT>
<%
response.End()
end if

end if




If session("user_adminlevel")=2 Then

if UName<>"" then
FUpFolderPath_num=0
AUpFolderPath=Split(UName, "\", -1, 1)
for i=0 to ubound(AUpFolderPath)
if AUpFolderPath(i)=session("user_allname") then AUpFolderPath_num = AUpFolderPath_num + 1
next
else
AUpFolderPath_num=100
end if
if AUpFolderPath_num<1 then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' �����ϴ���Χ��');
window.history.go( -1 );
</SCRIPT>
<%
response.End()
end if

end if



If session("user_adminlevel")=9 Then

if UName<>"" then
FUpFolderPath_num=0
AUpFolderPath=Split(UName, "\", -1, 1)
for i=0 to ubound(AUpFolderPath)
for j=1 to ubound(Aguest_allname)
if AUpFolderPath(i)=Aguest_allname(j) then AUpFolderPath_num = AUpFolderPath_num + 1
next
next
else
AUpFolderPath_num=100
end if
if AUpFolderPath_num<1 then
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' �����ϴ���Χ��');
window.history.go( -1 );
</SCRIPT>
<%
response.End()
end if

end if




    if UName="" Or file.filesize < 1 or file.filesize>2097152  then  
	Response.Write "<font size=2>��ѡ����ʵ��ļ�(ҪС��2M)��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
    response.end
    else
	  file.SaveAs UName   ''�����ļ�
%>
<SCRIPT LANGUAGE=JavaScript>
alert(' �ļ��ϴ��ɹ�. ');
window.history.go( -1 );
</SCRIPT>
<%	  
	End If
    set file=nothing
    set upload=nothing
%>