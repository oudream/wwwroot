<!--#include file="dbm_upload.inc"-->
<!--#include file="dbm_adminconn.asp" -->
<%
dim arr(8)
arr(0)=""
arr(1)=""
arr(2)=""
arr(3)=""
arr(4)=""
arr(5)=""
arr(6)=""
arr(7)=""

dim upload,file,formName,formPath,iCount,filename,fileExt,i
set upload=new upload_5xSoft ''�����ϴ�����

sname=upload.form("sname")
memo=upload.form("memo")

formPath=upload.form("filepath")
''��Ŀ¼���(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
i=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
    set file=upload.file(formName)  ''����һ���ļ�����

    if file.filesize<1 then exit for
   
    if file.filesize>153600 then
        response.write "<font size=2>ͼƬ��С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
        response.end
    end if

    fileExt=lcase(right(file.filename,4))
    filename=formPath&file.FileName
    arr(i)=file.FileName
    i=i+1
	
if fileEXT<>".gif" and fileEXT<>".jpg" then
    response.write "<font size=2>�ļ���ʽ���ԡ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
    response.end
end if 

    if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
        file.SaveAs Server.mappath(filename)   ''�����ļ�
        ' response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" �ɹ�!<br>"

          iCount=iCount+1
    end if
    set file=nothing
next

set upload=nothing  ''ɾ���˶���

for i=0 to UBound(arr)
if arr(i)="" then arr(i)=" "
next

sql = "insert into dbm_subs (sname,sfile1,sfile2,sfile3,sfile4,sfile5,sfile6,sfile7,sfile8,memo,inserttime) values ('"& sname &"','"&arr(0)&"','"&arr(1)&"','"&arr(2)&"','"&arr(3)&"','"&arr(4)&"','"&arr(5)&"','"&arr(6)&"','"&arr(7)&"','"&memo&"','"&now()&"')"
conn.Execute(sql)

'response.write "<html><head><meta  http-equiv='Refresh' content='3 url=""javascript:window.close();""'></head><body><center><br><br>�ļ��ϴ��ɹ�<br><br>"&arr(0)&"<br>"&arr(1)&"<br><br>"&filedata&"<br><br>"&sname&"<br>"&"<br>лл���֧�֣�<br>������������Զ��ر�</center></body></html>"
response.write "<html><head><meta  http-equiv='Refresh' content='3 url=""javascript:history.go( -1 );""'></head><body><center><br><br>�ļ��ϴ��ɹ�<br><br>"&arr(0)&"<br>"&arr(1)&"<br>"&arr(2)&"<br>"&arr(3)&"<br>"&arr(4)&"<br>"&arr(5)&"<br>"&arr(6)&"<br>"&arr(7)&"<br>"&filedata&"<br><br>"&sname&"<br>"&"<br>лл���֧�֣�<br>������������Զ�����</center></body></html>"
response.end
%>
