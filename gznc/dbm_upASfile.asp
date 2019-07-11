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
set upload=new upload_5xSoft ''建立上传对象

sname=upload.form("sname")
memo=upload.form("memo")
subs_id=upload.form("subs_id")

formPath=upload.form("filepath")
''在目录后加(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
i=0
for each formName in upload.file ''列出所有上传了的文件
    set file=upload.file(formName)  ''生成一个文件对象 

    if file.filesize>0 then
   
    if file.filesize>153600 then
        response.write "<font size=2>图片大小超过了限制　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
        response.end
    end if

    fileExt=lcase(right(file.filename,4))
    filename=formPath&file.FileName
    arr(i)=file.FileName
    i=i+1
	
if fileEXT<>".gif" and fileEXT<>".jpg" then
    response.write "<font size=2>文件格式不对　[ <a href=# onclick=history.go(-1)>重新上传</a> ]</font>"
    response.end
end if 

    if file.FileSize>0 then         ''如果 FileSize > 0 说明有文件数据
        file.SaveAs Server.mappath(filename)   ''保存文件
        ' response.write file.FilePath&file.FileName&" ("&file.FileSize&") => "&formPath&File.FileName&" 成功!<br>"

          iCount=iCount+1
    end if
    set file=nothing
	
	else
	i=i+1
	end if
	
next

set upload=nothing

sql="select * from dbm_subs where subs_id=" & subs_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1

if arr(0)="" then arr(0)=rs("sfile1")
if arr(0)="" then arr(0)=" "
if arr(1)="" then arr(1)=rs("sfile2")
if arr(1)="" then arr(1)=" "
if arr(2)="" then arr(2)=rs("sfile3")
if arr(2)="" then arr(2)=" "
if arr(3)="" then arr(3)=rs("sfile4")
if arr(3)="" then arr(3)=" "
if arr(4)="" then arr(4)=rs("sfile5")
if arr(4)="" then arr(4)=" "
if arr(5)="" then arr(5)=rs("sfile6")
if arr(5)="" then arr(5)=" "
if arr(6)="" then arr(6)=rs("sfile7")
if arr(6)="" then arr(6)=" "
if arr(7)="" then arr(7)=rs("sfile8")
if arr(7)="" then arr(7)=" "

rs.close
set rs=nothing


sql = "update dbm_subs set sname='"&sname&"',sfile1='"&arr(0)&"',sfile2='"&arr(1)&"',sfile3='"&arr(2)&"',sfile4='"&arr(3)&"',sfile5='"&arr(4)&"',sfile6='"&arr(5)&"',sfile7='"&arr(6)&"',sfile8='"&arr(7)&"',memo='"&memo&"' where subs_id=" & subs_id
conn.Execute(sql)

'response.write "<html><head><meta  http-equiv='Refresh' content='3 url=""javascript:window.close();""'></head><body><center><br><br>文件上传成功<br><br>"&arr(0)&"<br>"&arr(1)&"<br><br>"&filedata&"<br><br>"&sname&"<br>"&"<br>谢谢你的支持！<br>本窗口三秒后自动关闭</center></body></html>"
response.write "<html><head><meta  http-equiv='Refresh' content='3 url=""dbm_subs_view.asp""'></head><body><center><br><br>文件上传成功<br><br>"&arr(0)&"<br>"&arr(1)&"<br><br>"&filedata&"<br><br>"&sname&"<br>"&"<br>谢谢你的支持！<br>本窗口三秒后自动返回</center></body></html>"
response.end
%>
