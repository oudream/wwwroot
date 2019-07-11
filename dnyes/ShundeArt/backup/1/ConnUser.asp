<%
'DbUserType = "MSSQL" '链接MSSQL数据库
'DbUserType = "MYSQL" '链接MYSQL数据库
DbUserType = "ACCESS" '链接ACCESS数据库

if DbUserType="ACCESS" then
	UserDB = "data/news3000.mdb"
	ConnUserStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(UserDB)
end if

if DbUserType="MSSQL" then
	ConnUserStr = "driver={SQL Server};server=127.0.0.1;uid=feitium;pwd=12345678;database=FeiTium"
end if

if DbUserType="MYSQL" then
	ConnUserStr = "driver=MySQL;server=SERVER_NAME;uid=UID;pwd=PWD;database=DB_NAME"
end if

On Error Resume Next
Set ConnUser = Server.CreateObject("ADODB.Connection")
ConnUser.open ConnUserStr

If Err Then
	err.Clear
	Set ConnUser = Nothing
	Response.Write "数据库连接出错[代码：02]，请检查数据库链接文件中的连接字串。"
	Response.End
End If
%>