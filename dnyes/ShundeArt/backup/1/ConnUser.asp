<%
'DbUserType = "MSSQL" '����MSSQL���ݿ�
'DbUserType = "MYSQL" '����MYSQL���ݿ�
DbUserType = "ACCESS" '����ACCESS���ݿ�

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
	Response.Write "���ݿ����ӳ���[���룺02]���������ݿ������ļ��е������ִ���"
	Response.End
End If
%>