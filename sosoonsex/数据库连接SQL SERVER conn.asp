<%
'全测试通过的

'dim conn
'dim sqlstr
'set conn =server.CreateObject("adodb.connection")
'sqlstr="driver={sql server};server=192.168.1.193;uid=sa;pwd=123456;database=sosoonsex;"
'conn.open sqlstr

dbuid="sa"    			
dbpwd="123456"     			
dbname="sosoonsex"  		
set conn=Server.CreateObject("adodb.Connection")
Conn.Open  "PROVIDER=SQLOLEDB.1;Initial Catalog="&dbname&";User ID="&dbuid&";Password="&dbpwd
'Conn.Open  "PROVIDER=SQLOLEDB.1;DSN=sosoonsex;Initial Catalog="&dbname&";Persist Security Info=True;User ID="&dbuid&";Password="&dbpwd&";Connect Timeout=30"
%>
