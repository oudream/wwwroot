<!--#include file="conn.asp" -->
<%
dim LASTIP,NEWIP
set rs=server.CreateObject("ADODB.RecordSet")
Rs.Open "Select * From counters" ,conn,1,3

LASTIP = RS("LASTIP")
NEWIP = REQUEST.servervariables("REMOTE_ADDR")

IF CSTR(Month(RS("DATE"))) <> CSTR(Month(DATE())) THEN
	RS("DATE") = DATE()
	RS("YESTERDAY") = RS("TODAY")
	RS("BMONTH") = RS("MONTH")
	RS("MONTH") = 0
	RS("TODAY") = 0
	RS("TOTAL")  =  RS("TOTAL") + 1
	RS("TODAY") =  RS("TODAY") + 1
	RS("MONTH")  =  RS("MONTH") + 1
	RS("LASTIP") = NEWIP
	RS.Update
	session("IsFirst")=true
ELSE
	IF CSTR(Day(RS("DATE"))) <> CSTR(Day(DATE())) THEN
		RS("DATE") = DATE()
		RS("YESTERDAY") = RS("TODAY")
		RS("TODAY") = 0
		RS("TOTAL")  =  RS("TOTAL") + 1
		RS("TODAY") =  RS("TODAY") + 1
		RS("MONTH")  =  RS("MONTH") + 1
		RS("LASTIP") = NEWIP
		RS.Update
		session("IsFirst")=true
	END IF
END IF
if session("IsFirst")="" then		'判断是否是第一次登陆，如果是则计数器加一
	RS("TOTAL")  =  RS("TOTAL") + 1
	RS("TODAY") =  RS("TODAY") + 1
	RS("MONTH")  =  RS("MONTH") + 1
	RS("LASTIP") = NEWIP
	RS.Update
	session("IsFirst")=true
end if
%>
document.write('&nbsp;&nbsp;○- 昨日访问：<font color=red><%=rs("YESTERDAY")%></font><br>');
document.write('&nbsp;&nbsp;○- 今日访问：<font color=red><%=rs("TODAY")%></font><br>');
document.write('&nbsp;&nbsp;○- 上月访问：<font color=red><%=rs("BMONTH")%></font><br>');
document.write('&nbsp;&nbsp;○- 本月访问：<font color=red><%=rs("MONTH")%></font><br>');
document.write('&nbsp;&nbsp;○- 访问总数：<font color=red><%=rs("TOTAL")%></font> '); 
<%rs.Close
set rs=nothing%>