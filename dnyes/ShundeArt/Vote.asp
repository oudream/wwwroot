<% Response.Buffer=True %>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<%
if request.QueryString("stype")="" then
	if Request.ServerVariables("REMOTE_ADDR")=request.cookies(Forcast_SN)("IPAddress") then
		response.write"<SCRIPT language=JavaScript>alert('��л����֧�֣����Ѿ�Ͷ��Ʊ�ˣ������ظ�ͶƱ��лл��');"
		response.write"javascript:window.close();</SCRIPT>"
	else
		options=request.form("options")
		response.cookies(Forcast_SN)("IPAddress")=Request.ServerVariables("REMOTE_ADDR") 
		set rs=server.createobject("adodb.recordset")
		sql="update "& db_Vote_Table &" set answer"&options&"=answer"&options&"+1 where IsChecked=1"
		rs.open sql,conn,1,3
		set rs=nothing
	end if
end if
%>

<head> 
<title><%=redcaff%>ͶƱ���</title>
<LINK href=site.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file=User_Top.asp-->
<div align="center"></center> 
	<table border="0" cellpadding="0" cellspacing="0" width="500" height="48" bordercolor=<%=border%> style="border-collapse: collapse">
    		<br><br><br><br><%
		total=0
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Vote_Table &" where IsChecked=1"
		rs.open sql,conn,1,1
		%>
		<tr align="center"> 
			<td height="24" colspan="4" valign="top"><font color="#000000">��</font><font color="#000073"><%=rs("Title")%></font><font color="#000000">��ͶƱ���</font><br>
			</td>
		</tr>
		<tr>
			<td height="24" valign="top" colspan="4"><font color="#000000">=============================================================================</font><br>
			</td>
		</tr>
		<tr>
			<td width="22%" valign="top">���</td>
			<td valign="top">�ٱȷ�</td>
			<td colspan="2" align="center" valign="top">����</td>
		</tr>
		<%
		for i=1 to 8
			if rs("Select"&i)<>"" then
				total=total+rs("answer"&i)
			end if
		next
		for i=1 to 8
			if rs("Select"&i)<>"" then
				if total=0 then
					answer=0
				else
					answer=(rs("answer"&i)/total)*100
				end if%>
		<tr>
			<td valign="top"><%=i%>.<%=rs("select"&i)%>��</td>
			<td width="56%" valign="top"><img src=images/RSCount.gif width=<%=int(answer*2)%> height=8></td>
			<td width="12%" valign="top"><%=round(answer,3)%>%</td>
			<td width="10%" valign="top"><%=rs("answer"&i)%>��</td>
			<%end if
		next%>
		</tr>
		<tr>
			<td colspan="4"> ���С�<%=total%>���˲μ�ͶƱ<br>=============================================================================</font></center></td>
		</tr>
	</table>
</div>
<BR>
<div align="center">�ǳ���л����֧�����������²�����<BR><BR>��<a href="guestbook.asp" target="_blank">���½���</a>����<a href="javascript:window.close()">�رմ���</a>��<br><br><br><br>
	<% rs.close     
	set rs=nothing     
	conn.close     
	set conn=nothing %>
</div>
</body>
<!--#include file=User_Bottom.asp-->
