<Html>
<Head>
<Title>
��ʾ��ĳһҳͣ����ʱ��
</Title>
</Head>
<Body>
<%
If Request.QueryString("time")="" then
%>
�㻹δ�������������ӡ�<BR>
<%
Else
%>
������ҳͣ����<%=DateDiff("s",Request.QueryString("time"),Now())%>
�롣
<BR>
<%End if%>
<Br>
<A href=onlinetime.asp?time=<%=server.URLEncode(Now())%>>
������һҳͣ���˶����?</A><BR>
<BR>
</Body>
</Html>

