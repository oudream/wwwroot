<Html>
<Head>
<Title>
显示在某一页停留的时间
</Title>
</Head>
<Body>
<%
If Request.QueryString("time")="" then
%>
你还未点击过下面的链接。<BR>
<%
Else
%>
你在上页停留了<%=DateDiff("s",Request.QueryString("time"),Now())%>
秒。
<BR>
<%End if%>
<Br>
<A href=onlinetime.asp?time=<%=server.URLEncode(Now())%>>
我在这一页停留了多久了?</A><BR>
<BR>
</Body>
</Html>

