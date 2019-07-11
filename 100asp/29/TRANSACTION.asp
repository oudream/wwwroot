<%@ TRANSACTION=Required %>

<%
'缓存输出以便显示不同的页。
Response.Buffer = True
%>

<HTML>
<TITLE>联机银行</TITLE>
<BODY>
<H1>欢迎使用联机银行服务</H1>

<P>谢谢。正在处理您的事务。</P>
</BODY>
</HTML>

<%
'如果事务成功则显示该页。
Sub OnTransactionCommit()
%>
<HTML>
<BODY>

谢谢。您的帐号已获得信任。

</BODY>
</HTML>

<%
Response.Flush()
End Sub
%>

<%
'如果事务失败则显示该页。
Sub OnTransactionAbort()
Response.Clear()
%>
<HTML>
<BODY>

无法完成您的事务。

</BODY>
</HTML>
<%
Response.Flush()
End Sub
%>
