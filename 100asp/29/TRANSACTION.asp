<%@ TRANSACTION=Required %>

<%
'��������Ա���ʾ��ͬ��ҳ��
Response.Buffer = True
%>

<HTML>
<TITLE>��������</TITLE>
<BODY>
<H1>��ӭʹ���������з���</H1>

<P>лл�����ڴ�����������</P>
</BODY>
</HTML>

<%
'�������ɹ�����ʾ��ҳ��
Sub OnTransactionCommit()
%>
<HTML>
<BODY>

лл�������ʺ��ѻ�����Ρ�

</BODY>
</HTML>

<%
Response.Flush()
End Sub
%>

<%
'�������ʧ������ʾ��ҳ��
Sub OnTransactionAbort()
Response.Clear()
%>
<HTML>
<BODY>

�޷������������

</BODY>
</HTML>
<%
Response.Flush()
End Sub
%>
