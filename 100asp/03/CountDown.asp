<%@ Language=VBScript %>
<html>
<title>
����ʱ��
</title>
<body>
<%
response.write "������"
response.write formatDateTime(Date(),1) & "," 
'��ʽ��Ϊ�����ڸ�ʽ�����ʾ
response.write " ��߿�����"
response.write "<font color=blue><u>"
'����DateDiff����,�������ڼ��.
response.write DateDiff("d",Date(),"01-07-07")
response.write "</font></u>"
response.write "��"
%>
</body>
</html>
