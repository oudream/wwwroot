<%@ Language=VBScript %>
<html>
<title>
��վ��ҳ
</title>
<body>
<%
'��������Զ��������ַ�������ж�,���Ϊ����
'��ַ����뻶ӭҳ��,������ʾ������Ϣ.
dim address
address = request.servervariables("REMOTE_ADDR")
if address="127.0.0.1" then 
    '���Ϊ����,����ʾ��ӭҳ��.
    response.write "���,��ӭ���뱾վ��."
else
    '������ʾ������Ϣ.
    response.write "�Բ���,����Ȩ�鿴�ڲ�վ��."
end if
%>
</body>
</html>