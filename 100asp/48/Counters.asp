<html>
<head>
<title>��վ������ͳ��
</title>
</head>
<body bgcolor=f0f0fc>
<%
set Counter=server.createobject("MSWC.Counters")
counter.Set "boat",31
counter.set "fly",433
counter.set "fang",5321
%>
<h2>������ˮ֮�۵���վ</h2>
����<%= Counter.Increment("boat") %>���ʱ���վ
<h2>��������������վ</h2>
����<%= Counter.Increment("fly") %>���ʱ���վ
<h2>�����ҵ���վ</h2>
����<%= Counter.Increment("fang") %>���ʱ���վ
</body>
</html>