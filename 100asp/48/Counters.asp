<html>
<head>
<title>网站访问量统计
</title>
</head>
<body bgcolor=f0f0fc>
<%
set Counter=server.createobject("MSWC.Counters")
counter.Set "boat",31
counter.set "fly",433
counter.set "fang",5321
%>
<h2>这是在水之舟的网站</h2>
共有<%= Counter.Increment("boat") %>访问本网站
<h2>这是轻舞飞扬的网站</h2>
共有<%= Counter.Increment("fly") %>访问本网站
<h2>这是我的网站</h2>
共有<%= Counter.Increment("fang") %>访问本网站
</body>
</html>