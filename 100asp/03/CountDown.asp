<%@ Language=VBScript %>
<html>
<title>
倒记时间
</title>
<body>
<%
response.write "今天是"
response.write formatDateTime(Date(),1) & "," 
'格式化为长日期格式输出显示
response.write " 离高考还有"
response.write "<font color=blue><u>"
'调用DateDiff函数,计算日期间隔.
response.write DateDiff("d",Date(),"01-07-07")
response.write "</font></u>"
response.write "天"
%>
</body>
</html>
