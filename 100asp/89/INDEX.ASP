<script LANGUAGE="javascript">
<!--
function research_onsubmit() {
 if(document.research.Options[0].checked==false &&document.research.Options[1].checked==false&&document.research.Options[2].checked==false&&document.research.Options[3].checked==false)
{
alert("错误：没有选择！")
return false
 }
 }
//-->
</script>
</head>
<body>
<P align=center><FONT face=华文行楷 size=5><STRONG>实例   网上测试</STRONG></FONT></P>
<%
'随机数的产生
dim i
randomize
i=round(rnd(1)*15+1)

'定义题号
if session("countnum")="" then
	session("countnum")=1
else
	session("countnum")=session("countnum")+1
end if
if session("countnum")<=5 then
	'数据的读取
	set rs=server.createobject("adodb.recordset") 
	conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
	sql="select * from problem where id="+cstr(i)
	rs.open sql,conn,1,1
	if rs.EOF or rs.BOF then
		Response.Write "数据库中暂无资料!"
	else
	'数据的显示%>

		<table align=center width=450>
		<TR><TD><STRONG>每轮考试为 5 道题!</STRONG><BR></td></tr>
		<tr><TD><p><STRONG>第<%Response.Write session("countnum")%>题&nbsp;&nbsp;<%=rs("problem")%></STRONG><p></td></tr>
		<tr><td><STRONG>Choices：</STRONG><ul>
		<form method="post" action="check.asp?IID=<%=rs("id")%>" name="research" LANGUAGE="javascript" onsubmit="return research_onsubmit()">
		 <input type="radio" value="A" name="Options">A：&nbsp;&nbsp;<%=rs("selectA")%><BR>
		 <input type="radio" value="B" name="Options">B：&nbsp;&nbsp;<%=rs("selectB")%><BR>
		 <input type="radio" value="C" name="Options">C：&nbsp;&nbsp;<%=rs("selectC")%><BR>
		 <input type="radio" value="D" name="Options">D：&nbsp;&nbsp;<%=rs("selectD")%></ul></td></tr>
		 <tr><td align="center"><input type="submit" value=" 提交 " name="cmdok" class="buttonface"></TD></tr></FORM></table>
	<%
	end if
	set rs=nothing 
	set conn=nothing
else
	Response.Write "<P align=center><STRONG>此轮考试已经结束<BR></STRONG>"
	msg="您的考试成绩是"&session("acceptnum")*20&"分"
	session("countnum")=""
	session("acceptnum")=""
	Response.Write msg&"</p>"
	Response.End
end if
%>
</body>