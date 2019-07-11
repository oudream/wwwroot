<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>正在登录</title>
<script language=JavaScript>
function agree(ischecked){
if(ischecked) bbs.innerHTML="<input type='submit' value='手动登录论坛验证'>";
else bbs.innerHTML="<input type='submit' value='自动提交论坛验证' disabled>";
}
</script>

<script languge=javascript>
function go()
	{
	document.bbslogin.buttomto.click();
	document.bbslogin.buttomto.disabled=true;
	//location.href="<%=ViewUrl%>"
	}
</script>
<%
response.write "<meta http-equiv=""refresh"" content="""&freetime&";url="& ViewUrl &""">"
response.write "<link rel='stylesheet' type='text/css' href='style.css'>"
response.write "<body onload=go()>"
response.write "<table width=253 border=0 bgcolor=#000000 cellspacing=0 cellpadding=0 align=center><tr><td align=center bgcolor=#FFFFFF><p>"
Response.Write "<iframe name='framename' width='760' height='340' frameborder=no scrolling=no src='othermessage.htm'></iframe>"%>
<form name="bbslogin" method="post" action="<%=BbsPath%>login.asp" target="framename">
<input type="hidden" name="action" value="chk">
<input type="hidden" name="username" value="<%=UserName1%>">
<SPAN id=bbspwd>
	<input type="hidden" name="password" value="<%=trim(request.form("passwd"))%>">
</SPAN>
<SPAN id=bbs>
	<input type="submit" name="buttomto" value="自动提交论坛验证" >
</SPAN>
<SPAN class=style1> 
	<input onclick=agree(this.checked) type=checkbox value=checkbox name=checkbox>手动</SPAN>
</form>
</td></tr></table>
<p align=center>请等待页面转向，如未转向到登录前的页面，请点击  <a href="<%=ViewUrl%>">这里</a>  。</p>
</body>
</html>