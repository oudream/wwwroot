<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>���ڵ�¼</title>
<script language=JavaScript>
function agree(ischecked){
if(ischecked) bbs.innerHTML="<input type='submit' value='�ֶ���¼��̳��֤'>";
else bbs.innerHTML="<input type='submit' value='�Զ��ύ��̳��֤' disabled>";
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
	<input type="submit" name="buttomto" value="�Զ��ύ��̳��֤" >
</SPAN>
<SPAN class=style1> 
	<input onclick=agree(this.checked) type=checkbox value=checkbox name=checkbox>�ֶ�</SPAN>
</form>
</td></tr></table>
<p align=center>��ȴ�ҳ��ת����δת�򵽵�¼ǰ��ҳ�棬����  <a href="<%=ViewUrl%>">����</a>  ��</p>
</body>
</html>