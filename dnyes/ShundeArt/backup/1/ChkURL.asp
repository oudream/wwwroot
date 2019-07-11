<%'非法提交信息检测
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
	response.write "<META http-equiv=Content-Type content='text/html; charset=gb2312'>"
	response.write "<META content='MSHTML 6.00.3790.118' name=GENERATOR>"
	response.write "<head>"
	response.write "<LINK href='site.css' type=text/css rel=stylesheet>"
	response.write "<title>"& copyright & version & ver &"</title>"
	response.write "</head>"
	response.write "<body bgcolor='#CCCCCC' topmargin='0' marginheight='0'>"
	response.write "<table border=0 cellpadding=1 hei='150%' style='border-collapse: collapse' bordercolor='#C0C0C0' width='100%' id='AutoNumber1'>"
	response.write "<tr><td colspan=1 height='50' width='100%'>"
	response.write "</td></tr>"
	response.write "<tr class=TDtop><td height='25'>&nbsp;"
	response.write "<div align=center>┊ <strong><font color=red>访问错误</font></strong> ┊</div>"
	response.write "</td></tr></table>"
	response.write "<table border=1 cellpadding=1 hei='100%' style='border-collapse: collapse' bordercolor='#C0C0C0' width='100%' id='AutoNumber1'>"
	response.write "<tr><td height='150' bgcolor='ffb100'>"
	response.write "<center><font size=4 color='#ffffff'>请通过正确的链接进入本站！</font><br><br><br><br><a href='./' title='进入本站面页'><font size=2 color='#000000'>打开首页</font></a></center>"
	response.write "</td></tr></table>"
	response.write "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse: collapse' bordercolor='#C0C0C0' width='100%' id='AutoNumber1'>"
	response.write "<tr><td bgcolor='EAEAF4'><div align='center'>版权："& copyright &" 授权使用："& jjgn &"</div></td></tr>"
	response.write "<tr><td background='IMAGES/top2.gif'>&nbsp;</td></tr>"
	response.write "</table>"
	response.end
end if
%>