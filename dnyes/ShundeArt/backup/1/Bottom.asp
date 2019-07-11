<style type="text/css">
<!--
.style1 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
	color: #B9CFF3;
}
.style2 {color: #000000}
.style4 {
	color: #000000;
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style>
<div align="center"> 
<TABLE WIDTH=760 BORDER=0 CELLPADDING=0 CELLSPACING=0>
<TR> 
	<TD bgcolor="#7C7CB5"> 
		<div align="center">
			<p><br>
				<span class="style4">Copyright</span>
				<span class="style1"> &copy;</span>
				<span class="style4">2002-2004</span>
				<span class="style1">FORECAST <span class="style2">NEWS</span> </span>
				<a href="http://feitium.yeah.net" target="_blank" class=bottom>[沸腾展望新闻多媒体系统]</a>&nbsp;<font color="#B9CFF3">&nbsp;<%=version%>&nbsp;<%=ver%></font>&nbsp;<span class="style4">All Rights Reserved</span></p>
				<p>
<%
dim mymenu1
mymenu1=BOTTOMmenu
Response.Write mymenu1
%>
				</P>
				<p>新闻系统核心:&nbsp;<a href="http://www.5share.com" title="尘缘雅境图文系统官方论坛" target="_blank" class=bottom>尘缘雅境</a>&nbsp; 制作：<a class=bottom>杨正炎(雪域一线天)</a> 
					<a href="login.asp" target="_blank" class=bottom>[后台管理]</a><br>
<%dim endtime
endtime=timer()
response.write "页面执行时间："&FormatNumber((endtime-starttime)*1000,3)&"毫秒"%>
			</p><br>
		</div>
	</TD>
</TR>
<TR> 
	<TD>
		<IMG SRC="images/news2003-2d_07.gif" WIDTH=760 HEIGHT=31 ALT="">
	</TD>
</TR>
</TABLE>
</div>