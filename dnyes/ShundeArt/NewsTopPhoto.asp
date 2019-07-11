<!--最新图文代码开始-->

<%
set rs3=server.CreateObject("ADODB.RecordSet")
if uselevel=1 then
	if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
		rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="" then
		rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and newslevel=0  and picname<>'null' order by NewsID DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="selfreg" then
		if Request.cookies(Forcast_SN)("reglevel")=3 then
			rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=2 then
			rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=1 then
			rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
		end if
	end if
else
	rs3.Source ="select top " & top_img & " * from "& db_News_Table &" where picnews=1 and checkked=1 and picname<>'null' order by NewsID DESC"
end if
rs3.Open rs3.Source,conn,3,3
if not rs3.EOF then%>
<div align='center' id='demo' style='overflow:hidden;height:125px;width:590px;'><!--滚动区的高度和宽度-->
<table align='center' cellpadding='0' cellspace='0' border='0'>
<tr>
	<td id='demo1' valign='top'>
		<table width='100%' cellpadding='0' cellspacing='0' border='0' align='center'>
		<tr valign='top'>
	<%
	while not rs3.EOF
		fileExt=lcase(getFileExtName(rs3("picname")))
		%>
			<td align='center'>
				<TABLE width=100% border=0 align=center cellPadding=0 cellSpacing=0>
				<TR>
					<TD width=8 rowspan=3 >&nbsp;</TD>
					<TD vAlign=top width=8><img src='Images/bg_0ltop.gif' width=10 height=10></TD>
					<TD background=images/bg_01.gif></TD>
					<TD vAlign=top width=7><img src='Images/bg_0rtop.gif' width=10 height=10></TD>
					<TD width=7 rowspan=3 vAlign=top>&nbsp;</TD>
				</TR>
				<TR>
					<TD background='Images/bg_03.gif'>&nbsp;</TD>
					<TD align="center" bgcolor="#E9E9E9">
						<a class=middle href='ReadNews.asp?NewsID=<%=rs3("NewsID")%>' target=_blank title='<%=rs3("title")%>'>
		<%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
							<img  src='<%=FileUploadPath & rs3("picname")%>' width='105' height='80' border=1 style=border-color:black >
		<%else
			if fileext="swf" then%>
							<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width='105' height='80'>
								<param name=movie value='<%=FileUploadPath & rs3("picname")%>'>
								<param name=quality value=high>
								<param name='Play' value='-1'>
								<param name='Loop' value='0'>
								<param name='Menu' value='-1'>
								<param name='wmode' value='transparent'>
								<embed src='<%=FileUploadPath & rs3("picname")%>' width='105' height='80' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
							</object>
						</a>
						<a class=middle href='ReadNews.asp?NewsID=<%=rs3("NewsID")%>' target=_blank title='<%=rs3("title")%>'>
			<%end if
		end if%>
						</a>
	  			</TD>
	  			<TD background='Images/bg_04.gif'>&nbsp;</TD>
				</TR>
				<TR>
					<TD><img src='Images/bg_0lbottom.gif' width=10 height=10></TD>
					<TD><img src='Images/bg_02.gif'></TD>
					<TD><img src='Images/bg_0rbottom.gif' width=10 height=10></TD>
	  		</TR>
	  		<TR>
	  			<TD>&nbsp;</TD>
					<TD colspan=3 align=center height=20 valign='top' background='Images/bg_05.gif'>
	  				<a class=middle href='ReadNews.asp?NewsID=<%=rs3("NewsID")%>' target=_blank title='<%=rs3("title")%>'><%=CutStr(htmlencode4(rs3("title")),14)%></a>
	 				</TD>
	 				<TD>&nbsp;</TD>
				</TR>
	 			</TABLE>
			</td>
	  <%
	  rs3.MoveNext
	wend
	%>
		</tr>
		</table>
	</td>
	<td id=demo2 valign=top></td>
</tr>
</table>
</div>
	<%if top_img >4 then%>
<script>
var speed=15
demo2.innerHTML=demo1.innerHTML
function Marquee1(){
if(demo2.offsetWidth-demo.scrollLeft<=0)
demo.scrollLeft-=demo1.offsetWidth
else{
demo.scrollLeft++
}
}
var MyMar1=setInterval(Marquee1,speed)
demo.onmouseover=function() {clearInterval(MyMar1)}
demo.onmouseout=function() {MyMar1=setInterval(Marquee1,speed)}
</script>
	<%end if
else
	Response.Write "暂 无 最 新 图 文"
end if
rs3.close
set rs3=nothing
%>
<!--最新图文代码结束-->