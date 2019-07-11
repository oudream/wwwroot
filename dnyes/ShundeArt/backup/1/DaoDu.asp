<!--#include file="function.asp"-->
<%
set rs=server.CreateObject("ADODB.RecordSet") 
if uselevel=1 then
	if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="typemaster" or Request.cookies(Forcast_SN)("key")="bigmaster" or Request.cookies(Forcast_SN)("key")="smallmaster" or Request.cookies(Forcast_SN)("key")="check" then
		rs.Source="select top 1 * from "& db_News_Table &" where checkked=1 and goodnews=1 and picname is not null order by NewsId DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="" then
		rs.Source="select top 1 * from "& db_News_Table &" where checkked=1 and goodnews=1 and picname is not null order by NewsId DESC"
	end if
	if Request.cookies(Forcast_SN)("key")="selfreg" then
		if Request.cookies(Forcast_SN)("reglevel")=3 then
			rs.Source="select top 1 * from "& db_News_Table &" a1,type a2 where a1.typeid=a2.typeid and checkked=1 and goodnews=1 and typeview=1  and picname is not null order by NewsId DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=2 then
			rs.Source="select top 1 * from "& db_News_Table &" a1,type a2 where a1.typeid=a2.typeid and checkked=1 and goodnews=1 and typeview=1  and picname is not null order by NewsId DESC"
		end if
		if Request.cookies(Forcast_SN)("reglevel")=1 then
			rs.Source="select top 1 * from "& db_News_Table &" a1,type a2 where a1.typeid=a2.typeid and checkked=1 and goodnews=1 and typeview=1  and picname is not null order by NewsId DESC"
		end if
	end if
else
	rs.Source="select top 1 * from "& db_News_Table &" a1,type a2 where a1.typeid=a2.typeid and checkked=1 and goodnews=1 and typeview=1 and picname is not null order by NewsId DESC"
end if
rs.Open rs.Source,conn,3,3
if rs.EOF then
	Response.Write "<center>目前暂无推荐内容</center>"
else
	title=rs("title")
	fileExt=lcase(getFileExtName(rs("picname")))
	%>
<table width="450" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="52" background="pic/hotnews_b.gif">&nbsp;</td>
  </tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="450" id="AutoNumber3">
	<tr> 
		<td width="263" align="left" valign="top" bgcolor="#FFFFFF">
			<a class=class target=_blank href="ReadNews.asp?NewsId=<%=rs("newsid")%>" title="<%=htmlencode4(title)%>"> 
			<!-- 推荐无图 -->
			<%if fileext="" then%>
				<TABLE cellSpacing=0 cellPadding=0 width=263 border=0>
					<TBODY>
						<TR>
							<TD vAlign=top width=8><IMG height=7 src="images/index_t_center_1.gif" width=8></TD>
							<TD width=248 bgColor=#c9d5ae></TD>
							<TD vAlign=top width=7><IMG height=7 src="images/index_t_center_2.gif" width=7></TD>
						</TR>
						<TR>
							<TD bgColor=#c9d5ae>&nbsp;</TD>
							<TD>
								<TABLE class=56 width="100%">
									<TBODY>
										<TR>
											<TD width="240" height="180" align="right" valign="bottom" background="images/nophoto.jpg" class=newstop></a><strong>推荐新闻无图</strong>&nbsp;<br>
											</TD>
										</TR>
										<TR>
											<TD align=right></TD>
										</TR>
									</TBODY>
								</TABLE>
							</TD>
							<TD bgColor=#c9d5ae>&nbsp;</TD>
						</TR>
						<TR>
							<TD><IMG height=8 src="images/index_t_center_4.gif" width=8></TD>
							<TD width=248 bgColor=#c9d5ae></TD>
							<TD><IMG height=8 src="images/index_t_center_3.gif" width=7></TD>
						</TR>
					</TBODY>
				</TABLE>
        
			<%else%>
				<!-- 推荐无图 -->      
				<%if fileext="jpg" or fileext="bmp" or fileext="png" or fileext="gif" then%>
					<TABLE cellSpacing=0 cellPadding=0 width=263 border=0>
						<TBODY>
							<TR>
								<TD vAlign=top width=8><IMG height=7 src="images/index_t_center_1.gif" width=8></TD>
								<TD width=248 bgColor=#c9d5ae></TD>
								<TD vAlign=top width=7><IMG height=7 src="images/index_t_center_2.gif" width=7></TD>
							</TR>
							<TR>
								<TD bgColor=#c9d5ae>&nbsp;</TD>
								<TD>
									<TABLE class=56 width="100%">
										<TBODY>
											<TR>
												<TD width="240" height="180" align="right" valign="bottom" background="<%=FileUploadPath & rs("picname")%>" class=newstop></a><strong>SHUNDE ART</strong>&nbsp;<br>
												</TD>
											</TR>
											<TR>
												<TD align=right></TD>
											</TR>
										</TBODY>
									</TABLE>
								</TD>
								<TD bgColor=#c9d5ae>&nbsp;</TD>
							</TR>
							<TR>
								<TD><IMG height=8 src="images/index_t_center_4.gif" width=8></TD>
								<TD width=248 bgColor=#c9d5ae></TD>
								<TD><IMG height=8 src="images/index_t_center_3.gif" width=7></TD>
							</TR>
						</TBODY>
					</TABLE>
      
				<%end if%>
			<%end if%>
			<%if fileext="swf" then%>
				<TABLE cellSpacing=0 cellPadding=0 width=263 border=0>
					<TBODY>
						<TR>
							<TD vAlign=top width=8><IMG height=7 src="images/index_t_center_1.gif" width=8></TD>
							<TD width=248 bgColor=#c9d5ae></TD>
							<TD vAlign=top width=7><IMG height=7 src="images/index_t_center_2.gif" width=7></TD>
						</TR>
						<TR>
							<TD bgColor=#c9d5ae>&nbsp;</TD>
							<TD>
								<TABLE class=56 width="100%">
									<TBODY>
										<TR>
											<TD width="240" height="180" a class=newstop >
												<CENTER>
												<object classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0' width="240" height="180">
												<param name=movie value="<%=FileUploadPath & rs("picname")%>">
												<param name=quality value=high>
												<param name='Play' value='-1'>
												<param name='Loop' value='0'>
												<param name='Menu' value='-1'>
												<param name="wmode" value="transparent">
												<embed src="<%=FileUploadPath & rs("picname")%>" width="240" height="180" pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed>
												</object>
												</CENTER>
											</TD>
										</TR>
										<TR>
											<TD align=right></TD>
										</TR>
									</TBODY>
								</TABLE>
							</TD>
							<TD bgColor=#c9d5ae>&nbsp;</TD>
						</TR>
						<TR>
							<TD><IMG height=8 src="images/index_t_center_4.gif" width=8></TD>
							<TD width=248 bgColor=#c9d5ae></TD>
							<TD><IMG height=8 src="images/index_t_center_3.gif" width=7></TD>
						</TR>
					</TBODY>
				</TABLE>
			<%end if%>
		</td>
		<td width=1  valign=top background="images/t.gif">&nbsp;</td>
		<td width=168  valign=top bgcolor="#FFFFFF" style="LEFT: 0px; WIDTH: 168px; WORD-WRAP: break-word">
			<strong>
				<font size=+1 face="楷体_GB2312" color="<%=rs("titlecolor")%>">
					<a title="<%=htmlencode4(title)%>"><%=htmlencode4(CutStr(title,24))%>
					</a>
				</font>
			</strong>
			<br>
			<br>
			<a class=class target=_blank href="ReadNews.asp?NewsId=<%=rs("newsid")%>" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs("Content")),220)%> </a>
			<table width=160 height="0" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="16" align="right">
						<a class=class target=_blank href="ReadNews.asp?newsid=<%=rs("newsid")%>" title="<%=htmlencode4(title)%>"> 
							<font color="#FF0000">[详细内容]</font>
						</a>
					</td>
				</tr>
			</table> 
		</td>
	</tr>
	<tr bgcolor="#FFFFFF" >
		<td colspan="3" background="images/t1.gif" align="left" valign="top"><img src="images/kb.gif" width="9" height="9"></td>
	</tr>
	<tr>
		<td colspan="3" align="left" valign="top" bgcolor="#FFFFFF">
			<table width="98%" border="0" cellspacing="0" cellpadding="0">
				<tr> 
					<td> 
						<% set rs9=server.CreateObject("ADODB.RecordSet") 
						rs9.Source="select top 6 * from "& db_News_Table &" where istop=1 order by newsid desc"
						rs9.Open rs9.Source,conn,1,1
						do while not rs9.EOF
							title=trim(rs9("title"))%>
							<img src="images/dot_blue.gif">
							<a class=middle target="_blank" href="ReadNews.asp?NewsId=<%=rs9("newsid")%>" title="<%=htmlencode4(title)%>">
							<font color=<%=rs9("titlecolor")%>>[<%=CutStr(htmlencode4(rs9("title")),10)%>]</font> 
							</a>
							<a class=middle target="_blank" href="ReadNews.asp?NewsId=<%=rs9("newsid")%>" title="<%=htmlencode4(title)%>"><%=CutStr(nohtml(rs9("Content")),44)%></a>
							
                          <!--标题后评论提示-->

						  <% if rs9("titlesize")>=1 then %>
						  <A class=middle HREF="<%=path%>Review.asp?NewsID=<%=rs9("NewsID")%>" target="_blank" >
						  <font color=red><b>评</b></font></A>
						  <%end if %>

						  <!--标题后评论提示-->
						  <br>

							<%rs9.movenext
						loop
						rs9.close
						set rs9=nothing%>
						<BR>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if
rs.close
set rs=nothing
%>
