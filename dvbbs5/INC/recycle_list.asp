<%
	sub AnnounceList1()
	
	sql="select AnnounceID,parentID,boardID,UserName,child,Topic,body,DateAndTime,hits,RootID,Expression,times,locktopic,istop,isbest,isvote from bbs1 where locktopic=2 order by announceid desc"
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		'论坛无内容
		call showEmptyBoard1()
		bBoardEmpty = true
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      		totalrec=rs.recordcount
            	call showpagelist1() 
	end if
	end sub


	REM 显示贴子列表	
	sub showPageList1()
	dim body
	dim vrs
	dim votenum,votenum_1
	dim pnum
	i=0
response.write "<form name=recycle action=admin_recycle.asp method=post><table cellspacing=0 border=0 width=95% bgcolor="&Forum_body(0)&" align=center><tr><td height=1></td></tr></table>"&_
				"<TABLE style=color:"&Forum_body(6)&"  border=1 cellPadding=0 cellSpacing=0 width=95% align=center bordercolor="&Forum_body(0)&">"&_
  				"<TBODY>"&_
				"<TR align=middle>"&_
				"<TD height=27 width=32 bgColor="&Forum_body(2)&"><font color="&Forum_body(6)&">状态</TD>"&_ 
				"<TD bgColor="&Forum_body(2)&" width=*><font color="&Forum_body(6)&">主 题</TD>"&_
				"<TD bgColor="&Forum_body(2)&" width=80><font color="&Forum_body(6)&">作 者</TD>"&_
				"<TD bgColor="&Forum_body(2)&" width=64><font color="&Forum_body(6)&">回复/人气</TD>"&_
				"<TD bgColor="&Forum_body(2)&" width=195><font color="&Forum_body(6)&">最后更新 | 回复人</TD>"&_
				"</TR>"&_ 
				"</TBODY></TABLE>"
		while (not rs.eof) and (not page_count = rs.PageSize)
		

response.write "<TABLE style=color:"&Forum_body(7)&" border=1 cellPadding=0 cellSpacing=0 width=95% align=center bordercolor="&Forum_body(0)&">"&_
				"<TBODY><TR align=middle>"&_
				"<TD bgColor="&Forum_body(5)&" width=32 height=27>"

response.write "<input type=checkbox name=topicid value="&rs(0)&">"

response.write "</TD>"&_
				"<TD align=left bgcolor="&Forum_body(4)&" width=* onmouseover=javascript:this.bgColor='"&Forum_body(5)&"' onmouseout=javascript:this.bgColor='"&Forum_body(4)&"'>"
	response.write "<img src='face/"&rs("Expression")&"' width=15 height=15>"

	body=replace(replace(htmlencode(left(rs(6),20)),"<BR>",""),"</P><P>","")
	
response.write "<a href=""javascript:openScript('viewrecycle.asp?id="&rs(0)&"',500,400)"">"
if rs(5)="" or isnull(rs(5)) then
	response.write body
else
	if len(rs(5))>26 then
	response.write ""&left(htmlencode(rs(5)),26)&"..."
	else
	response.write htmlencode(rs(5))
	end if
end if
	response.write "</a>"

response.write "</TD>"&_
				"<TD bgColor="&Forum_body(5)&" width=80><a href=dispuser.asp?name="& rs(3) &">"& rs(3) &"</a></TD>"&_
				"<TD bgColor="&Forum_body(4)&" width=64>"
if rs(15)=1 then
	set vrs=conn.execute("select votenum from vote where announceid="& rs(0) &"")
	if not(vrs.eof and vrs.bof) then
	votenum=vrs("votenum")
	votenum=split(votenum,"|")
	dim iu
	for iu = 0 to ubound(votenum)
		votenum_1=cint(votenum_1)+votenum(iu)
	next
	response.write "<FONT color=#990000><b>"&votenum_1&"</b></font>  票"
	votenum_1=0
	end if
	vrs.close
	set vrs=nothing
else
	response.write "<font color="& Forum_body(7) &">"& rs(4) &"/"& rs(8) &"</font>"
end if

response.write "</TD><TD align=left bgColor="& Forum_body(5) &" width=195>&nbsp;"

	'on error resume next
		response.write "&nbsp;"&_
						FormatDateTime(rs(7),2)&"&nbsp;"&FormatDateTime(rs(7),4)&_
						"&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;------"

response.write "</TD></TR>"&_
				"</TBODY></TABLE>"

	  page_count = page_count + 1
          rs.movenext
        wend
	if err.number<>0 then err.clear
	end sub


	sub listPages3()
	dim endpage
	'on error resume next
	if master then
response.write "<table cellspacing=0 border=0 width=95% bgcolor="&Forum_body(0)&" align=center><tr><td height=1></td></tr></table>"&_
				"<TABLE style=color:"&Forum_body(6)&"  border=1 cellPadding=0 cellSpacing=0 width=95% align=center bordercolor="&Forum_body(0)&">"&_
  				"<TBODY>"&_
				"<TR>"&_
				"<TD height=27 width=""100%"" bgColor="&Forum_body(2)&"><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">选中所有显示帖子&nbsp;<input type=submit name=action onclick=""{if(confirm('确定还原选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=还原>"
	if instr(session("flag"),"06")>0 then
	response.write "&nbsp;<input type=submit name=action onclick=""{if(confirm('确定删除选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=删除>&nbsp;<input type=submit name=action onclick=""{if(confirm('确定清除回收站所有的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=清空回收站>"
	end if
	response.write "</TD>"&_ 
				"</TR></form>"&_ 
				"</TBODY></TABLE>"
	end if
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<strong>"&Forum_Setting(11)&"</b> 贴数<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： <b>"

	if currentpage > 4 then
	response.write "<a href=""?page=1"">[1]</a> ..."
	end if
	if Pcount>currentpage+3 then
	endpage=currentpage+3
	else
	endpage=Pcount
	end if
	for i=currentpage-3 to endpage
	if not i<1 then
		if i = clng(currentpage) then
        response.write " <font color="&Forum_body(8)&">["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a></b>"
	end if
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing
	end sub 

	sub showEmptyBoard1()
		response.write "<TABLE style=color:"&Forum_body(6)&" bgColor='"&Forum_body(0)&"' border=0 cellPadding=4 cellSpacing=1 width=95% align=center>"&_
			"<TBODY>"&_
			"<TR align=middle bgColor='"&Forum_body(2)&"'>"&_
			"<TD height=25><font color="&Forum_body(6)&">状态</font></TD>"&_
			"<TD><font color="&Forum_body(6)&">主 题  (点心情符为开新窗浏览)</TD>"&_
			"<TD><font color="&Forum_body(6)&">作 者</TD> "&_
			"<TD><font color="&Forum_body(6)&">回复/人气</TD> "&_
			"<TD><font color="&Forum_body(6)&">最新回复</TD></TR> "&_
			"<tr bgColor="&Forum_body(4)&"><td colSpan=5 width=100% >论坛回收站暂无内容。</td></tr>"&_
			"</TBODY></TABLE>"
	end sub
%>