<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim totalrec
	dim n
	dim orders,ordername
	dim currentpage,page_count,Pcount
	currentPage=request("page")
	if not founduser then
		Errmsg="您还没有登陆，请<a href=login.asp>登陆</a>后浏览"
		founderr=true
	end if
	orders=request("s")
	if orders="" or not isInteger(orders) then
		orders=1
	else
		orders=clng(orders)
	end if
	if orders=1 then
	ordername="我参与的主题"
	sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" order by announceid desc) order by topicid desc"
	elseif orders=2 then
	sql="select top 200 topicID,boardID,postUserName,child,title,DateAndTime,hits,Expression,locktopic,istop,isbest,isvote from topic where postuserid="&userid&" and child>0 ORDER BY topicid desc"
	ordername="已被回复的我的发言"
	else
	ordername="我参与的主题"
	sql="select top 200 * from topic where topicid in (select top 200 rootid from "&NowUseBBS&" where postuserid="&userid&" order by announceid desc) order by topicid desc"
	end if	
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	stats=ordername
if founderr then
	call nav()
	call head_var(1)
else
	call nav()
	call head_var(1)
	call AnnounceList1()
	call listPages3()
end if
call activeonline()
call footer()

sub AnnounceList1()
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		call showEmptyBoard1()
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
	i=0
	response.write "<p align=center>当前论坛数据很可能进行了分表保存数据<BR>当前仅显示您在当前数据表中最前200帖，如果您需要查询更多数据请到搜索页面</p>"
	response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
	response.write "<TR align=middle>"
	response.write "<Th height=27 width=32 id=tabletitlelink>状态</th>"
	response.write "<Th width=*>主 题  (点<img src="&Forum_info(7)&"plus.gif>即可展开贴子列表)</Th>"
	response.write "<Th width=80>作 者</Th>"
	response.write "<Th width=64>回复/人气</Th>"
	response.write "<Th width=195>最后更新 | 回复人</Th>"
	response.write "</TR>"
	dim Ers, Eusername, Edateandtime,body
	dim vrs,votenum,pnum,p,iu,votenum_1
       while (not rs.eof) and (not page_count = rs.PageSize)
	boardid=rs("boardid")
response.write "<TR align=middle>"&_
				"<TD class=tablebody2 width=32 height=27>"

if rs("istop")<>1 and lockboard<>1 and rs("locktopic")<>1 and rs("isvote")<>1 and rs("isbest")<>1 and rs("child")<10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=开放主题>"
elseif rs("isvote")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(12)&""" alt=投票贴子>"
elseif rs("istop")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(3)&""" alt=固顶主题>"
elseif rs("isbest")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=精华帖子>"
elseif rs("child")>=10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(1)&""" alt=热门主题>"
elseif rs("locktopic")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=本主题已锁定>"
elseif lockboard=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=本论坛已锁定>"
else
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=开放主题>"
end if

response.write "</TD>"&_
				"<TD align=left  class=tablebody1 width=*>"

	response.write "<img src='"& Forum_info(7) &"nofollow.gif' id='followImg"& rs("topicid") &"'>"

response.write "<a href=""dispbbs.asp?boardID="& rs("boardid") &"&ID="& rs("topicid") &""">"

if len(rs("title"))>26 then
	response.write ""&left(htmlencode(rs("title")),26)&"..."
else
	response.write htmlencode(rs("title"))
end if
	response.write "</a>"
if (rs("child")+1)>cint(Forum_Setting(12)) then
	response.write "&nbsp;&nbsp;[分页："
	Pnum=(Cint(rs("child")+1)/cint(Forum_Setting(12)))+1
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><FONT color="&Forum_body(8)&"><b>"&p&"</b></font></a> "
	next
	response.write "]"
end if

response.write "</TD>"&_
				"<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs("postusername") &">"& rs("postusername") &"</a></TD>"&_
				"<TD class=tablebody1 width=64>"
	response.write ""& rs("child") &"/"& rs("hits") &""

response.write "</TD><TD align=left  class=tablebody2 width=195>&nbsp;"

	response.write "&nbsp;"&_
						FormatDateTime(rs("dateandtime"),2)&"&nbsp;"&FormatDateTime(rs("dateandtime"),4)&_
						"&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;------"

response.write "</TD></TR>"

	  page_count = page_count + 1
          rs.movenext
        wend
	response.write "</table>"
	end sub

	sub listPages3()
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"<font color="&Forum_body(13)&">页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<strong>"&Forum_Setting(11)&"</b> 贴数<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： "

	if currentpage > 4 then
	response.write "<a href=""?page=1&s="&orders&""">[1]</a> ..."
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
        response.write " <a href=""?page="&i&"&s="&orders&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&s="&orders&""">["&Pcount&"]</a>"
	end if
	response.write "</p></div></font></td></tr></table>"
	rs.close
	set rs=nothing	
	end sub 

	sub showEmptyBoard1()
%>
<TABLE cellPadding=4 cellSpacing=1 class=tableborder1 align=center>
<TBODY>
<TR align=middle>
<Th height=25>状态</Th>
<Th>主 题  (点心情符为开新窗浏览)</Th>
<Th>作 者</Th> 
<Th>回复/人气</Th> 
<Th>最新回复</Th>
</TR> 
<tr><td colSpan=5 width=100% class=tablebody1>暂无内容，欢迎发贴：）</td></tr>
</TBODY></TABLE>
<%
	end sub
%>