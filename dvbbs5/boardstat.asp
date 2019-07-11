<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: boardstat.asp
' Version:5.0
' Date: 2002-9-12
' Script Written by fssunwin,modify by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

stats="论坛状态图例"
dim bbsnums,topicnums,todaynums
dim allbbsnums,alltopicnums,alltodaynums
dim sinfo
dim allnum
dim td1,td2,td3

set rs=conn.execute("select TopicNum,BbsNum,TodayNum,UserNum from config where active=1")
allbbsnums=rs("bbsnum")
alltopicnums=rs("topicnum")
alltodaynums=rs("todaynum")
set rs=nothing

if boardid<>"" and boardid<>0 then
	bbsnums=lastbbsnum
	topicnums=LastTopicNum
	todaynums=todaynum
	boardtype=boardtype
else
	bbsnums=allbbsnums
	topicnums=alltopicnums
	todaynums=alltodaynums
	boardtype=Forum_info(0)
end if

td1="id=tdbg"
td2="id=tdbg"
td3="id=tdbg"


if request("action")="lastbbsnum" then
	sinfo="lastbbsnum"
	allnum=allbbsnums
	stats="总帖数"
	td1=""
elseif request("action")="lasttopicnum" then
	sinfo="lasttopicnum"
	allnum=alltopicnums
	stats="主题数"
	td2=""
elseif request("reaction")<>"" then
	if boardid=0 then
		stats="论坛总在线"
	else
		stats=BoardType & "在线"
	end if
else
	sinfo="todayNum"
	allnum=alltodaynums
	stats="今日帖数"
	td3=""
end if

call nav()
if boardid=0 then
	call head_var(1)
else
	call head_var(2)
end if

call stat_var()

if request("reaction")="onlineinfo" then
	call onlinemain()
elseif request("reaction")="onlineUserinfo" then
	call userinfo(BOARDID)
elseif request("reaction")="online" then
	call onlineinfo()
else
	call main(sinfo)
end if

call activeonline()
call footer()

sub main(SINFO)
%>
<table cellpadding=5 cellspacing=1 class=tableborder1 align=center>
<tr>
<th width="20%">论坛名称</th>
<th width="70%">图形比例</th>
<th width="10%"><%=stats%></th>
</tr>
<%
dim g,allrs
dim wid
g=1
set rs=server.createobject("adodb.recordset")
if boardmaster or master then
	sql="select "&SINFO&",boardid,BoardType from board order by "&SINFO&" desc"
else
	sql="select "&SINFO&",boardid,BoardType from board where lockboard=0 order by "&SINFO&" desc"
end if
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	response.write "<tr><td colspan=3>　　还没有任何数据。</td></tr>"
else  
	do while not rs.eof 
	g=g+1 
	if g=11 then g=1
	if allnum=0 then allnum=1
	if cint(rs(0)/allnum)>1 then
		allrs=conn.execute("Select sum("&SINFO&") from board  ") 
		allnum=allrs(0)
		set allrs=nothing
		if isnull(allnum) then allnum=1
	end if
	wid=Cint(replace(FormatPercent(rs(0)/allnum),"%",""))*5.1

	response.write "<tr>"
	response.write "<td class=tablebody2 align=center><a href=""list.asp?boardid="&rs(1)&""">"&rs(2)&"</a></td>"
	response.write "<td "
	if clng(BoardID)=rs(1) then
	response.write "class=tablebody2"
	else
	response.write "class=tablebody1"
	end if
	response.write "><img src="""&Forum_info(7)&"bar"&g&".gif"" width="""&wid&""" height=8></td>"
	response.write "<td class=tablebody2 align=center>"&rs(0)&"</td>"
	response.write "</tr>"

    rs.MoveNext
	loop
end if
rs.close
set rs=nothing
response.write "</table>"
END SUB

sub stat_var()
%>

<table cellpadding=5 cellspacing=1  class=tableborder1 align=center>
<tr>
<th id=TableTitleLink width="16%">
<a href="?boardid=<%=clng(BoardID)%>"><b>今日帖数图例</b></a>
</th>
<th id=TableTitleLink width="16%">
<a href="?action=lasttopicnum&boardid=<%=clng(BoardID)%>"><b>主题数图例</b></a>
</th>
<th id=TableTitleLink width="16%">
<a href="?action=lastbbsnum&boardid=<%=clng(BoardID)%>"><b>总帖数图例</b></a>
</th>
<th id=TableTitleLink width="16%">
<a href="?reaction=online&boardid=<%=clng(BoardID)%>"><b>在线图例</b></a>
</th>
<th id=TableTitleLink width="16%">
<a href="?reaction=onlineinfo&boardid=<%=clng(BoardID)%>"><b>在线情况</b></a>
</th>
<th id=TableTitleLink width="17%">
<a href="?reaction=onlineUserinfo&boardid=<%=clng(BoardID)%>"><b>用户组在线图例</b></a>
</th>
</tr>
<tr> 
<td  align=center colspan=6 class=tablebody1>
<%
if request("reaction")<>"" then
	if boardid>0 then
	response.write "当前论坛共有<B>"&allonline()&"</B>位用户在线，其中"&boardtype&"有<B>"&online(boardid)&"</B>位在线用户与<B>"&guest(boardid)&"</B>位客人"
	else
	response.write "当前论坛共有<B>"&allonline()&"</B>位用户在线"
	end if
else
%>
目前<b><%=boardtype%></b>  今日帖数：<%=todaynums%> 帖子总数： <%=bbsnums%> 主题总数：<%=topicnums%>
<%end if%>
</td>
</tr>
</table>
<BR>
<%
end sub

sub onlinemain()
	dim guests,page_count,Pcount
	dim totalrec,endpage
	call activeonline()
	dim mysql
	if BoardID="" then
		mysql="select username,stats,browser,startime,ip from online order by usergroupid"
	elseif not isInteger(BoardID) then
		mysql="select username,stats,browser,startime,ip from online order by usergroupid"
	else
		BoardID=clng(request("BoardID"))
		mysql="select username,stats,browser,startime,ip from online where boardid="&boardid&" order by usergroupid"
	end if
	dim totalPages,currentPage
	currentPage=request.querystring("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if

	response.write "<table cellpadding=6 cellspacing=1 align=center class=tableborder1>"
	response.write "<tr> "
	response.write "<th width=10% >用户名</th>"
	response.write "<th width=35% >当前位置</th>"
	response.write "<th width=20% >用户信息</th>"
	response.write "<th width=18% >来源鉴定</th>"
	response.write "<th width=22% >登陆时间</th>"
	response.write "</tr>"
	set rs=server.createobject("adodb.recordset")
	rs.open mysql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr> <td colspan=5 class=tablebody1>　　还没有任何用户数据。</td></tr>"
	else
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
		response.write "<tr> "
		response.write "<td align=center class=tablebody1><a href=dispuser.asp?name="&htmlencode(rs("username"))&" target=_blank>"&htmlencode(rs("username"))&"</a></td>"
		response.write "<td align=center class=tablebody2>"&htmlencode(rs("stats"))&"</td>"
		response.write "<td align=center class=tablebody1>"&replace(usersysinfo(rs("browser"),2),"操作系统：","")&"，"&replace(replace(usersysinfo(rs("browser"),1),"浏 览 器：",""),"Internet Explorer","IE")&"</td>"
		response.write "<td align=center class=tablebody2>"
		if Cint(GroupSetting(30))=1 then
			response.write "<a href=look_ip.asp?action=lookip&ip="&rs("ip")&">"&rs("ip")&"</a>"
		else
			response.write "已设置保密"
		end if
		response.write "</td><td align=center class=tablebody1>"&formatdatetime(rs("startime"),2)&"&nbsp;"&formatdatetime(rs("startime"),4)&"</td></td>"
		response.write "</tr>"
	  	page_count = page_count + 1
		rs.movenext
		wend
	end if
	response.write "</table>"
  	if totalrec mod Forum_Setting(11)=0 then
     		Pcount= totalrec \ Forum_Setting(11)
  	else
     		Pcount= totalrec \ Forum_Setting(11)+1
  	end if
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"
	response.write "<tr><td valign=middle nowrap>"
	response.write "页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"
	response.write "&nbsp;每页<b>"&Forum_Setting(11)&"</b> 总数<b>"&totalrec&"</b></td>"
	response.write "<td valign=middle nowrap align=right>分页："
	if currentpage > 4 then
	response.write "<a href=""?page=1&boardid="&boardid&"&reaction="&request("reaction")&""">[1]</a> ..."
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
		response.write " <a href=""?page="&i&"&boardid="&boardid&"&reaction="&request("reaction")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&boardid="&boardid&"&reaction="&request("reaction")&""">["&Pcount&"]</a>"
	end if
	response.write "</td></tr></table>"
	rs.close
	set rs=nothing
end sub

function INDEXonline()
	dim tmprs
	tmprs=conn.execute("Select count(id) from online WHERE boardid=0") 
	INDEXonline=tmprs(0) 
	set tmprs=nothing 
	if isnull(INDEXonline) then INDEXonline=0
end function
sub onlineinfo()
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<tr>
<th width="20%">论坛名称</th>
<th width="70%">图形比例</th>
<th width="10%">在线人数</th>
</tr>
<tr>
<td class=tablebody2 align=center>论坛其它位置</td>
<td class=tablebody1><img src=<%=Forum_info(7)%>bar10.gif width="<%=Cint(replace(FormatPercent(INDEXonline()/allonline()),"%",""))*5.1%>" height=8></td>
<td class=tablebody2 align=center><%=INDEXonline()%></td>
</tr>
<%
dim rsinfo1,rsinfo2,sqlinfo,onlineNum
dim bid
dim tmprs
dim g
g=1
if boardmaster or master then
	sqlinfo="select boardid,BoardType from board "
else
	sqlinfo="select boardid,BoardType from board where lockboard=0"
end if
set rsinfo1=conn.execute(sqlinfo)
if rsinfo1.eof and rsinfo1.bof then
	response.write "<tr> <td colspan=3 class=tablebody1>　　还没有任何用户数据。</td></tr>"
else  
do while not rsinfo1.eof 
	bid=rsinfo1(0)	
	set rsinfo2=conn.execute("Select Count(id) From online Where boardid="&bid&" ")
	onlineNum=rsinfo2(0)
	if isnull(onlineNum) then onlineNum=0
	g=g+1 
	if g=11 then g=1
	response.write "<tr>"
	response.write "<td class=tablebody2 align=center><a href=""list.asp?boardid="&bid&""">"&rsinfo1(1)&"</a></td>"
	response.write "<td class=tablebody1><img src="""&Forum_info(7)&"bar"&g&".gif"" width="""&Cint(replace(FormatPercent(onlineNum/allonline()),"%",""))*5.1&""" height=8></td>"
	response.write "<td class=tablebody2 align=center>"&onlineNum&"</td>"
	response.write "</tr>"

rsinfo1.MoveNext
loop
end if
set rsinfo1=nothing
set rsinfo2=nothing
response.write "</table>"
end sub

sub userinfo(boardid)
dim sfind
if boardid<>"" and boardid<>0 then
	sfind=" boardid="&boardid&" and "
else
	sfind=""
end if
%>
<table cellpadding=6 cellspacing=1 align=center class=tableborder1>
<tr>
<th width="20%">用户组别</th>
<th width="70%">图形比例</th>
<th width="10%" >在线人数</th>
</tr>
<%
dim trs,g,onlineNum
g=1
set rs=conn.execute("select UserGroupID,title from UserGroups order by UserGroupID")
do while not rs.eof
	set trs=conn.execute("select count(*) from online where UserGroupID="&rs(0))
	onlineNum=trs(0)
	if isnull(onlineNum) then onlineNum=0
	g=g+1 
	if g=11 then g=1
	response.write "<tr>"
	response.write "<td class=tablebody2 align=center>"&rs(1)&"</td>"
	response.write "<td class=tablebody1><img src="""&Forum_info(7)&"bar"&g&".gif"" width="""&Cint(replace(FormatPercent(onlineNum/allonline()),"%",""))*5.1&""" height=8></td>"
	response.write "<td class=tablebody2 align=center>"&onlineNum&"</td>"
	response.write "</tr>"
rs.movenext
loop
set trs=nothing
set rs=nothing
%>
</table>
<%END SUB%>
<!--#include file="online_l.asp"-->