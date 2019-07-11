<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
'=========================================================
' File: list.asp
' Version:5.0
' Date: 2002-9-20
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

Dim currentPage,tl,limitime
Dim Fast_action
Fast_action=false

if boardmaster or master then
	guestlist=""
	Fast_action=true
else
	guestlist=" lockboard<>1 and "
	Fast_action=false
end if

BoardID = Request("BoardID")
currentPage=request("page")
if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>错误的版面参数！请确认您是从有效的连接进入。"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("selTimeLimit")="all" then
	tl=""
elseif request("selTimeLimit")="" then
	tl=""
else
	limitime=request("selTimeLimit")
	tl=" and dateandtime>='"&cstr(cdate(now()-limitime))&"' "
end if
if cint(lockboard)=2 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请<a href=login.asp>登陆</a>并确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>本论坛为认证论坛，请确认您的用户名已经得到管理员的认证后进入。"
		founderr=true
		end if
	end if
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(1)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
Dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
dim Pnum,p,replynum
dim totalrec,ii,page_count
dim n,pi
dim call_info
dim onlinenum,guestnum
dim rs1,sql1
onlinenum=online(boardid)
guestnum=guest(boardid)
if trim(boardmasterlist)<>"" then
	master_1=split(boardmasterlist, "|")
	for i = 0 to ubound(master_1)
	master_2=""+master_2+"<a href=""dispuser.asp?name="+master_1(i)+""" target=_blank title=点击查看该版主资料>"+master_1(i)+"</a>&nbsp;"
	next
else
	master_2="无版主"
end if
call activeonline()
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=100% valign=middle>
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<%
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='当前没有公告'>当前没有公告</ACRONYM></a></b> ("& now() &")"
else
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
%>","","欢迎光临<%=Forum_info(0)%> <%=boardtype%>","",
"严禁恶意使用粗言秽语，违者经劝告无效，立即封ID。",""
]
</SCRIPT>
<div id="elFader" style="position:relative;visibility:hidden; height:16" ></div>
</td>
</tr>
</table><BR>
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	if (targetImg.src.indexOf("nofollow")!=-1){return false;}
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="pic/minus.gif";
			if (targetImg.loaded=="no"){
				document.frames["hiddenframe"].location.replace("loadtree1.asp?boardid="+b_id+"&rootid="+t_id);
			}
		}else{
			targetDiv.style.display="none";
			targetImg.src="pic/plus.gif";
		}
	}
}
</script>

<iframe width=0 height=0 src="" id="hiddenframe"></iframe>
<!--<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=34 valign=middle> <img src="<%=Forum_info(7)&Forum_boardpic(16)%>" border=0  width=20 height=17>
</td><td valign=middle align=left>
<%
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='当前没有公告'>当前没有公告</ACRONYM></a></b> ("& now() &")"
else
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
%>
</td>
<td align=right valign=middle><p>
<form action="list.asp" method=get>
<input type=hidden name=boardid value="<%= boardid %>">
<select name=selTimeLimit onchange='javascript:submit()'>
<option value=all>查看所有的主题
<option value=1>查看一天内的主题
<option value=2>查看两天内的主题
<option value=7>查看一星期内的主题
<option value=15>查看半个月内的主题
<option value=30>查看一个月内的主题
<option value=60>查看两个月内的主题
<option value=180>查看半年内的主题
</select>
</form></p>
</td></tr>
</TABLE>
--><BR>
<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TR>
<Th height=25 width="100%" align=left id=tabletitlelink style="font-weight:normal">总在线<b><%=allonline()%></b>人，其中<%= boardtype %>上共有 <b><%= onlinenum %></b> 位会员与 <b><%= guestnum %></b> 位客人.今日贴子 <b><font color=<%=Forum_body(8)%>><%= todaynum %></font></b> 
<%
if request("action")="show" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
else
	if cint(Forum_Setting(14))=1 and request("action")<>"off" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">关闭详细列表</a>]"
	else
	response.write "[<a href=list.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">显示详细列表</a>]"
	end if
end if
%>
</Th></TR>
<%
if request("action")="off" then
	call onlineuser(0,0,boardid)
elseif request("action")="show" then
	call onlineuser(1,1,boardid)
else
	call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
end if
%>
</td></tr>
</TABLE>
<BR>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center valign=middle><tr>
<td align=center width=2> </td>
<td align=left> <a href="announce.asp?boardid=<%= boardid %>"><img src="<%=Forum_info(7)&Forum_boardpic(1)%>" border=0 alt="发新帖"></a>
&nbsp;&nbsp;<a href="vote.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(2)%>" border=0 alt="发起新投票"></a>
&nbsp;&nbsp;<a href="smallpaper.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(3)%>" border=0 alt="发布小字报"></a></td>
<td align=right><img src="<%=Forum_info(7)%>team2.gif" align=absmiddle>  <%= master_2 %></td></tr></table>

<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>
<tr><td class=tablebody1 colspan=5 height=20>
<table width=100% >
<tr>
<td valign=middle height=20 width=50> <a href=AllPaper.asp?boardid=<%=boardid%> title=点击查看本论坛所有小字报><b>广播：</b></a> </td><td width=*> 
<marquee scrolldelay=150 scrollamount=4 onmouseout="if (document.all!=null){this.start()}" onmouseover="if (document.all!=null){this.stop()}">
<%
set rs=conn.execute("SELECT TOP 5 s_id,s_username,s_title FROM smallpaper where datediff('d',s_addtime,Now())<=1 and s_boardid="&boardid&" ORDER BY s_addtime desc")
do while not rs.eof
	response.write "<font color="&Forum_pic(5)&">"&htmlencode(rs(1))&"</font>说：<a href=javascript:openScript('viewpaper.asp?id="&rs(0)&"&boardid="&boardid&"',500,400)>"&htmlencode(rs(2))&"</a>"
rs.movenext
loop
set rs=nothing
%>
</marquee><td align=right width=200><a href='elist.asp?boardid=<%= boardid %>' title=查看本版精华><font color=<%=Forum_body(8)%>><B>精华</B></font></a> | <a href=boardstat.asp?reaction=online&boardid=<%=boardid%> title=查看本版在线详细情况>在线</a> | <a href=bbseven.asp?boardid=<%=boardid%> title=查看本版事件>事件</a> | <a href=Boardhelp.asp?boardid=<%=boardid%> title=查看本版帮助>帮助</a> | <a href='?boardid=<%=boardid%>&action=batch'>管理</a></td></tr>
</table>
</td></tr>
<form action="admin_batch.asp" method=post name=batch>
<TR align=middle>
<Th height=25 width=32 id=tabletitlelink><a href="list.asp?boardid=<%=boardid%>&page=<%=request("page")%>&action=batch">状态</a></th>
<Th width=*>主 题  (点<img src=<%=Forum_info(7)%>plus.gif>即可展开贴子列表)</Th>
<Th width=80>作 者</Th>
<Th width=64>回复/人气</Th>
<Th width=195>最后更新 | 回复人</Th>
</TR>
<%
if limitime="" then
	totalrec=lasttopicnum
else
	set rs=conn.execute("select count(topicid) from topic where boardID="&cstr(boardID)&" "&tl&"")
	totalrec=rs(0)
end if
set rs=server.createobject("adodb.recordset")
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where boardid="&boardid&" and not locktopic=2 order by istop desc,lastposttime desc"

rs.open sql,conn,1
if rs.bof and rs.eof then
	response.write "<tr><td colSpan=5 width=100% class=tablebody1>本论坛暂无内容，欢迎发贴：）</td></tr>"
else
	if totalrec mod Forum_Setting(11)=0 then
		n= totalrec \ Forum_Setting(11)
	else
	     	n= totalrec \ Forum_Setting(11)+1
  	end if
	RS.MoveFirst
	if currentpage > n then currentpage = n
      	if currentpage<1 then currentpage=1
	RS.Move (currentpage-1) * Forum_Setting(11)
	do while not rs.eof and page_count<Clng(Forum_Setting(11))
	page_count=page_count+1
'========================TopicInfo=========================
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"

if request("action")="batch" and Cint(GroupSetting(45))=1 then
	response.write "<input type=checkbox name=Announceid value="&rs("topicid")&">"
else
	if rs("istop")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(3)&""" alt=固顶主题>"
	elseif rs("isbest")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=精华帖子>"
	elseif rs("isvote")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(12)&""" alt=投票贴子>"
	elseif rs("locktopic")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=本主题已锁定>"
	elseif lockboard=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=本论坛已锁定>"
	elseif rs("child")>=10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(1)&""" alt=热门主题>"
	else
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=开放主题>"
	end if
end if

response.write "</TD><TD align=left class=tablebody1 width=*>"

if rs(6)=0 then
	response.write "<img src="""& Forum_info(7) &"nofollow.gif"" id=""followImg"& rs("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7) &"plus.gif"" id=""followImg"& rs("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& rs("topicid") &","& boardid &")"" title=展开贴子列表>"
end if

if not isnull(rs("lastpost")) then
	LastPost=split(rs("lastpost"),"$")
	if ubound(LastPost)=6 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=rs("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src="&Forum_info(7)&LastPost(4)&".gif width=16 height=16>"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
	end if
else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
end if
if LastTime=rs("dateandtime") then LastUser=""

response.write "<a href=""dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &""" title=""《"& htmlencode(rs("title")) &"》<br>作者："& htmlencode(rs("postusername")) &"<br>发表于"& rs("dateandtime") &"<br>最后跟贴："& body &"..."">"

if len(rs("title"))>26 then
	response.write left(htmlencode(rs("title")),26)&"..."
else
	response.write htmlencode(rs("title"))
end if

response.write "</a>"
replynum=rs("child")+1
if replynum>Cint(Forum_Setting(12)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&"multipage.gif><b>"
  	if replynum mod Cint(Forum_Setting(12))=0 then
     		Pnum= replynum \ Cint(Forum_Setting(12))
  	else
     		Pnum= replynum \ Cint(Forum_Setting(12))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 then
	response.write "... <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if

response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& rs("postuserid") &""" target=_blank>"& htmlencode(rs("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=64>"

if rs("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&rs("votetotal")&"</b></font>  票"
else
	response.write rs("child") &"/"& rs("hits")
end if

response.write "</TD>"
response.write "<TD align=left class=tablebody2 width=195>&nbsp;&nbsp;<a href=""dispbbs.asp?boardid="& boardid &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write FormatDateTime(LastTime,2)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a>&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& rs("topicid") &"""><td colspan=5 id=""followTd"& rs("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& rs("topicid") &","&boardid&")"">正在读取关于本主题的跟贴，请稍侯……</div></td></tr>"
'===========================End============================
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	rs.movenext
	loop

	if request("action")="batch" then
%>
<tr><td height=30 width=100% class=tablebody1 colspan=5><select name=newboard size=1>
<option selected>移动帖子请选择</option>
<%
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option>╋ "& rs(1) &"</option>"

		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			response.write "<option>没有论坛</option>"
		else
			do while not rs1.eof
				response.write "<option value="""&rs1(0)&""">　├" & rs1(1) & "</option>"
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
%>
<input type=hidden value="<%=boardid%>" name=boardid><input type=checkbox name=chkall value=on onclick='CheckAll(this.form)'>全选/取消 <input type=radio name=action value=dele>批量删除 <input type=radio name=action value=move>批量移动 <input type=radio name=action value=isbest>批量精华 <input type=radio name=action value=lock>批量锁定 <input type=radio name=action value=istop>批量固顶 <input type=submit name=Submit value=执行  onclick="{if(confirm('您确定执行的操作吗?')){this.document.batch.submit();return true;}return false;}"></font></td></tr></form>
<%
	end if
	set rs1=nothing
end if
set rs=nothing
if currentpage-1 mod 10=0 then
	p=(currentpage-1) \ 10
else
	p=(currentpage-1) \ 10
end if
dim pagelist,pagelistbit
%>
</table>

<table border=0 cellpadding=0 cellspacing=3 width=<%=Forum_body(12)%> align=center >
</form>
<form method=post action=list.asp>
<input type=hidden name=selTimeLimit value='<%= request("selTimeLimit") %>'><tr>
<td valign=middle>页次：<b><%= currentPage %></b>/<b><%= n %></b>页 每页<b><%= Forum_Setting(11) %></b> 主题数<b><%= totalrec %></b></td>
<td valign=middle><div align=right >分页：
<%
	if currentPage=1 then
	response.write "<font face=webdings color="&Forum_body(8)&">9</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page=1&selTimeLimit="&request("selTimeLimit")&"' title=首页><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?boardid="&boardid&"&page="&Cstr(p*10)&"&selTimeLimit="&request("selTimeLimit")&"' title=上十页><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		   if ii=currentPage then
	          response.write "<font color="&Forum_body(8)&">"+Cstr(ii)+"</font> "
		   else
		      response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&selTimeLimit="&request("selTimeLimit")&"'>"+Cstr(ii)+"</a>   "
		   end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&selTimeLimit="&request("selTimeLimit")&"' title=下十页><font face=webdings>8</font></a>   "
	if currentPage=n then
	response.write "<font face=webdings color="&Forum_body(8)&">:</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page="&Cstr(n)&"&selTimeLimit="&request("selTimeLimit")&"' title=尾页><font face=webdings>:</font></a>   "
	end if
%>
转到:<input type=text name=Page size=3 maxlength=10  value='<%= currentpage %>'><input type=submit value=Go name=submit>
</div></td></tr>
<input type=hidden name=BoardID value='<%= BoardID %>'>
</form></table>

<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr>
<FORM METHOD=POST ACTION="queryresult.asp?boardid=<%=boardid%>&sType=2&SearchDate=30&pSearch=1">
<td width=50% valign=middle nowrap height=40>
快速搜索：<input type=text name="keyword">&nbsp;<input type=submit name=submit value="搜索">
</td>
</FORM>
<td valign=middle nowrap width=50%> <div align=right>
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option selected>跳转论坛至...</option>
<%
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option>╋ "& rs(1) &"</option>"

		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			response.write "<option>没有论坛</option>"
		else
			do while not rs1.eof
				response.write "<option value=""list.asp?boardid="&rs1(0)&""">　├" & rs1(1) & "</option>"
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
	set rs=nothing
	set rs1=nothing
%>
</select><div></td></tr></table>
<BR>
<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>
<th width=80% align=left>　-=> <%= Forum_info(0) %>图例</th>
<th noWrap width=20% align=right>所有时间均为 - <%=Forum_info(9)%> &nbsp;</th>
</tr>
<tr><td colspan=2 class=tablebody1>
<table cellspacing=4 cellpadding=0 width=92% border=0 align=center>
<tr><td><img src=<%=Forum_info(7)&Forum_statePic(0)%>> 开放的主题</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(1)%>> 回复超过10贴</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(2)%>> 锁定的主题</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(3)%>> 固定顶端的主题 </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(4)%>> 精华帖子 </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(12)%>> 投票帖子 </td>
</tr>
</table>
</td></tr></table>
<%
end sub

if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
	if Forum_ads(2)=1 then call admove()
	if Forum_ads(13)=1 then call fixup()
end if
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->