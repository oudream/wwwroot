<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
'=========================================================
' File: printpage.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim urs,usql
dim rsboard,boardsql
dim announceid
dim username
dim rootid
dim topic
dim abgcolor
abgcolor="#FFFFFF"
Dim TotalUseTable
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>请指定相关贴子。"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>非法的贴子参数。"
else
	AnnounceID=request("id")
end if

if foundErr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call announceinfo()
	if founderr then call dvbbs_error()
end if
call footer()

sub announceinfo()
	if clng(Announceid)>clng(MaxOldID) then
	TotalUseTable=NowUseBbs
	else
	TotalUseTable="bbs1"
	end if
    set rs=conn.execute("select title from topic where topicID="&AnnounceID)
	if not(rs.bof and rs.eof) then
		topic=rs(0)
	else
		foundErr = true
		ErrMsg=ErrMsg+"<br>"+"<li>您指定的贴子不存在</li>"
		exit sub
	end if
	rs.close
%>
<HTML><HEAD><TITLE><%=Forum_info(0)%>--显示贴子</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<!--#include file="inc/Forum_css.asp"-->
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgcolor="#ffffff" alink="#333333" vlink="#333333" link="#333333" topmargin="10" leftmargin="10">
<TABLE border=0 width="<%=Forum_body(12)%>" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
<b>以文本方式查看主题</b><br><br>
-&nbsp;&nbsp;<b><%=Forum_info(0)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>index.asp)<br>
--&nbsp;&nbsp;<b><%=boardtype%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>bbs.asp?boardid=<%=boardid%>)<br>
----&nbsp;&nbsp;<b><%=htmlencode(topic)%></b>&nbsp;&nbsp;(<%=Forum_info(1)%>dispbbs.asp?boardid=<%=boardid%>&rootid=<%=rootid%>&id=<%=announceid%>)
      </TD></TR></TBODY></TABLE> 
<br>
<hr>
<%
	Rs.open "Select b.UserName,b.Topic,b.dateandtime,b.body,u.UserGroupID from "&TotalUseTable&" b inner join [user] u on b.PostUserID=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" order by b.announceid",conn,1,1
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>没有找到相关帖子。"
		exit sub
	end if
	do while not rs.eof
%>
<TABLE border=0 width="750" align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top>
--&nbsp;&nbsp;作者：<%=rs("username")%><br>
--&nbsp;&nbsp;发布时间：<%=rs("dateandtime")%><br><br>
--&nbsp;&nbsp;<%=htmlencode(rs("topic"))%><br>
<%=dvbcode(rs("body"),rs("UserGroupID"))%>
	<hr>
    </TD></TR></TBODY></TABLE> 
<%
          rs.movenext
        loop	
	rs.close
	set rs=nothing
	call activeonline()
end sub    
%>