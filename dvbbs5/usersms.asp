<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/usercp.asp"-->
<%Response.Buffer = true%>
<%
'=========================================================
' File: usersms.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim tablebody
dim boxName,smscount,smstype,readaction,turl
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if

stats="用户短信服务"
call nav()
if founderr then
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	call main()
end if
call activeonline()
call footer()

sub main()
smscount=1
select case request("action")
case "inbox"
	boxName="收件箱"
	smstype="inbox"
	readaction="read"
	turl="readsms"
	sql="select * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	call smsbox()
case "outbox"
	boxName="草稿箱"
	smstype="outbox"
	readaction="edit"
	turl="sms"
	sql="select * from message where sender='"&trim(membername)&"' and issend=0 and delS=0 order by sendtime desc"
	call smsbox()
case "issend"
	boxName="已发送的消息"
	smstype="issend"
	readaction="outread"
	turl="readsms"
	sql="select * from message where sender='"&trim(membername)&"' and issend=1 and delS=0 order by sendtime desc"
	call smsbox()
case "recycle"
	boxName="垃圾箱"
	smstype="recycle"
	readaction="read"
	turl="readsms"
	sql="select * from message where ((sender='"&trim(membername)&"' and delS=1) or (incept='"&trim(membername)&"' and delR=1)) and not delS=2 order by sendtime desc"
	call smsbox()
case else
	boxName="收件箱"
	smstype="inbox"
	readaction="read"
	turl="readsms"
	sql="select * from message where incept='"&trim(membername)&"' and issend=1 and delR=0 order by flag,sendtime desc"
	call smsbox()
end select
end sub

sub smsbox()
dim newstyle
response.write "<TABLE cellpadding=6 cellspacing=1 align=center class=tableborder1><TBODY>"&_
				"<TR>"&_
				"<TD align=center class=tablebody1><a href=usersms.asp?action=inbox><img src=pic/m_inbox.gif border=0 alt=收件箱></a> &nbsp; <a href=usersms.asp?action=outbox><img src=pic/M_outbox.gif border=0 alt=发件箱></a> &nbsp; <a href=usersms.asp?action=issend><img src=pic/M_issend.gif border=0 alt=已发送邮件></a>&nbsp; <a href=usersms.asp?action=recycle><img src=pic/M_recycle.gif border=0 alt=废件箱></a>&nbsp; <a href=friendlist.asp><img src=pic/M_address.gif border=0 alt=地址簿></a>&nbsp;<a href=JavaScript:openScript('messanger.asp?action=new',500,400)><img src=pic/m_write.gif border=0 alt=发送消息></a>"&_
                           "</td></tr></TBODY></TABLE>"
%>
<form action="messanger.asp" method=post name=inbox>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th valign=middle width=30 height=25>已读</th>
<th valign=middle width=100>
<%if smstype="inbox" or smstype="recycle" then response.write "发件人" else response.write "收件人"%>
</th>
<th valign=middle width=300>主题</th>
<th valign=middle width=150>日期</th>
<th valign=middle width=50>大小</th>
<th valign=middle width=30>操作</th>
</tr>
<%
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
<tr>
<td class=tablebody1 align=center valign=middle colspan=6>您的<%=boxname%>中没有任何内容。</td>
</tr>
<%else%>
<%do while not rs.eof%>
<%
if rs("flag")=0 then
tablebody="tablebody2"
newstyle="font-weight:bold"
else
tablebody="tablebody1"
newstyle="font-weight:normal"
end if
%>
<tr>
<td class=<%=tablebody%> align=center valign=middle>
<%
select case smstype
case "inbox"
	if rs("flag")=0 then
	response.write "<img src="""&Forum_info(7)&"m_news.gif"">"
	else
	response.write "<img src="""&Forum_info(7)&"m_olds.gif"">"
	end if
case "outbox"
	response.write "<img src="""&Forum_info(7)&"m_issend_2.gif"">"
case "issend"
	response.write "<img src="""&Forum_info(7)&"m_issend_1.gif"">"
case "recycle"
	if rs("flag")=0 then
	response.write "<img src="""&Forum_info(7)&"m_news.gif"">"
	else
	response.write "<img src="""&Forum_info(7)&"m_olds.gif"">"
	end if
end select
%>
</td>
<td class=<%=tablebody%> align=center valign=middle style="<%=newstyle%>">
<%if smstype="inbox" or smstype="recycle" then%>
<a href="dispuser.asp?name=<%=htmlencode(rs("sender"))%>" target=_blank><%=htmlencode(rs("sender"))%></a>
<%else%>
<a href="dispuser.asp?name=<%=htmlencode(rs("incept"))%>" target=_blank><%=htmlencode(rs("incept"))%></a>
<%end if%>
</td>
<td class=<%=tablebody%> align=left style="<%=newstyle%>"><a href="JavaScript:openScript('messanger.asp?action=<%=readaction%>&id=<%=rs("id")%>&sender=<%=rs("sender")%>',500,400)"><%=htmlencode(rs("title"))%></a>	</td>
<td class=<%=tablebody%> style="<%=newstyle%>"><%=rs("sendtime")%></td>
<td class=<%=tablebody%> style="<%=newstyle%>"><%=len(rs("content"))%>Byte</td>
<td align=center valign=middle width=30 class=<%=tablebody%>><input type=checkbox name=id value=<%=rs("id")%>></td>
</tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
<tr> 
<td align=right valign=middle colspan=6 class=tablebody2>节省每一分空间，请及时删除无用信息&nbsp;<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除<%=replace(boxname,"箱","")%>">&nbsp;<input type=submit name=action onclick="{if(confirm('确定清除<%=boxname%>所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空<%=boxname%>"></td>
</tr>
</table>
</form>
<%
end sub
%>