<%
sub cp_var()
%>
<BR>
<table cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<tr>
<%if not founduser then%>
<td height=25>
>> 欢迎您 , 请您先 <a href="login.asp">登陆</a> 或 <a href="reg.asp">注册</a>
<%else%>
<td width=65% >
>> 欢迎 <B><%=membername%></B>  [ <a href="JavaScript:openScript('messanger.asp?action=new',500,400)">发短信</a> :: <a href="topicwithme.asp?s=2">发表的主题</a> :: <a href="topicwithme.asp?s=1">参与的主题</a> :: <a href="dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission">我能做什么</a> ]
</td><td width=35% align=right>
<%if Cint(newincept())>Cint(0) then%>
<bgsound src="pic/mail.wav" border=0>
<img src=<%=Forum_info(7)%>msg_new_bar.gif> <a href="usersms.asp?action=inbox">我的收件箱</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> 新</font></a>)
<%else%>
<img src=<%=Forum_info(7)%>msg_no_new_bar.gif> <a href="usersms.asp?action=inbox">我的收件箱</a> (<font color=gray>0 新</font>)
<%end if%>
<%end if%>
</td></tr>
</table>

<table cellspacing=1 cellpadding=3 align=center class=tableBorder2>
<tr><td height=25 valign=middle>
<img src="<%=Forum_info(7)%>Forum_nav.gif" align=absmiddle> <a href=index.asp><%=Forum_info(0)%></a> → 
<a href=usermanager.asp>用户<%=Membername%>的控制面板</a> → <%=stats%>
<a name=top></a>
</td></td>
</table>
<br>

<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th width=14% height=25 id=tabletitlelink><a href=usermanager.asp>控制面板首页</a></th>
<th width=14%  id=tabletitlelink><a href=mymodify.asp>基本资料修改</a></th>
<th width=14%  id=tabletitlelink><a href=modifypsw.asp>用户密码修改</a></th>
<th width=14%  id=tabletitlelink><a href=modifyadd.asp>联系资料修改</a></th>
<th width=14%  id=tabletitlelink><a href=usersms.asp>用户短信服务</a></th>
<th width=14%  id=tabletitlelink><a href=friendlist.asp>编辑好友列表</a></th>
<th width=14%  id=tabletitlelink><a href=favlist.asp>用户收藏管理</a></th>
</tr>
</table>
<hr size=1 width="<%=Forum_body(12)%>">
<%
end sub
%>