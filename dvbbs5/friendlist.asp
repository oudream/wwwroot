<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: friendlist.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim msg
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	founderr=true
end if
stats="好友列表"
call nav()
if founderr=true then
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	select case request("action")
	case "info"
		call info()
	case "addF"
		call addF()
	case "saveF"
		call saveF()
	case "删除"
		call DelFriend()
	case "清空好友"
		call AllDelFriend()
	case else
		call info()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub info()
%>
<form action="friendlist.asp" method=post name=inbox>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
            <tr>
                <th valign=middle width="25%">姓名</td>
                <th valign=middle width="25%">邮件</td>
                <th valign=middle width="25%">主页</td>
                <th valign=middle width="10%">OICQ</td>
                <th valign=middle width="10%">发短信</td>
                <th valign=middle width="5%">操作</td>
            </tr>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select F.*,U.useremail,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&trim(membername)&"' order by F.f_addtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tablebody1 align=center valign=middle colspan=6>您的好友列表中没有任何内容。</td>
                </tr>
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tablebody1><a href="dispuser.asp?name=<%=htmlencode(rs("F_friend"))%>" target=_blank><%=htmlencode(rs("F_friend"))%></a></td>
                    <td align=center valign=middle class=tablebody1><a href="mailto:<%=htmlencode(rs("useremail"))%>"><%=htmlencode(rs("useremail"))%></a></td>
                    <td align=center class=tablebody1><a href="<%=htmlencode(rs("homepage"))%>" target=_blank><%=htmlencode(rs("homepage"))%></a></td>
                    <td align=center class=tablebody1><%=htmlencode(rs("oicq"))%></td>
                    <td align=center class=tablebody1><a href="usersms.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>">发送</a></td>
                <td align=center class=tablebody1><input type=checkbox name=id value=<%=rs("f_id")%>></td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
                
        <tr> 
          <td align=right valign=middle colspan=6 class=tablebody2><input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">选中所有显示记录&nbsp;<input type=button name=action onclick="location.href='friendlist.asp?action=addF'" value="添加好友">&nbsp;<input type=submit name=action onclick="{if(confirm('确定删除选定的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="删除">&nbsp;<input type=submit name=action onclick="{if(confirm('确定清除所有的纪录吗?')){this.document.inbox.submit();return true;}return false;}" value="清空好友"></td>
                </tr>
                </table>
</form>
<%
end sub

sub delFriend()
dim delid
delid=replace(request.form("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"请选择相关参数。"
founderr=true
exit sub
else
	conn.execute("delete from Friend where F_username='"&trim(membername)&"' and F_id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除选定的好友记录。"
	call dvbbs_suc()
end if
end sub
sub AllDelFriend()
	conn.execute("delete from Friend where F_username='"&trim(membername)&"'")
	sucmsg=sucmsg+"<br>"+"<li><b>您已经删除了所有好友列表。"
	call dvbbs_suc()
end sub

sub addF()
%>
<form action="Friendlist.asp" method=post name=messager>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=2> 
              <input type=hidden name="action" value="saveF">
              加入好友--请完整输入下列信息</th>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle width=70><b>好友：</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="touser" size=50 value="<%=request("myFriend")%>">
			  &nbsp;使用逗号（,）分开，最多5位用户
            </td>
          </tr>
          <tr> 
            <td valign=middle colspan=2 align=center class=tablebody2> 
              <input type=Submit value="保存" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="清除">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub saveF()
dim incept
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>您忘记填写发送对象了吧。"
	founderr=true
	exit sub
else
	incept=checkStr(request("touser"))
	incept=split(incept,",")
end if

for i=0 to ubound(incept)
set rs=server.createobject("adodb.recordset")
sql="select username from [user] where username='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>论坛没有这个用户，操作未成功。"
	founderr=true
	exit sub
end if
set rs=nothing

sql="select F_friend from friend where F_username='"&membername&"' and  F_friend='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	sql="insert into friend (F_username,F_friend,F_addtime) values ('"&membername&"','"&incept(i)&"',Now())"
	conn.execute(sql)
end if
if i>4 then
	errmsg=errmsg+"<br>"+"<li>每次最多只能添加5位用户，您的名单5位以后的请重新填写。"
	founderr=true
	exit sub
	exit for
end if
next

sucmsg=sucmsg+"<br>"+"<li><b>恭喜您，好友添加成功。"
call dvbbs_suc()
end sub
%>