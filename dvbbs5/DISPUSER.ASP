<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/birthday.asp"-->
<%
'=========================================================
' File: dispuser.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
dim username
dim abgcolor
dim masterinfo
dim usertopic,usertopic_1
dim delpoint
dim userclass
dim suserid
username=trim(checkStr(request("name")))
stats="查询"&username&"个人资料"
founderr=false
dim usersign
usersign=true
call nav()
call head_var(1)
if founderr then
	call dvbbs_error()
else
	if request("action")="permission" then
		call mypermission()
	else
		call main()
	end if
	if founderr then call dvbbs_error()
end if
call footer()
sub main()
if Cint(GroupSetting(1))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛会员资料的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	founderr=true
	exit sub
end if
if isnumeric(request("id")) then
	suserid=request("id")
else
	if username="" then
   		ErrMsg=Errmsg+"<br>"+"<li>错误的参数。"
		founderr=true
		exit sub
	end if
end if
dim stat,OnlineTime
dim Srs
dim Stockmoney,bankmoney,usergrade
dim userinfo,userbaseinfo,usertrueinfo
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,useraddress
dim usersetting,setuserinfo,setusertrue
set rs=server.createobject("adodb.recordset")
if suserid="" then
sql="select * from [User] where username='"&UserName&"'"
else
sql="select * from [user] where userid="&suserid
end if
rs.open sql,conn,1,1
if err.number<>0 then 
	ErrMsg=Errmsg+"<br>"+"<li>数据库操作失败："&err.description
	founderr=true
	exit sub
end if
if rs.eof and rs.bof then
	ErrMsg=Errmsg+"<br>"+"<li>您查询的名字不存在"
	founderr=true
	exit sub
end if
suserid=rs("userid")
userclass=rs("userclass")
username=rs("username")
if rs("usersetting")<>"" then
	usersetting=split(rs("usersetting"),"|||")
	if ubound(usersetting)=1 then
	setuserinfo=usersetting(0)
	setusertrue=usersetting(1)
	else
	setuserinfo=1
	setusertrue=0
	end if
else
	setuserinfo=1
	setusertrue=0
end if
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		useraddress=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		useraddress=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	useraddress=""
end if
	if boardmaster or master then
	set Srs=conn.execute("select stats,startime from online where userid="&rs("userid"))
	else
	set Srs=conn.execute("select stats,startime from online where userid="&rs("userid")&" and userhidden=2")
	end if
	if Srs.eof and Srs.bof then
	stat=""
	OnlineTime="[离线]"
	else
	stat="当前位置：["&Srs(0)&"]"
	OnlineTime="[在线：" & datediff("n",Srs(1),Now()) & "Mins]"
	end if
	Srs.close
	set Srs=nothing

	Stockmoney=0
	bankmoney=0
%>
<table width=<%=Forum_body(12)%> border=0 cellspacing=0 cellpadding=3 align=center>
  <tr> 
    <td><img src=<%=htmlencode(rs("face"))%> width=20 height=21 align=absmiddle> 
<b><%=htmlencode(username)%></b> 
</td>
    <td align=right>
<%=stat%><img src=<%=Forum_info(7)%>zhuangtai.gif width=16 height=16 align=absmiddle> 
      状态：
<%
	if rs("lockuser")=1 then
		response.write "锁定  " & onlinetime
	elseif rs("lockuser")=2 then
		response.write "屏蔽  " & onlinetime
	else
		response.write "正常  " & onlinetime
	end if
%>
  </td>
  </tr>
</table>
<%if setuserinfo=1 then%>
<table cellspacing=1 cellpadding=3 align=center  style="table-layout:fixed;word-break:break-all" class=tableborder1>
  <col width=20% ><col width=*><col width=40% > 
  <tr> 
    <th colspan=2 align=left>基本资料</th>
    <td rowspan=9 align=center class=tablebody1 width=40% valign=top>
<%=iimg(rs("userphoto"),"<font color=gray>无</font>","<img src="&htmlencode(FilterJS(rs("userPhoto")))&">")%>
    </td>
  </tr>   
  <tr> 
    <td class=tablebody1 width=20% align=right>性 别：</td>
    <td class=tablebody1><%=iif(rs("sex"),"男","女")%> </td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>出 生：</td>
    <td class=tablebody2>
<%=iimg(rs("birthday"),"<font color=gray>未填</font>",rs("birthday"))%>
 </td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>星 座：</td>
    <td class=tablebody1>
<%=iimg(rs("birthday"),"<font color=gray>未填</font>",astro(rs("birthday")))%>
</td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>Ｅｍａｉｌ：</td>
    <td class=tablebody2>
<%=iimg(rs("useremail"),"<font color=gray>未填</font>","<a href=mailto:"&htmlencode(rs("useremail"))&">"&htmlencode(rs("useremail"))&"</a>")%>
</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>Ｑ Ｑ：</td>
    <td class=tablebody1>
<%=iimg(rs("oicq"),"<font color=gray>未填</font>",htmlencode(rs("oicq")))%>
</td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>ＩＣＱ：</td>
    <td class=tablebody2>
<%=iimg(rs("icq"),"<font color=gray>未填</font>",htmlencode(rs("icq")))%>
</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>ＭＳＮ：</td>
    <td class=tablebody1>
<%=iimg(rs("msn"),"<font color=gray>未填</font>",htmlencode(rs("msn")))%>
 </td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>主 页：</td>
    <td class=tablebody2>
<%=iimg(rs("homepage"),"<font color=gray>未填</font>","<a href="&htmlencode(rs("homepage"))&" target=_blank>"&htmlencode(rs("homepage"))&"</a> ")%>
</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right valign=top>&nbsp;</td>
    <td class=tablebody1>&nbsp;</td>
    <td class=tablebody1 align=center width=40% >
      <b><a href="javascript:openScript('messanger.asp?action=new&touser=<%=htmlencode(username)%>',500,400)">给他留言</a> | <a href="friendlist.asp?action=addF&myFriend=<%=HTMLEncode(username)%>" target=_blank>加为好友</a></b></td>
  </tr>
</table>
<br>
<%end if%>

<%if setusertrue=1 then%>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1 style="table-layout:fixed;word-break:break-all">
  <col width=20% ><col width=*><col width=40% > 
  <tr> 
    <th colspan=2 align=left>
      用户详细资料</th>
    <td rowspan=14 class=tablebody1 width=40% valign=top>
<b>性格：</b>
<br>
<%=character%>
<br><br><br>
<b>个人简介：</b><br>
<%=personal%>
<br>
</td>
  </tr>   
  <tr> 
    <td class=tablebody1 width=20% align=right>真实姓名：</td>
    <td class=tablebody1><%=realname%></td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>国　　家：</td>
    <td class=tablebody2><%=country%> </td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>联系电话：</td>
    <td class=tablebody1><%=userphone%></td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>通信地址：</td>
    <td class=tablebody2><%=useraddress%></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>省　　份：</td>
    <td class=tablebody1><%=province%></td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>城　　市：</td>
    <td class=tablebody2><%=city%></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>生　　肖：</td>
    <td class=tablebody1><%=shengxiao%> </td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>血　　型：</td>
    <td class=tablebody2><%=blood%></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>信　　仰：</td>
    <td class=tablebody1><%=belief%></td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>职　　业：</td>
    <td class=tablebody2><%=occupation%></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>婚姻状况：</td>
    <td class=tablebody1><%=marital%></td>
  </tr>
  <tr> 
    <td class=tablebody2 width=20% align=right>最高学历：</td>
    <td class=tablebody2><%=education%></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=20% align=right>毕业院校：</td>
    <td class=tablebody1><%=college%></td>
  </tr></table>
<br>
<%end if%>

<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
  <tr>
    <th align=left colspan=6> 论坛属性</th>
  </tr>
  <tr>
    <td class=tablebody1 width=15% align=right>经验值：</td>

    <td  width=35%  class=tablebody1><b><%= rs("userEP") %> </b></td>
    <td width=15% align=right class=tablebody1>精华帖子：</td>
    <td width=35%  class=tablebody1> <b><%=rs("userisbest")%></b>篇</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>魅力值：</td>
    <td  width=35%  class=tablebody1><b><%= rs("userCP") %> </b></td>
    <td width=15% align=right class=tablebody1>帖子总数：</td>
    <td width=35%  class=tablebody1><b><%=rs("article")%></b> 篇</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>论坛等级：</td>
    <td  width=35%  class=tablebody1><b><%=rs("userclass")%> </b></td>
    <td width=15% align=right class=tablebody1>被删主题：</td>
    <td width=35%  class=tablebody1><b><font color=<%=Forum_body(8)%>><%=rs("userDel")%></font></b> 
      篇</td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>威望值：</td>
    <td  width=35%  class=tablebody1><b><font color=<%=Forum_body(8)%>><%=rs("userPower")%></font> </b></td>
    <td width=15% align=right class=tablebody1>被删除率：</td>
<td width=35%  class=tablebody1><b></b> <font color=<%=Forum_body(8)%>><b>
<%
	if rs("userDel")>0 or rs("article")>0 then
		response.write FormatPercent(-rs("userDel")/rs("article"))
	else
		response.write "%"
	end if
%>
</b></font> 
    </td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>门  派：</td>
    <td  width=35%  class=tablebody1><b>
<%=iimg(rs("usergroup"),"<font color=gray>无</font>",rs("usergroup"))%>
 </b></td>
    <td class=tablebody1 width=15% align=right>登陆次数：</td>
    <td width=35%  class=tablebody1><b><%=rs("logins")%></b> 
    </td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>注册日期：</td>
    <td  width=35%  class=tablebody1><b><%=rs("addDAte")%></b></td>
    <td width=15% align=right class=tablebody1>上次登录：</td>
    <td width=35%  class=tablebody1><b><%=rs("lastlogin")%></b></td>
  </tr>
</table>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
  <tr> 
    <th align=left colspan=4>
      资产情况</th>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>现金货币：</td>
    <td width=35%  class=tablebody1><b><%= rs("userWealth") %></b></td>
    <td colspan=2 valign=top rowspan=4 class=tablebody1>论坛职务：
      <hr size=1 width=100 align=left>
<%
	if rs("UserGroupID")=1 then
		response.write "管理员<br>"
	end if
	if rs("UserGroupID")<=3 then
	set Srs=conn.execute("select boardmaster,boardid,boardtype from board order by boardid")
	do while not srs.eof
	boardmaster=split(srs(0),"|")
	for i=0 to ubound(boardmaster)
	if boardmaster(i)=username then response.write "<a href=list.asp?boardid="&srs(1)&">"&srs(2)&"</a>  版主<br>"
	next
	srs.movenext
	loop
	srs.close
	set srs=nothing
	end if
	if rs("userGroupID")>3 then
	response.write "<font color=gray>无职务</font>"
	end if
%>
      </td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>股票市值：</td>
    <td  width=35%  class=tablebody1><b><%=Stockmoney%></b></td>
  </tr>

  <tr> 
    <td class=tablebody1 width=15% align=right>银行存款：</td>
    <td width=35%  class=tablebody1><b><%=bankmoney%></b></td>
  </tr>
  <tr> 
    <td class=tablebody1 width=15% align=right>总 资 产：</td>
    <td width=35%  class=tablebody1><b><%=rs("userWealth")+stockmoney+bankmoney%></b></td>
  </tr>
</table>
<br>
<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
  <tr> 
    <th align=left id=tabletitlelink>
      最近活动－－<a href="queryResult.asp?stype=1&nSearch=3&keyword=<%=HTMLEncode(UserName)%>&SearchDate=ALL&boardid=0">搜索该用户所有帖子</a></th>
  </tr>
<tr> 
    <td class=tablebody1 align=left>
<%
	set srs=conn.execute("select top 5 rootid,boardid,dateandtime,topic,expression,announceid,body from "&NowUseBBS&" where postuserid="&rs("userid")&" and not locktopic=2 order by announceid desc")
	do while not srs.eof
	response.write "&nbsp;<img src=face/"&srs(4)&" width=14 height=14>&nbsp;<a href=dispbbs.asp?boardid="&srs(1)&"&replyid="&srs(5)&"&id="&srs(0)&"&skin=1>"
	if isnull(srs(3)) or srs(3)="" then
	response.write left(htmlencode(replace(srs(6),chr(10)," ")),26) & "..."
	else
	response.write htmlencode(srs(3))
	end if
	response.write "</a>&nbsp;--&nbsp;"&srs(2)&"<br>"
	srs.movenext
	loop
	srs.close
	set srs=nothing
%>
</td>
  </tr>
</table>
<BR>
<%if Cint(GroupSetting(28))=1 or Cint(GroupSetting(29))=1 or Cint(GroupSetting(42))=1 or Cint(GroupSetting(43))=1  or Cint(GroupSetting(30))=1  or Cint(GroupSetting(31))=1 then%>
<table class=tableborder1 cellspacing=1 cellpadding=3 align=center>
<tr><th height="25" align=left>快捷管理选项</th></tr>
<%if Cint(GroupSetting(42))=1 or Cint(GroupSetting(43))=1 or Cint(GroupSetting(28))=1 then%>
<tr><td class=tablebody1 height=25>
<B>用户管理选项</B>：   『 <a href=admin_lockuser.asp?action=lock_1&name=<%=htmlencode(username)%> title=锁定该用户不允许登陆和发言>锁定</a> | <a href=admin_lockuser.asp?action=lock_2&name=<%=htmlencode(username)%> title=屏蔽该用户在论坛的发言>屏蔽</a> | <a href=admin_lockuser.asp?action=lock_3&name=<%=htmlencode(username)%> title=解除该用户在论坛的锁定和屏蔽>解除</a> | <a href="admin_lockuser.asp?action=power&name=<%=htmlencode(username)%>" title=对用户进行分值奖励>奖励</a> | <a href="admin_lockuser.asp?action=getpermission&name=<%=htmlencode(username)%>&userid=<%=rs("userid")%>">编辑该用户论坛权限</a> 』
</td></tr>
<%end if%>
<%if Cint(GroupSetting(29))=1 then%>
<FORM METHOD=POST ACTION="admin_lockuser.asp?action=DelTopic">
<tr><td class=tablebody1 valign=middle>
<B>帖子管理选项</B>：   删除该用户&nbsp;
<input type="hidden" value="<%=suserid%>" name="SetUserID">
<input type="hidden" value="<%=htmlencode(username)%>" name="name">
<select name="delTopicDate" size=1>
<%for i=1 to 10%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>&nbsp;天内的主题&nbsp;<input type="submit" name="submit" value="执行">
</td></tr>
</FORM>
<%end if%>
<%if Cint(GroupSetting(30))=1 then%>
<tr><td class=tablebody1 valign=middle height=25>
<B>用户最后来访IP</B>：   <a href="look_ip.asp?action=lookip&ip=<%=rs("UserLastIP")%>"><%=rs("UserLastIP")%></a>&nbsp;&nbsp;点击IP查看用户来源及操作
</td></tr>
<%end if%>
</table>
<%
end if
rs.close
set rs=nothing
call activeonline()
end sub

sub mypermission()
dim myper_1,myper_2,myper_3
myper_1=false
myper_2=false
myper_3=false
set rs=conn.execute("select uc_userid from UserAccess where uc_boardid="&boardid&" and uc_userid="&userid)
if not(rs.eof and rs.bof) then
	myper_1=true
end if
if not myper_1 then
	set rs=conn.execute("select pid from BoardPermission where Boardid="&boardid&" and GroupID="&UserGroupID)
	if not(rs.eof and rs.bof) then
		myper_2=true
	end if
end if
if not(myper_1 and myper_2) then
	set rs=conn.execute("select UserGroupID from UserGroups where UserGroupID="&UserGroupID)
	if not(rs.eof and rs.bof) then
		myper_3=true
	end if
end if
set rs=nothing
%>
<table class=tableborder1 cellspacing=1 cellpadding=3  align=center>
<tr> 
<td height="25" colspan="2" class=tablebody1>
<%
dim myper_4
if boardid=0 then
	myper_4="本论坛"
else
	myper_4=boardtype
end if
if myper_1 then
	response.write "您在 <B>"&myper_4&"</B> 被系统定义了单独的用户权限，具体的权限列表如下"
elseif myper_2 then
	response.write "您所属的用户组在 <B>"&myper_4&"</B> 有单独的权限设置，具体的权限列表如下"
else
	response.write "您所属的用户组在 <B>"&myper_4&"</B> 使用的是默认设置，具体的权限列表如下"
end if
%>
</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝查看权限</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以浏览论坛</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(0)=1 then%>√<%end if%><%if GroupSetting(0)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以查看会员信息(包括其他会员的资料和会员列表)
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(1)=1 then%>√<%end if%><%if GroupSetting(1)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以查看其他人发布的主题
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(2)=1 then%>√<%end if%><%if GroupSetting(2)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以浏览精华帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(41)=1 then%>√<%end if%><%if GroupSetting(41)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>发帖权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以发布新主题</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(3)=1 then%>√<%end if%><%if GroupSetting(3)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以回复自己的主题
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(4)=1 then%>√<%end if%><%if GroupSetting(4)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以回复其他人的主题
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(5)=1 then%>√<%end if%><%if GroupSetting(5)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以在论坛允许评分的时候参与评分(鲜花和鸡蛋)?
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(6)=1 then%>√<%end if%><%if GroupSetting(6)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>参与评分所需金钱
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(47)%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以上传附件
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(7)=1 then%>√<%end if%><%if GroupSetting(7)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>最多上传文件个数
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(40)%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>上传文件大小限制
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(44)%></B> KB</td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以发布新投票</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(8)=1 then%>√<%end if%><%if GroupSetting(8)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以参与投票</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(9)=1 then%>√<%end if%><%if GroupSetting(9)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以发布小字报</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(17)=1 then%>√<%end if%><%if GroupSetting(17)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>发布小字报所需金钱</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(46)%></B></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>帖子/主题编辑权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以编辑自己的帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(10)=1 then%>√<%end if%><%if GroupSetting(10)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以删除自己的帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(11)=1 then%>√<%end if%><%if GroupSetting(11)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以移动自己的帖子到其他论坛
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(12)=1 then%>√<%end if%><%if GroupSetting(12)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以打开/关闭自己发布的主题
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(13)=1 then%>√<%end if%><%if GroupSetting(13)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝<b>其他权限</b></th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以搜索论坛
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(14)=1 then%>√<%end if%><%if GroupSetting(14)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以使用'发送本页给好友'功能
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(15)=1 then%>√<%end if%><%if GroupSetting(15)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以修改个人资料
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(16)=1 then%>√<%end if%><%if GroupSetting(16)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>帖子中可以查看其它人头衔
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(36)=1 then%>√<%end if%><%if GroupSetting(36)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>帖子中可以查看其它人头像
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(37)=1 then%>√<%end if%><%if GroupSetting(37)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>帖子中可以可以查看其它人签名
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(38)=1 then%>√<%end if%><%if GroupSetting(38)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以浏览论坛事件
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(39)=1 then%>√<%end if%><%if GroupSetting(39)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝管理权限</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以删除其它人帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(18)=1 then%>√<%end if%><%if GroupSetting(18)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以移动其它人帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(19)=1 then%>√<%end if%><%if GroupSetting(19)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以打开/关闭其它人帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(20)=1 then%>√<%end if%><%if GroupSetting(20)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以固顶/解除固顶帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(21)=1 then%>√<%end if%><%if GroupSetting(21)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以奖励/惩罚发贴用户
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(22)=1 then%>√<%end if%><%if GroupSetting(22)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以奖励/惩罚用户
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(43)=1 then%>√<%end if%><%if GroupSetting(43)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以编辑其它人帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(23)=1 then%>√<%end if%><%if GroupSetting(23)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以加入/解除精华帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(24)=1 then%>√<%end if%><%if GroupSetting(24)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以发布公告
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(25)=1 then%>√<%end if%><%if GroupSetting(25)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以管理公告
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(26)=1 then%>√<%end if%><%if GroupSetting(26)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以管理小字报
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(27)=1 then%>√<%end if%><%if GroupSetting(27)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以锁定/屏蔽/解除锁定用户
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(28)=1 then%>√<%end if%><%if GroupSetting(28)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以删除用户1－10天内所发帖子
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(29)=1 then%>√<%end if%><%if GroupSetting(29)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以查看来访IP及来源
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(30)=1 then%>√<%end if%><%if GroupSetting(30)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以限定IP来访
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(31)=1 then%>√<%end if%><%if GroupSetting(31)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以管理用户权限
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(42)=1 then%>√<%end if%><%if GroupSetting(42)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以批量删除帖子（前台）
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(45)=1 then%>√<%end if%><%if GroupSetting(45)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>＝＝短信权限</th>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>可以发送短信
</td>
<td height="23" width="40%" class=tablebody1><B><%if GroupSetting(32)=1 then%>√<%end if%><%if GroupSetting(32)=0 then%><font color=<%=Forum_body(8)%>>×</font><%end if%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>最多发送用户
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(33)%></B></td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>短信内容大小限制
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(34)%></B> byte</td>
</tr>
<tr>
<td height="23" width="60%" class=tablebody1>信箱大小限制
</td>
<td height="23" width="40%" class=tablebody1><B><%=GroupSetting(35)%></B> KB</td>
</tr>
</table>
<%
end sub
%>