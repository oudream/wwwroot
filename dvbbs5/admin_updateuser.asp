<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char.asp" -->
<!--#include file="inc/grade.asp"-->
<title><%=Forum_info(0)%>--管理页面</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	server.scriptTimeout=999999
	if not master or instr(session("flag"),"15")=0 then
		Errmsg=Errmsg+"<br>"+"<li>本页面为管理员专用，请<a href=admin_index.asp target=_top>登陆</a>后进入。<br><li>您没有管理本页面的权限。"
		call dvbbs_error()
	else
		dim body
		dim tmprs
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td align=center colspan="2"><font color="<%=Forum_body(6)%>">欢迎<b><%=membername%></b>进入管理页面</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><p><font color="<%=Forum_body(7)%>">
<%
	if request("action")="updat" then
		if request("submit")="重新计算用户发贴" then
		call updateTopic()
		elseif request("submit")="更新用户等级" then
		call updategrade()
		else
		call updatemoney()
		end if
		response.write ""&body&""
	else
%>
              <table width="100%" border="0" cellspacing="3" cellpadding="0">
                <tr> 
                  <td bgcolor=<%=Forum_body(3)%>> <font color="<%=Forum_body(7)%>">
                    <p><b>更新用户数据</b>：执行下列操作比较消耗服务器资源，如果数据量非常大的情况下将消耗更多的资源和时间，尽量减少使用次数。</p></font>
                  </td>
                </tr>
                <tr> 
                  <td>
                    <p><b>填写更新用户ID范围</b>：填1000时，即更新ID号1000以下的用户，2000时；即更新ID号1000-2000以内的用户。一般您论坛的最大用户ID号为当前论坛总用户数量左右。每次更新数值最好在1000以内。</p>
                  </td>
                </tr>
                <tr> 
                  <td> <font color="<%=Forum_body(7)%>">
<form action="admin_updateuser.asp?action=updat" method=post>
 请填写更新用户ID范围：<input type="text" name="useridx" size="20"> <input type="submit" name="Submit" value="重新计算用户发贴">&nbsp;<BR><BR>执行本操作将按照论坛发贴重新计算所有用户发表帖子数量。<BR><BR>
 请填写更新用户ID范围：<input type="text" name="useridy" size="20"> <input type="submit" name="Submit" value="更新用户等级">&nbsp;<BR><BR>执行本操作将按照用户发贴数量和论坛的等级设置重新计算用户等级，本操作不影响等级为贵宾、版主、总版主的数据。<BR><BR>
 请填写更新用户ID范围：<input type="text" name="useridz" size="20"> <input type="submit" name="Submit" value="更新用户金钱/经验/魅力">&nbsp;<BR><BR>执行本操作将按照用户的发贴数量和论坛的相关设置重新计算用户的金钱/经验/魅力，本操作也将重新计算贵宾、版主、总版主的数据，注意：<font color=red>不推荐用户进行本操作，本操作在数据很多的时候请尽量不要使用，并且本操作对各个版面删除帖子等所扣相应分值不做运算，只是按照发贴和总的论坛分值设置进行运算，请大家慎重操作</font>。
</form></font>
                  </td>
                </tr>
<%
	end if
%></font>
</p></td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
'	response.write ""&body&""
	end sub

	sub updateTopic()

	dim userTopic
        dim useridx,userida,useridb
        useridx=request("useridx")
    if useridx="" then
          useridb=""
          
       else
          userida=useridx-1000
          IF USERIDA<="0" THEN
          USERIDA=0
          END IF

          useridb="where userid>"&userida&" and userid<"&useridx&""
           
    end if
	'conn.execute("update [user] set article="&allnum(username)&"")
	set rs=server.createobject("adodb.recordset")
	sql="select username,userid from [user] "&useridb&" order by userid desc"
	rs.open sql,conn,1,1
	do while not rs.eof
		userTopic=allnum(rs(0))
		conn.execute("update [user] set article="&userTopic&" where userid="&rs(1))
	rs.movenext
	loop
	rs.close
	set rs=nothing
	body=body+"<br>"+"更新用户数据成功("&now()&")。"
	end sub

	sub updatemoney()
	dim userTopic,userReply,userWealth
	dim userEP,userCP

       dim useridy,userida,useridb
        useridy=request("useridy")
    if useridy="" then
          useridb=""
          
       else
          userida=useridy-1000
          IF USERIDA<="0" THEN
          USERIDA=0
          END IF

          useridb="where userid>"&userida&" and userid<"&useridy&""
           
    end if
	set rs=server.createobject("adodb.recordset")
	sql="select username,logins,userid from [user] "&useridb&" order by userid desc"
	rs.open sql,conn,1,1
	do while not rs.eof
		userTopic=TopicNum(rs(0))
		userreply=replyNum(rs(0))
		userwealth=rs(1)*Forum_user(4) + userTopic*Forum_user(1) + userreply*Forum_user(2)
		userEP=rs(1)*Forum_user(9) + userTopic*Forum_user(6) + userreply*Forum_user(7)
		userCP=rs(1)*Forum_user(14) + userTopic*Forum_user(11) + userreply*Forum_user(12)
		conn.execute("update [user] set userWealth="&userWealth&",userep="&userep&",usercp="&usercp&" where userid="&rs(2))
	rs.movenext
	loop
	rs.close
	set rs=nothing
	body=body+"<br>"+"更新用户数据成功("&now()&")。"
	end sub

	sub updategrade()
        dim useridz,userida,useridb
        useridz=request("useridz")
    if useridz="" then
          useridb=""
          
       else
          userida=useridz-1000
          IF USERIDA<="0" THEN
          USERIDA=0
          END IF

          useridb=" userid>"&userida&" and userid<"&useridz&" and "
           
    end if
	conn.execute("update [user] set userclass=1 where "&useridb&" article<0")

	conn.execute("update [user] set userclass=1 where "&useridb&" (article>="&point(1)&" and article<"&point(2)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=2 where "&useridb&" (article>="&point(2)&" and article<"&point(3)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=3 where "&useridb&" (article>="&point(3)&" and article<"&point(4)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=4 where "&useridb&" (article>="&point(4)&" and article<"&point(5)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=5 where "&useridb&" (article>="&point(5)&" and article<"&point(6)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=6 where "&useridb&" (article>="&point(6)&" and article<"&point(7)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=7 where "&useridb&" (article>="&point(7)&" and article<"&point(8)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=8 where "&useridb&" (article>="&point(8)&" and article<"&point(9)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=9 where "&useridb&" (article>="&point(9)&" and article<"&point(10)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=10 where "&useridb&" (article>="&point(10)&" and article<"&point(11)&") and not (userclass=18 or userclass=19 or  userclass=20)")	
	conn.execute("update [user] set userclass=11 where "&useridb&" (article>="&point(11)&" and article<"&point(12)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=12 where "&useridb&" (article>="&point(12)&" and article<"&point(13)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=13 where "&useridb&" (article>="&point(13)&" and article<"&point(14)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=14 where "&useridb&" (article>="&point(14)&" and article<"&point(15)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=15 where "&useridb&" (article>="&point(15)&" and article<"&point(16)&") and not (userclass=18 or userclass=19 or  userclass=20)")
	conn.execute("update [user] set userclass=16 where "&useridb&" (article>="&point(16)&" and article<"&point(17)&") and not (userclass=18 or userclass=19 or  userclass=20)")

	conn.execute("update [user] set userclass=17 where "&useridb&" article>="&point(17)&" and not (userclass=18 or userclass=19 or  userclass=20)")
	body=body+"<br>"+"更新用户数据成功("&now()&")。"
	end sub

	function TopicNum(username)
	set tmprs=conn.execute("select count(announceid) from bbs1 where ParentID=0 and username='"&username&"'")
	TopicNum=tmprs(0)
	if isnull(TopicNum) then TopicNum=0
	set tmprs=nothing
	end function
	function replyNum(username)
	set tmprs=conn.execute("select count(announceid) from bbs1 where not ParentID=0 and username='"&username&"'")
	replyNum=tmprs(0)
	if isnull(replyNum) then replyNum=0
	set tmprs=nothing
	end function
	function allnum(username)
	set tmprs=conn.execute("select count(announceid) from bbs1 where username='"&username&"'")
	allnum=tmprs(0)
	if isnull(allnum) then allnum=0
	set tmprs=nothing
	end function
%>