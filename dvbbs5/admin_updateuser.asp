<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char.asp" -->
<!--#include file="inc/grade.asp"-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY <%=Forum_body(11)%>>
<%
	server.scriptTimeout=999999
	if not master or instr(session("flag"),"15")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
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
        <td align=center colspan="2"><font color="<%=Forum_body(6)%>">��ӭ<b><%=membername%></b>�������ҳ��</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><p><font color="<%=Forum_body(7)%>">
<%
	if request("action")="updat" then
		if request("submit")="���¼����û�����" then
		call updateTopic()
		elseif request("submit")="�����û��ȼ�" then
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
                    <p><b>�����û�����</b>��ִ�����в����Ƚ����ķ�������Դ������������ǳ��������½����ĸ������Դ��ʱ�䣬��������ʹ�ô�����</p></font>
                  </td>
                </tr>
                <tr> 
                  <td>
                    <p><b>��д�����û�ID��Χ</b>����1000ʱ��������ID��1000���µ��û���2000ʱ��������ID��1000-2000���ڵ��û���һ������̳������û�ID��Ϊ��ǰ��̳���û��������ҡ�ÿ�θ�����ֵ�����1000���ڡ�</p>
                  </td>
                </tr>
                <tr> 
                  <td> <font color="<%=Forum_body(7)%>">
<form action="admin_updateuser.asp?action=updat" method=post>
 ����д�����û�ID��Χ��<input type="text" name="useridx" size="20"> <input type="submit" name="Submit" value="���¼����û�����">&nbsp;<BR><BR>ִ�б�������������̳�������¼��������û���������������<BR><BR>
 ����д�����û�ID��Χ��<input type="text" name="useridy" size="20"> <input type="submit" name="Submit" value="�����û��ȼ�">&nbsp;<BR><BR>ִ�б������������û�������������̳�ĵȼ��������¼����û��ȼ�����������Ӱ��ȼ�Ϊ������������ܰ��������ݡ�<BR><BR>
 ����д�����û�ID��Χ��<input type="text" name="useridz" size="20"> <input type="submit" name="Submit" value="�����û���Ǯ/����/����">&nbsp;<BR><BR>ִ�б������������û��ķ�����������̳������������¼����û��Ľ�Ǯ/����/������������Ҳ�����¼��������������ܰ��������ݣ�ע�⣺<font color=red>���Ƽ��û����б������������������ݺܶ��ʱ���뾡����Ҫʹ�ã����ұ������Ը�������ɾ�����ӵ�������Ӧ��ֵ�������㣬ֻ�ǰ��շ������ܵ���̳��ֵ���ý������㣬�������ز���</font>��
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
	body=body+"<br>"+"�����û����ݳɹ�("&now()&")��"
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
	body=body+"<br>"+"�����û����ݳɹ�("&now()&")��"
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
	body=body+"<br>"+"�����û����ݳɹ�("&now()&")��"
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