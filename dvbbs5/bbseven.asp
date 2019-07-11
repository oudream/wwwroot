<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	dim isdisp,bbseveninfo
	dim endpage
	dim totalrec
	dim n
	dim currentpage,page_count,Pcount
	dim bgcolor
	if boardid=0 then
	stats="论坛总事件记录"
	call nav()
	call head_var(1)
	else
	stats=BoardType & "事件记录"
	call nav()
	call head_var(2)
	end if
	if Cint(GroupSetting(39))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有浏览本论坛事件的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
	end if
	if master then founderr=false
	if founderr then
		call dvbbs_error()
	else
		if request("action")="dellog" then
			call batch()
		else
			call boardeven()
		end if
		if founderr then call dvbbs_error()
		call activeonline()
	end if
	call footer()
	REM 显示版面信息---Headinfo
	sub boardeven()
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
	if master then
	response.write "<div align=center>版主或管理员请点击操作时间切换到管理状态</div>"
	end if

	response.write "<form action=bbseven.asp?action=dellog method=post name=even>"
	response.write "<input type=hidden name=boardid value="&boardid&">"
	response.write "<table cellspacing=1 cellpadding=3 align=center class=tableborder1>"
	response.write "<tr align=center>"
	response.write "<th width=15% height=25>对象</td>"
	response.write "<th width=50% id=tabletitlelink>事件内容(<a href=bbseven.asp>查看所有事件记录</a>)</td>"
	response.write "<th width=20% id=tabletitlelink><a href=bbseven.asp?action=batch&boardid="&boardid&">操作时间</a></td>"
	response.write "<th width=15% >操作人</td></tr>"


	set rs=server.createobject("adodb.recordset")
	if boardid>0 then
	sql="select * from log where l_boardid="&boardid&" order by l_addtime desc"
	else
	sql="select * from log order by l_addtime desc"
	end if
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		response.write "<tr><td class=tablebody1 colspan=4 height=25>本版还没有任何事件</td></tr>"
	else
		rs.PageSize = Forum_Setting(11)
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = rs.PageSize)

		if bgcolor=Forum_body(4) then
		bgcolor=Forum_body(5)
		else
		bgcolor=Forum_body(4)
		end if

		response.write "<tr>"
		response.write "<td class=tablebody1 align=center height=24><a href=dispuser.asp?name="&htmlencode(rs("l_touser"))&" target=_blank>"&htmlencode(rs("l_touser"))&"</a></td>"
		response.write "<td class=tablebody1>"&htmlencode(rs("l_content"))&"</td>"
		response.write "<td class=tablebody1>"
		if request("action")="batch" and (master or boardmaster) then
		response.write "<input type=checkbox name=lid value="&rs("l_id")&">"
		end if
		response.write rs("l_addtime")
		response.write "</td>"
		response.write "<td align=center class=tablebody1>"
		if Forum_Setting(72)=1 or master then
		response.write "<a href=dispuser.asp?name="&htmlencode(rs("l_username"))&" target=_blank>"&htmlencode(rs("l_username"))&"</a>"
		else
		response.write "保密"
		end if
		response.write "</td></tr>"
		page_count = page_count + 1
		rs.movenext
		wend
	end if
	if request("action")="batch" then
		response.write "<tr><td class=tablebody2 colspan=4>请选择要删除的事件，<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">全选 <input type=submit name=Submit value=执行  onclick=""{if(confirm('您确定执行的操作吗?')){this.document.even.submit();return true;}return false;}""></td></tr>"
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
	response.write "<a href=""?page=1&boardid="&boardid&""">[1]</a> ..."
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
		response.write " <a href=""?page="&i&"&boardid="&boardid&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&boardid="&boardid&""">["&Pcount&"]</a>"
	end if
	response.write "</td></tr></table>"
	rs.close
	set rs=nothing
	end sub

	sub batch()
	dim lid
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请登陆后进行操作。"
	end if
	if boardid=0 then
		if not master then
			founderr=true
			Errmsg=Errmsg+"<br>"+"<li>您不是系统管理员，不能管理所有日志。"
		end if
	else
		if not boardmaster then
			Errmsg=Errmsg+"<br>"+"<li>您不是该版面斑竹或者系统管理员。<br><li>或者您没有使用该功能的权限。"
			founderr=true
		end if
	end if

	if request.form("lid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>请指定相关事件。"
	else
		lid=replace(request.Form("lid"),"'","")
	end if
	if founderr then exit sub
	conn.execute("delete from log where l_id in ("&lid&")")
	sucmsg="<li>删除指定事件成功"
	call dvbbs_suc()
	end sub
%>