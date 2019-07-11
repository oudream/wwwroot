<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	stats="论坛搜索"
	if boardmaster or master then
		guestlist=""
	else
		guestlist=" lockboard<>1 and "
	end if
	call nav()
	if BoardID="" or not isInteger(BoardID) then
		boardid=0
		call head_var(1)
	else
		boardid=cLng(Boardid)
		call head_var(2)
	end if
	call main()
	if founderr then call dvbbs_error()
	call footer()

sub main()
	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛搜索的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
		founderr=true
		exit sub
	end if
%>
<form action=queryresult.asp method=post>
    <table cellpadding=5 cellspacing=1 align=center class=tableborder1>
    	<tr>
	<th valign=middle colspan=2 >请输入要搜索的关键字</th></tr>
	<tr>
	<td class=tablebody1 colspan=2 align=center valign=middle>(同时查询多个条件使用'<font  color=<%=Forum_body(8)%>>or</font>' 分隔关键字，查询同时满足某条件使用'<font  color=<%=Forum_body(8)%>>and</font>'分隔关键字)<br><br><input type=text size=40 name=keyword></td></tr>
           <tr>
	<td class=tablebody2 valign=middle colspan=2 align=center><b>搜索选项</b></td></tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>作者搜索</b>&nbsp;<input name=sType type=radio value=1>
           </td>
           <td class=tablebody1 align=left valign=middle>
           <select name=nSearch>
                <option value=1>搜索主题作者
                 <option value=2>搜索回复作者
                  <option value=3>两者都搜索
           </select>
	</td>
           </tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>关键字搜索</b>&nbsp;<input name=sType type=radio value=2 checked>
           </td>
           <td class=tablebody1 align=left valign=middle>
           <select name=pSearch>
	   <option value=1>在主题中搜索关键字
	   <!--<option value=2>在贴子内容中搜索关键字
	       <option value=3>两者都搜索-->
	</select>
	</td>
           </tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>日期范围</b>
           </td>
           <td class=tablebody1 align=left valign=middle>
<select name=SearchDate class=smallsel> <option value=ALL>所有日期<option value=1>昨天以来<option  selected value=5>5天以来<option value=10>10天以来<option value=30>30天以来</option></select> 
           <select name=Stable size=1>
<%
For i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""""
	if trim(NowUseBBS)=trim(AllPostTable(i)) then response.write " selected "
	response.write ">"&AllPostTablename(i)&"</option>"
next
%>
           </select>
	 </td>
           </tr>
           <tr>
	<td class=tablebody1 valign=middle colspan=2 align=center><b>请选择要搜索的论坛 (不要选那些用 >> 和 << 括起来的，那只是类别名，不是论坛)</b></td></tr>
	<tr>
	<td class=tablebody1 colspan=2 valign=middle align=center>
           <b>搜索论坛： &nbsp; 
<select name=boardid size=1>
<option value=0>所有论坛</option>
<%
	dim rs1,sql1
	dim searchboard,searchboarda,searchclass,searchinfo
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option value=0>>>&nbsp;"& rs(1) &" <<"
		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			searchboard="<option value=0>没有论坛"
		else
			do while not rs1.eof
				response.write "<option value="& rs1(0) &""
				if boardid<>"" then
				if cint(boardid)=cint(rs1(0)) then response.write " selected "
				end if
				response.write ">"& rs1(1)
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
	set rs=nothing
	set rs1=nothing
%>
</select>
	</b></td>
	</tr>
	<tr>
	<td class=tablebody2 valign=middle colspan=2 align=center>
	<input type=submit value=开始搜索>
	</td></form></tr></table>
<%
	call activeonline()
end sub
%>