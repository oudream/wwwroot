<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
   	dim bBoardEmpty
	dim totalrec
	dim n,RowCount
	dim p
	dim currentpage,page_count,Pcount
	dim tablename
	if not master then
		Errmsg=Errmsg+"<br><li>您没有权限浏览本页。"
		founderr=true
	end if

	stats="论坛回收站"
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if

	call nav()
	call head_var(1)
	if founderr then
	call dvbbs_error()
	else
	call AnnounceList1()
	call listPages3()
	if founderr then call dvbbs_error()
	end if
	call activeonline()
	call footer()

	sub AnnounceList1()
response.write "<table cellpadding=0 cellspacing=0 border=0 width="&Forum_body(12)&" align=center><tr>"&_
			"<td align=center width=2 valign=middle> </td>"&_
			"<td align=left valign=middle> 本页面只有系统管理员可进行操作，请选择需要的操作，列出所有：<a href=?tablename=topic>主题表数据</a>"
for i=0 to ubound(allposttable)
response.write " | <a href=?tablename="&allposttable(i)&">帖子"&allposttablename(i)&"</a>"
next
response.write "<BR>注意：还原主题表数据将连跟贴（帖子）表数据一起还原，删除主题表数据将连跟帖（帖子）表一起删除</td>"&_
			"<td align=right> </td></tr></table><BR>"

	if instr(request("tablename"),"bbs")>0 then
	sql="select AnnounceID,boardID,UserName,Topic,body,DateAndTime from "&request("tablename")&" where locktopic=2 and not parentid=0 order by announceid desc"
	tablename=request("tablename")
	else
	sql="select topicID,boardID,PostUserName,Title,title as body,DateAndTime from topic where locktopic=2 order by topicid desc"
	tablename="topic"
	end if
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
		'论坛无内容
		call showEmptyBoard1()
	else
		rs.PageSize = cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
      	totalrec=rs.recordcount
		call showpagelist1() 
	end if
	end sub


	REM 显示贴子列表	
	sub showPageList1()
	dim body
	dim vrs
	dim votenum,votenum_1
	dim pnum
	i=0
response.write "<form name=recycle action=admin_recycle.asp method=post><input type=hidden value="&tablename&" name=tablename>"&_
				"<TABLE cellPadding=1 cellSpacing=1 align=center class=tableborder1>"&_
  				"<TBODY>"&_
				"<TR align=middle>"&_
				"<Th height=25 width=32>状态</Th>"&_ 
				"<Th width=*>主 题</Th>"&_
				"<Th width=80>作 者</Th>"&_
				"<Th width=195>最后更新 | 回复人</Th>"&_
				"</TR>"
		while (not rs.eof) and (not page_count = rs.PageSize)
response.write "<TR align=middle>"&_
				"<TD class=tablebody2 width=32 height=27>"

response.write "<input type=checkbox name=topicid value="&rs(0)&">"

response.write "</TD>"&_
				"<TD align=left class=tablebody1 width=*>"

	body=replace(replace(htmlencode(left(rs(4),20)),"<BR>",""),"</P><P>","")
	
response.write "<a href=""javascript:openScript('viewrecycle.asp?id="&rs(0)&"&tablename="&tablename&"',500,400)"">"
if rs(3)="" or isnull(rs(3)) then
	response.write body
else
	if len(rs(3))>26 then
	response.write ""&left(htmlencode(rs(3)),26)&"..."
	else
	response.write htmlencode(rs(3))
	end if
end if
	response.write "</a>"

response.write "</TD>"&_
				"<TD class=tablebody2 width=80><a href=dispuser.asp?name="& rs(2) &">"& rs(2) &"</a></TD>"

response.write "<TD align=left class=tablebody1 width=195>&nbsp;"

	'on error resume next
		response.write "&nbsp;"&_
						FormatDateTime(rs(5),2)&"&nbsp;"&FormatDateTime(rs(5),4)&_
						"&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;------"

response.write "</TD></TR>"

	  page_count = page_count + 1
          rs.movenext
        wend
	end sub


	sub listPages3()
	dim endpage
	'on error resume next
	Pcount=rs.PageCount
	if master and totalrec>0 then
response.write "<TR><TD height=27 width=""100%"" class=tablebody2 colspan=4><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">选中所有显示帖子&nbsp;<input type=submit name=action onclick=""{if(confirm('确定还原选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=还原>"
	if session("flag")<>"" then
	response.write "&nbsp;<input type=submit name=action onclick=""{if(confirm('确定删除选定的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=删除>&nbsp;<input type=submit name=action onclick=""{if(confirm('确定清除回收站所有的纪录吗?')){this.document.recycle.submit();return true;}return false;}"" value=清空回收站>"
	end if
	response.write "</TD></TR></form>"
	end if
	response.write "</TBODY></TABLE>"
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"页次：<b>"&currentpage&"</b>/<b>"&Pcount&"</b>页"&_
			"每页<b>"&Forum_Setting(11)&"</b> 贴数<b>"&totalrec&"</b></td>"&_
			"<td valign=middle nowrap><div align=right><p>分页： <b>"

	if currentpage > 4 then
	response.write "<a href=""?page=1"">[1]</a> ..."
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
        response.write " <a href=""?page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a></b>"
	end if
	response.write "</p></div></td></tr></table>"
	rs.close
	set rs=nothing
	end sub 

	sub showEmptyBoard1()
		response.write "<TABLE class=tableborder1 cellPadding=4 cellSpacing=1 align=center>"&_
			"<TBODY>"&_
			"<TR align=middle>"&_
			"<Th height=25>状态</Th>"&_
			"<Th>主 题</Th>"&_
			"<Th>作 者</Th> "&_
			"<Th>回复/人气</Th> "&_
			"<Th>最新回复</Th></TR> "&_
			"<tr><td colSpan=5 width=100% class=tablebody1>论坛回收站暂无内容。</td></tr>"&_
			"</TBODY></TABLE>"
	end sub
%>
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>