<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'=========================================================
' File: allpaper.asp
' Version:5.0
' Date: 2002-9-10
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

if boardid=0 then
	stats="��̳ȫ��С�ֱ�"
	call nav()
	call head_var(1)
else
	stats=BoardType & "С�ֱ�"
	call nav()
	call head_var(2)
end if
if Forum_Setting(56)=0 then
	Errmsg=Errmsg+"<br>"+"<li>�ð���δ����С�ӱ���������������棡"
	Founderr=true
end if

if founderr then
	call dvbbs_error()
else
	if request("action")="delpaper" then
		call batch()
	else
		call boardeven()
	end if
	if founderr then call dvbbs_error()
	call activeonline()
end if
call footer()

sub boardeven()
	dim totalrec
	dim n
	dim currentpage,page_count,Pcount
	currentPage=request("page")
	if currentpage="" or not isInteger(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
	end if
if Cint(GroupSetting(27))=1 then
%>
<div align=center>���������л�������״̬</div>
<%end if%>
<form action=allpaper.asp?action=delpaper method=post name=paper>
<input type=hidden name=boardid value="<%=boardid%>">
<table class=tableborder1 cellspacing="1" cellpadding="3" align="center">
  <tr align=center> 
    <th width="15%" height=25>�û�</th>
    <th width="50%">����</th>
    <th width="20%">����ʱ��</th>
    <th width="15%" id=tabletitlelink><%if Cint(GroupSetting(27))=1 then%><a href="allpaper.asp?action=batch&boardid=<%=boardid%>"><%end if%>����</a></th>
  </tr> 
<%
	dim bgcolor
	set rs=server.createobject("adodb.recordset")
	if boardid=0 then
	sql="select * from smallpaper order by s_addtime desc"
	else
	sql="select * from smallpaper where s_boardid="&boardid&" order by s_addtime desc"
	end if
	rs.open sql,conn,1,1
	if rs.bof and rs.eof then
	response.write "<tr> <td class=tablebody1 colspan=4 height=25>���滹û��С�ֱ�</td></tr>"
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

		response.write "<tr> "
		response.write "<td class=tablebody1 align=center height=24>"
		response.write "<a href=dispuser.asp?name="&htmlencode(rs("s_username"))&" target=_blank>"&htmlencode(rs("s_username"))&"</a>"
		response.write "</td>"
		response.write "<td class=tablebody1>"
		response.write "<a href=javascript:openScript('viewpaper.asp?id="&rs("s_id")&"&boardid="&boardid&"',500,400)>"&htmlencode(rs("s_title"))&"</a>"
		response.write "</td>"
		response.write "<td class=tablebody1>"&rs("s_addtime")&"</td>"
		response.write "<td align=center class=tablebody1>"
		if request("action")="batch" then
		response.write "<input type=checkbox name=sid value="&rs("s_id")&">"
		else
		response.write rs("s_hits")
		end if
		response.write "</td></tr>"
	  	page_count = page_count + 1
		rs.movenext
		wend
	if request("action")="batch" then
	response.write "<tr><td class=tablebody2 colspan=4 align=right>��ѡ��Ҫɾ����С�ֱ���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=Submit value=ִ��  onclick=""{if(confirm('��ȷ��ִ�еĲ�����?')){this.document.paper.submit();return true;}return false;}""></td></tr>"
	end if
	response.write "</table>"
	dim endpage
	Pcount=rs.PageCount
	response.write "<table border=0 cellpadding=0 cellspacing=3 width="""&Forum_body(12)&""" align=center>"&_
			"<tr><td valign=middle nowrap>"&_
			"ҳ�Σ�<b>"&currentpage&"</b>/<b>"&Pcount&"</b>ҳ"&_
			"ÿҳ<b>"&Forum_Setting(11)&"</b> ����<b>"&totalrec&"</b>��</td>"&_
			"<td valign=middle nowrap><div align=right><p>��ҳ�� "

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
	response.write "</font></td></tr>"
	rs.close
	set rs=nothing	
	end if

	response.write "</table></form>"

	end sub

	sub batch()
	dim sid
	dim adminpaper
	adminpaper=false
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���½����в�����"
	end if
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(27))=1 then
		adminpaper=true
	else
		adminpaper=false
	end if
	if UserGroupID>3 and Cint(GroupSetting(27))=1 then
		adminpaper=true
	end if
	if FoundUserPer and Cint(GroupSetting(27))=1 then
		adminpaper=true
	elseif FoundUserPer and Cint(GroupSetting(27))=0 then
		adminpaper=false
	end if
	if not adminpaper then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��û��Ȩ��ִ�б�������"
	end if
	if request.form("sid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ�����С�ֱ���"
	else
		sid=replace(request.Form("sid"),"'","")
	end if
	if founderr then exit sub
	conn.execute("delete from smallpaper where s_boardid="&boardid&" and s_id in ("&sid&")")
	sucmsg="<li>��ѡ���С�ֱ��Ѿ�ɾ����"
	call dvbbs_suc()
	end sub
%>