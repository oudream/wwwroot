<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file=inc/char.asp-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		call main()
	end if

sub main()
dim userip,ips,GetIp1,GetIp2
if request("userip")<>"" then
userip=request("userip")
ips=Split(userIP,".")
GetIp1=ips(0)&"."&ips(1)&"."&ips(2)&".1"
GetIp2=ips(0)&"."&ips(1)&"."&ips(2)&".255"
else 
userip=""
GetIp1=""
GetIp2=""
end if
if request("action")="add" then
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=2>IP���ƹ��������</th>
</tr>
<%
dim sip,str1,str2,str3,str4,num_1,num_2
if request.querystring("reaction")="save" then
	sip=cstr(request.form("ip1"))
	'dot=instr(ip,".")-1
	'response.write dot
	str1=left(sip,cint(instr(sip,".")-1))
	sip=mid(sip,cint(instr(sip,"."))+1)
	str2=left(sip,cint(instr(sip,"."))-1)
	sip=mid(sip,cint(instr(sip,"."))+1)
	str3=left(sip,cint(instr(sip,"."))-1)
	str4=mid(sip,cint(instr(sip,"."))+1)
	num_1=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1

	sip=cstr(request.form("ip2"))
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	num_2=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
	'response.write num_1 &","& num_2
	'response.end

	set rs = server.CreateObject ("adodb.recordset")
	sql="select * from LockIP"
	rs.open sql,conn,1,3
	rs.addnew
	rs("ip1")=num_1
	rs("ip2")=num_2
	rs("sip1")=request.form("ip1")
	rs("sip2")=request.form("ip2")
	rs.update
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<tr>
<td width="100%" colspan=2 class=forumrow>��ӳɹ���</td>
</tr>
<%
else
%>
<form action="admin_LockIP.asp?action=add&reaction=save" method="post">
<tr>
<td width="20%" class=forumrow>��ʼI&nbsp;P</td>
<td width="80%" class=forumrow><input type="text" name="ip1" size="30" value="<%=GetIp1%>">&nbsp;��202.152.12.1</td>
</tr>
<tr>
<td width="20%" class=forumrow>��βI&nbsp;P</td>
<td width="80%" class=forumrow><input type="text" name="ip2" size="30" value="<%=GetIp2%>">&nbsp;��202.152.12.255</td>
</tr>
<tr>
<td width="20%" class=forumrow></td>
<td width="80%" class=forumrow>
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<%
end if
elseif request("action")="delip" then
	conn.execute("delete from lockip where id in ("&request.form("delid")&")")
	response.write "ɾ���ɹ���"
else
%>
<FORM METHOD=POST ACTION="?action=delip">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<th width="100%" colspan=4>IP���ƹ���������</th>
</tr>
<tr>
<td width="15%" align="center" height="20" class="forumHeaderBackgroundAlternate"><B>ID��</B></td>
<td width="35%" align="center" class="forumHeaderBackgroundAlternate"><B>��ʼIP</B></td>
<td width="35%" align="center" class="forumHeaderBackgroundAlternate"><B>��βIP</B></td>
<td width="15%" align="center" class="forumHeaderBackgroundAlternate"><B>����</B></td>
</tr>
<%
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
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
	sql="select id,sip1,sip2 from LockIP order by id desc"
	Set rs= Server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
<tr> 
<td width="100%" colspan=4 class=forumrow>��û���κ�IP�������ݡ�</td>
</tr>
<%
	else
		rs.PageSize = Cint(Forum_Setting(11))
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))
%>
<tr>
<td width="15%" align="center" height="20" class=forumrow><%=rs("id")%></td>
<td width="35%" align="center" class=forumrow><%=rs("sip1")%></td>
<td width="35%" align="center" class=forumrow><%=rs("sip2")%></td>
<td width="15%" align="center" class=forumrow><input type=checkbox name="delid" value="<%=cstr(rs("ID"))%>"></td>
</tr>
<%
		page_count = page_count + 1
		rs.movenext
		wend
%>
<tr><td colspan=3 class=forumrow align=center>��ҳ��
<%Pcount=rs.PageCount
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
        response.write " <font color=red>["&i&"]</font>"
		else
        response.write " <a href=""?page="&i&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
	end if
%>
</td>
<td align=center class=forumrow><input type=submit name=submit value="ɾ��"><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)"></td>
</tr>
<%
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
</FORM>
</table>
<%
end if
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