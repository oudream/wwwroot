<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
'=========================================================
' File: list.asp
' Version:5.0
' Date: 2002-9-20
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

Dim currentPage,tl,limitime
Dim Fast_action
Fast_action=false

if boardmaster or master then
	guestlist=""
	Fast_action=true
else
	guestlist=" lockboard<>1 and "
	Fast_action=false
end if

BoardID = Request("BoardID")
currentPage=request("page")
if BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("selTimeLimit")="all" then
	tl=""
elseif request("selTimeLimit")="" then
	tl=""
else
	limitime=request("selTimeLimit")
	tl=" and dateandtime>='"&cstr(cdate(now()-limitime))&"' "
end if
if cint(lockboard)=2 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��½</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		end if
	end if
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	call nav()
	call head_var(1)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
Dim LastPost,Lastuser,LastID
dim LastTime,LastUserid,LastRootid,body
dim Pnum,p,replynum
dim totalrec,ii,page_count
dim n,pi
dim call_info
dim onlinenum,guestnum
dim rs1,sql1
onlinenum=online(boardid)
guestnum=guest(boardid)
if trim(boardmasterlist)<>"" then
	master_1=split(boardmasterlist, "|")
	for i = 0 to ubound(master_1)
	master_2=""+master_2+"<a href=""dispuser.asp?name="+master_1(i)+""" target=_blank title=����鿴�ð�������>"+master_1(i)+"</a>&nbsp;"
	next
else
	master_2="�ް���"
end if
call activeonline()
%>
<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=100% valign=middle>
<SCRIPT LANGUAGE='JavaScript' SRC='js/fader.js' TYPE='text/javascript'></script>
<SCRIPT LANGUAGE='JavaScript' TYPE='text/javascript'>
prefix="";
arNews = ["<%
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='��ǰû�й���'>��ǰû�й���</ACRONYM></a></b> ("& now() &")"
else
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
%>","","��ӭ����<%=Forum_info(0)%> <%=boardtype%>","",
"�Ͻ�����ʹ�ô��Ի��Υ�߾�Ȱ����Ч��������ID��",""
]
</SCRIPT>
<div id="elFader" style="position:relative;visibility:hidden; height:16" ></div>
</td>
</tr>
</table><BR>
<script>
function loadThreadFollow(t_id,b_id){
	var targetImg =eval("document.all.followImg" + t_id);
	var targetDiv =eval("document.all.follow" + t_id);
	if (targetImg.src.indexOf("nofollow")!=-1){return false;}
	if ("object"==typeof(targetImg)){
		if (targetDiv.style.display!='block'){
			targetDiv.style.display="block";
			targetImg.src="pic/minus.gif";
			if (targetImg.loaded=="no"){
				document.frames["hiddenframe"].location.replace("loadtree1.asp?boardid="+b_id+"&rootid="+t_id);
			}
		}else{
			targetDiv.style.display="none";
			targetImg.src="pic/plus.gif";
		}
	}
}
</script>

<iframe width=0 height=0 src="" id="hiddenframe"></iframe>
<!--<TABLE cellSpacing=0 cellPadding=0 width=<%=Forum_body(12)%> border=0 align=center>
<tr>
<td align=center width=34 valign=middle> <img src="<%=Forum_info(7)&Forum_boardpic(16)%>" border=0  width=20 height=17>
</td><td valign=middle align=left>
<%
sql="select top 1 title,addtime from bbsnews where boardid="&boardid&"  order by id desc"
set rs=conn.execute(sql)
if rs.bof and rs.eof then
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank><ACRONYM TITLE='��ǰû�й���'>��ǰû�й���</ACRONYM></a></b> ("& now() &")"
else
	response.write "<b><a href='announcements.asp?boardid="& boardid &"' target=_blank>"& rs("title") &"</a></b> ("& rs("addtime") &")"
end if
set rs=nothing
%>
</td>
<td align=right valign=middle><p>
<form action="list.asp" method=get>
<input type=hidden name=boardid value="<%= boardid %>">
<select name=selTimeLimit onchange='javascript:submit()'>
<option value=all>�鿴���е�����
<option value=1>�鿴һ���ڵ�����
<option value=2>�鿴�����ڵ�����
<option value=7>�鿴һ�����ڵ�����
<option value=15>�鿴������ڵ�����
<option value=30>�鿴һ�����ڵ�����
<option value=60>�鿴�������ڵ�����
<option value=180>�鿴�����ڵ�����
</select>
</form></p>
</td></tr>
</TABLE>
--><BR>
<TABLE cellpadding=3 cellspacing=1 class=tableborder1 align=center>
<TR>
<Th height=25 width="100%" align=left id=tabletitlelink style="font-weight:normal">������<b><%=allonline()%></b>�ˣ�����<%= boardtype %>�Ϲ��� <b><%= onlinenum %></b> λ��Ա�� <b><%= guestnum %></b> λ����.�������� <b><font color=<%=Forum_body(8)%>><%= todaynum %></font></b> 
<%
if request("action")="show" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
else
	if cint(Forum_Setting(14))=1 and request("action")<>"off" then
	response.write "[<a href=list.asp?action=off&boardID="& boardid&"&amp;page=1&skin="& skin &">�ر���ϸ�б�</a>]"
	else
	response.write "[<a href=list.asp?action=show&boardID="& boardid&"&amp;page=1&skin="& skin &">��ʾ��ϸ�б�</a>]"
	end if
end if
%>
</Th></TR>
<%
if request("action")="off" then
	call onlineuser(0,0,boardid)
elseif request("action")="show" then
	call onlineuser(1,1,boardid)
else
	call onlineuser(Forum_Setting(14),Forum_Setting(15),boardid)
end if
%>
</td></tr>
</TABLE>
<BR>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center valign=middle><tr>
<td align=center width=2> </td>
<td align=left> <a href="announce.asp?boardid=<%= boardid %>"><img src="<%=Forum_info(7)&Forum_boardpic(1)%>" border=0 alt="������"></a>
&nbsp;&nbsp;<a href="vote.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(2)%>" border=0 alt="������ͶƱ"></a>
&nbsp;&nbsp;<a href="smallpaper.asp?boardid=<%=boardid%>"><img src="<%=Forum_info(7)&Forum_boardpic(3)%>" border=0 alt="����С�ֱ�"></a></td>
<td align=right><img src="<%=Forum_info(7)%>team2.gif" align=absmiddle>  <%= master_2 %></td></tr></table>

<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>
<tr><td class=tablebody1 colspan=5 height=20>
<table width=100% >
<tr>
<td valign=middle height=20 width=50> <a href=AllPaper.asp?boardid=<%=boardid%> title=����鿴����̳����С�ֱ�><b>�㲥��</b></a> </td><td width=*> 
<marquee scrolldelay=150 scrollamount=4 onmouseout="if (document.all!=null){this.start()}" onmouseover="if (document.all!=null){this.stop()}">
<%
set rs=conn.execute("SELECT TOP 5 s_id,s_username,s_title FROM smallpaper where datediff('d',s_addtime,Now())<=1 and s_boardid="&boardid&" ORDER BY s_addtime desc")
do while not rs.eof
	response.write "����<font color="&Forum_pic(5)&">"&htmlencode(rs(1))&"</font>˵��<a href=javascript:openScript('viewpaper.asp?id="&rs(0)&"&boardid="&boardid&"',500,400)>"&htmlencode(rs(2))&"</a>"
rs.movenext
loop
set rs=nothing
%>
</marquee><td align=right width=200><a href='elist.asp?boardid=<%= boardid %>' title=�鿴���澫��><font color=<%=Forum_body(8)%>><B>����</B></font></a> | <a href=boardstat.asp?reaction=online&boardid=<%=boardid%> title=�鿴����������ϸ���>����</a> | <a href=bbseven.asp?boardid=<%=boardid%> title=�鿴�����¼�>�¼�</a> | <a href=Boardhelp.asp?boardid=<%=boardid%> title=�鿴�������>����</a> | <a href='?boardid=<%=boardid%>&action=batch'>����</a></td></tr>
</table>
</td></tr>
<form action="admin_batch.asp" method=post name=batch>
<TR align=middle>
<Th height=25 width=32 id=tabletitlelink><a href="list.asp?boardid=<%=boardid%>&page=<%=request("page")%>&action=batch">״̬</a></th>
<Th width=*>�� ��  (��<img src=<%=Forum_info(7)%>plus.gif>����չ�������б�)</Th>
<Th width=80>�� ��</Th>
<Th width=64>�ظ�/����</Th>
<Th width=195>������ | �ظ���</Th>
</TR>
<%
if limitime="" then
	totalrec=lasttopicnum
else
	set rs=conn.execute("select count(topicid) from topic where boardID="&cstr(boardID)&" "&tl&"")
	totalrec=rs(0)
end if
set rs=server.createobject("adodb.recordset")
sql="select TopicID,boardid,title,postusername,postuserid,dateandtime,child,hits,votetotal,lastpost,lastposttime,istop,isvote,isbest,locktopic,Expression from topic where boardid="&boardid&" and not locktopic=2 order by istop desc,lastposttime desc"

rs.open sql,conn,1
if rs.bof and rs.eof then
	response.write "<tr><td colSpan=5 width=100% class=tablebody1>����̳�������ݣ���ӭ��������</td></tr>"
else
	if totalrec mod Forum_Setting(11)=0 then
		n= totalrec \ Forum_Setting(11)
	else
	     	n= totalrec \ Forum_Setting(11)+1
  	end if
	RS.MoveFirst
	if currentpage > n then currentpage = n
      	if currentpage<1 then currentpage=1
	RS.Move (currentpage-1) * Forum_Setting(11)
	do while not rs.eof and page_count<Clng(Forum_Setting(11))
	page_count=page_count+1
'========================TopicInfo=========================
response.write "<TR align=middle><TD class=tablebody2 width=32 height=27>"

if request("action")="batch" and Cint(GroupSetting(45))=1 then
	response.write "<input type=checkbox name=Announceid value="&rs("topicid")&">"
else
	if rs("istop")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(3)&""" alt=�̶�����>"
	elseif rs("isbest")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(4)&""" alt=��������>"
	elseif rs("isvote")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(12)&""" alt=ͶƱ����>"
	elseif rs("locktopic")=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=������������>"
	elseif lockboard=1 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(2)&""" alt=����̳������>"
	elseif rs("child")>=10 then
	response.write "<img src="""&Forum_info(7)&Forum_statePic(1)&""" alt=��������>"
	else
	response.write "<img src="""&Forum_info(7)&Forum_statePic(0)&""" alt=��������>"
	end if
end if

response.write "</TD><TD align=left class=tablebody1 width=*>"

if rs(6)=0 then
	response.write "<img src="""& Forum_info(7) &"nofollow.gif"" id=""followImg"& rs("topicid") &""">"
else
	response.write "<img loaded=""no"" src="""& Forum_info(7) &"plus.gif"" id=""followImg"& rs("topicid") &""" style=""cursor:hand;"" onclick=""loadThreadFollow("& rs("topicid") &","& boardid &")"" title=չ�������б�>"
end if

if not isnull(rs("lastpost")) then
	LastPost=split(rs("lastpost"),"$")
	if ubound(LastPost)=6 then
	Lastuser=htmlencode(LastPost(0))
	LastID=LastPost(1)
	LastTime=LastPost(2)
	if not isdate(LastTime) then LastTime=rs("dateandtime")
	body=htmlencode(LastPost(3))
	if not isnull(LastPost(4)) and LastPost(4)<>"" then response.write "<img src="&Forum_info(7)&LastPost(4)&".gif width=16 height=16>"
	LastUserid=LastPost(5)
	LastRootid=LastPost(6)
	else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
	end if
else
	Lastuser="------"
	LastID=rs("topicid")
	LastTime=rs("dateandtime")
	body="..."
	LastUserid=rs("postuserid")
	LastRootid=rs("topicid")
end if
if LastTime=rs("dateandtime") then LastUser=""

response.write "<a href=""dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &""" title=""��"& htmlencode(rs("title")) &"��<br>���ߣ�"& htmlencode(rs("postusername")) &"<br>������"& rs("dateandtime") &"<br>��������"& body &"..."">"

if len(rs("title"))>26 then
	response.write left(htmlencode(rs("title")),26)&"..."
else
	response.write htmlencode(rs("title"))
end if

response.write "</a>"
replynum=rs("child")+1
if replynum>Cint(Forum_Setting(12)) then
	response.write "&nbsp;&nbsp;[<img src="&Forum_info(7)&"multipage.gif><b>"
  	if replynum mod Cint(Forum_Setting(12))=0 then
     		Pnum= replynum \ Cint(Forum_Setting(12))
  	else
     		Pnum= replynum \ Cint(Forum_Setting(12))+1
  	end if
	for p=1 to Pnum
	response.write " <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&P&"'><font color="&Forum_body(8)&">"&p&"</font></a> "
	if p+1>7 then
	response.write "... <a href='dispbbs.asp?boardID="& boardID &"&ID="& rs("topicid") &"&star="&Pnum&"'><font color="&Forum_body(8)&">"&Pnum&"</font></a>"
	exit for
	end if
	next
	response.write "</b>]"
else
	pnum=1
end if

response.write "</TD>"
response.write "<TD class=tablebody2 width=80><a href=""dispuser.asp?id="& rs("postuserid") &""" target=_blank>"& htmlencode(rs("postusername")) &"</a></TD>"
response.write "<TD class=tablebody1 width=64>"

if rs("isvote")=1 then
	response.write "<FONT color="&Forum_body(8)&"><b>"&rs("votetotal")&"</b></font>  Ʊ"
else
	response.write rs("child") &"/"& rs("hits")
end if

response.write "</TD>"
response.write "<TD align=left class=tablebody2 width=195>&nbsp;&nbsp;<a href=""dispbbs.asp?boardid="& boardid &"&id="&LastRootID&"&star="&pnum&"#"& LastID&""">"
response.write FormatDateTime(LastTime,2)&"&nbsp;"&FormatDateTime(LastTime,4)&"</a>&nbsp;<font color="&Forum_body(8)&">|</font>&nbsp;"

if LastUser="" then
	response.write "------"
else
	response.write "<a href=dispuser.asp?id="&LastUserID&" target=_blank>"&LastUser&"</a>"
end if

response.write "</TD></TR>"
response.write "<tr style=""display:none"" id=""follow"& rs("topicid") &"""><td colspan=5 id=""followTd"& rs("topicid") &""" style=""padding:0px""><div style=""width:240px;margin-left:18px;border:1px solid black;background-color:lightyellow;color:black;padding:2px"" onclick=""loadThreadFollow("& rs("topicid") &","&boardid&")"">���ڶ�ȡ���ڱ�����ĸ��������Ժ��</div></td></tr>"
'===========================End============================
	Lastuser=""
	LastID=""
	LastTime=""
	body=""
	rs.movenext
	loop

	if request("action")="batch" then
%>
<tr><td height=30 width=100% class=tablebody1 colspan=5>��<select name=newboard size=1>
<option selected>�ƶ�������ѡ��</option>
<%
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option>�� "& rs(1) &"</option>"

		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			response.write "<option>û����̳</option>"
		else
			do while not rs1.eof
				response.write "<option value="""&rs1(0)&""">����" & rs1(1) & "</option>"
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
%>
<input type=hidden value="<%=boardid%>" name=boardid><input type=checkbox name=chkall value=on onclick='CheckAll(this.form)'>ȫѡ/ȡ�� <input type=radio name=action value=dele>����ɾ�� <input type=radio name=action value=move>�����ƶ� <input type=radio name=action value=isbest>�������� <input type=radio name=action value=lock>�������� <input type=radio name=action value=istop>�����̶� <input type=submit name=Submit value=ִ��  onclick="{if(confirm('��ȷ��ִ�еĲ�����?')){this.document.batch.submit();return true;}return false;}"></font></td></tr></form>
<%
	end if
	set rs1=nothing
end if
set rs=nothing
if currentpage-1 mod 10=0 then
	p=(currentpage-1) \ 10
else
	p=(currentpage-1) \ 10
end if
dim pagelist,pagelistbit
%>
</table>

<table border=0 cellpadding=0 cellspacing=3 width=<%=Forum_body(12)%> align=center >
</form>
<form method=post action=list.asp>
<input type=hidden name=selTimeLimit value='<%= request("selTimeLimit") %>'><tr>
<td valign=middle>ҳ�Σ�<b><%= currentPage %></b>/<b><%= n %></b>ҳ ÿҳ<b><%= Forum_Setting(11) %></b> ������<b><%= totalrec %></b></td>
<td valign=middle><div align=right >��ҳ��
<%
	if currentPage=1 then
	response.write "<font face=webdings color="&Forum_body(8)&">9</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page=1&selTimeLimit="&request("selTimeLimit")&"' title=��ҳ><font face=webdings>9</font></a>   "
	end if
	if p*10>0 then response.write "<a href='?boardid="&boardid&"&page="&Cstr(p*10)&"&selTimeLimit="&request("selTimeLimit")&"' title=��ʮҳ><font face=webdings>7</font></a>   "
	response.write "<b>"
	for ii=p*10+1 to P*10+10
		   if ii=currentPage then
	          response.write "<font color="&Forum_body(8)&">"+Cstr(ii)+"</font> "
		   else
		      response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&selTimeLimit="&request("selTimeLimit")&"'>"+Cstr(ii)+"</a>   "
		   end if
		if ii=n then exit for
		'p=p+1
	next
	response.write "</b>"
	if ii<n then response.write "<a href='?boardid="&boardid&"&page="&Cstr(ii)&"&selTimeLimit="&request("selTimeLimit")&"' title=��ʮҳ><font face=webdings>8</font></a>   "
	if currentPage=n then
	response.write "<font face=webdings color="&Forum_body(8)&">:</font>   "
	else
	response.write "<a href='?boardid="&boardid&"&page="&Cstr(n)&"&selTimeLimit="&request("selTimeLimit")&"' title=βҳ><font face=webdings>:</font></a>   "
	end if
%>
ת��:<input type=text name=Page size=3 maxlength=10  value='<%= currentpage %>'><input type=submit value=Go name=submit>
</div></td></tr>
<input type=hidden name=BoardID value='<%= BoardID %>'>
</form></table>

<table border=0 cellpadding=0 cellspacing=3 width="<%=Forum_body(12)%>" align=center>
<tr>
<FORM METHOD=POST ACTION="queryresult.asp?boardid=<%=boardid%>&sType=2&SearchDate=30&pSearch=1">
<td width=50% valign=middle nowrap height=40>
����������<input type=text name="keyword">&nbsp;<input type=submit name=submit value="����">
</td>
</FORM>
<td valign=middle nowrap width=50%> <div align=right>
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}">
<option selected>��ת��̳��...</option>
<%
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option>�� "& rs(1) &"</option>"

		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			response.write "<option>û����̳</option>"
		else
			do while not rs1.eof
				response.write "<option value=""list.asp?boardid="&rs1(0)&""">����" & rs1(1) & "</option>"
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
	set rs=nothing
	set rs1=nothing
%>
</select><div></td></tr></table>
<BR>
<table cellspacing=1 cellpadding=3 width=100% class=tableborder1 align=center><tr>
<th width=80% align=left>��-=> <%= Forum_info(0) %>ͼ��</th>
<th noWrap width=20% align=right>����ʱ���Ϊ - <%=Forum_info(9)%> &nbsp;</th>
</tr>
<tr><td colspan=2 class=tablebody1>
<table cellspacing=4 cellpadding=0 width=92% border=0 align=center>
<tr><td><img src=<%=Forum_info(7)&Forum_statePic(0)%>> ���ŵ�����</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(1)%>> �ظ�����10��</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(2)%>> ����������</td>
<td><img src=<%=Forum_info(7)&Forum_statePic(3)%>> �̶����˵����� </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(4)%>> �������� </td>
<td> <img src=<%=Forum_info(7)&Forum_statePic(12)%>> ͶƱ���� </td>
</tr>
</table>
</td></tr></table>
<%
end sub

if instr(scriptname,"index.asp")>0 or instr(scriptname,"list.asp")>0 then
	if Forum_ads(2)=1 then call admove()
	if Forum_ads(13)=1 then call fixup()
end if
%>
<!--#include file="online_l.asp"-->
<!--#include file="inc/ad_fixup.asp"-->