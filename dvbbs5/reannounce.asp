<!--#include file="conn.asp"-->
<!-- #include file="inc/char_board.asp" -->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<%
dim AnnounceID
dim replyID
dim username
dim dateandtime
dim bgcolor,abgcolor
dim PostUserGroup
dim con,content
dim TotalUseTable

if BoardID="" or not isInteger(BoardID) or BoardID=0 then
	Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
else
	BoardID=clng(BoardID)
end if
if request("id")="" then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
elseif not isInteger(request("id")) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
else
	AnnounceID=request("id")
end if
if request("replyid")="" then
	replyid=Announceid
elseif not isInteger(request("replyid")) then
	replyid=Announceid
else
	replyid=request("replyid")
end if
if founderr then
	call nav()
	call head_var(1)
	call dvbbs_error()
else
	stats=boardtype&"�ظ�����"
	call nav()
	call head_var(2)
	call main()
	if founderr then call dvbbs_error()
end if
call footer()

sub main()
if cint(lockboard)=2 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����<a href=login.asp>��½</a>��ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
	else
		if chkboardlogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>����̳Ϊ��֤��̳����ȷ�������û����Ѿ��õ�����Ա����֤����롣"
		founderr=true
		exit sub
		end if
	end if
end if

set rs=conn.execute("select PostTable,locktopic from topic where TopicID="&AnnounceID)
if not (rs.eof and rs.bof) then
	if rs(1)=1 or rs(1)=2 then
		Errmsg=Errmsg+"<br>"+"<li>�������ѱ���������ɾ�������ܷ���ظ���"
		founderr=true
		exit sub
	end if
	TotalUseTable=rs(0)
else
	ErrMsg=ErrMsg+"<br>"+"<li>��ָ�������Ӳ�����</li>"
	founderr=true
	exit sub
end if

if replyid=Announceid then
	set rs=conn.execute("select top 1 Announceid from "&TotalUseTable&" where rootID="&AnnounceID&" order by Announceid")
	if not(rs.bof and rs.eof) then
		replyID=rs(0)
	else
		ErrMsg=ErrMsg+"<br>"+"<li>��ָ�������Ӳ�����</li>"
		founderr=true
		exit sub
	end if
end if

sql="select b.body,b.topic,b.locktopic,b.username,b.dateandtime,b.isbest,u.lockuser,u.UserGroupID from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.AnnounceID="&replyID
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg=Errmsg+"<br>"+"<li>û���ҵ�������ӡ�"
	founderr=true
	exit sub
else
	if rs("locktopic")=1 or rs("locktopic")=2 then
		Errmsg=Errmsg+"<br>"+"<li>�������ѱ���������ɾ�������ܷ���ظ���"
		founderr=true
		exit sub
	elseif rs("lockuser")=2 then
   		con=""
	elseif rs("isbest")=1 and Cint(GroupSetting(41))=0 then
		con=""
	else
   		con=rs("body")
	end if
	username=rs("username")
	dateandtime=rs("dateandtime")
	PostUserGroup=rs("UserGroupID")
	if username=membername then
		if Cint(GroupSetting(4))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�ظ��Լ������Ȩ�ޡ�"
		founderr=true
		exit sub
		end if
	else
		if Cint(GroupSetting(5))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳�ظ����������ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
		end if
		if Cint(GroupSetting(2))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û������ڱ���̳�鿴�����˷��������ӵ�Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
		end if
	end if
end if
set rs=nothing

if request("reply")="true" then
	con = reubbcode(con)
	con = replace(con, "&gt;", ">")
	con = replace(con, "&lt;", "<")
	con = Replace(con, "</P><P>", CHR(10) & CHR(10))
	con = Replace(con, "<BR>", CHR(10))
	content = "[quote][b]����������[i]"&username&"��"&dateandtime&"[/i]�ķ��ԣ�[/b]" & chr(13)
	content = content & con & chr(13)
	content = content & "[/quote]"
else
	content=""
end if
%>
<script src="inc/ubbcode.js"></script>
<form action="SaveReAnnounce.asp?method=Topic&boardID=<%=request("boardid")%>" method=POST onSubmit="submitonce(this)" name=frmAnnounce>
<input type=hidden name="followup" value="<%=replyID%>">
<input type=hidden name="rootID" value="<%=AnnounceID%>">
<input type=hidden name="star" value="<%=request("star")%>">
<input type=hidden name="TotalUseTable" value="<%=TotalUseTable%>">

<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% align=left colspan=2>&nbsp;&nbsp;����ظ�</th>
    </tr>
        <tr>
          <td width=20% class=tablebody1><b>�û���</b></td>
          <td width=80% class=tablebody1><input name=username value=<%=membername%> class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font><a href=reg.asp>��û��ע�᣿</a> 
          </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>����</b></td>
          <td width=80% class=tablebody1><input name=passwd type=password value=<%=htmlencode(memberword)%> class=FormClass><font color=<%=Forum_body(8)%>>&nbsp;&nbsp;<b>*</b></font><a href=lostpass.asp>�������룿</a></td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>�������</b>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">ѡ����</OPTION> <OPTION value=[ԭ��]>[ԭ��]</OPTION> 
              <OPTION value=[ת��]>[ת��]</OPTION> <OPTION value=[��ˮ]>[��ˮ]</OPTION> 
              <OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[�Ƽ�]>[�Ƽ�]</OPTION> <OPTION value=[����]>[����]</OPTION> 
              <OPTION value=[ע��]>[ע��]</OPTION> <OPTION value=[��ͼ]>[��ͼ]</OPTION>
              <OPTION value=[����]>[����]</OPTION> <OPTION value=[����]>[����]</OPTION>
              <OPTION value=[����]>[����]</OPTION></SELECT>
	  </td>
		            <td width=80% class=tablebody1><input name=subject size=70 maxlength=80 class=FormClass>&nbsp;&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>���ó��� 25 �����ֻ�50��Ӣ���ַ�
	 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1><b>��ǰ����</b><br><li>���������ӵ�ǰ��</td>
          <td width=80% class=tablebody1>
<%for i=1 to 9%>
	<input type="radio" value="face<%=i%>" name="Expression" <%if i=1 then response.write "checked"%>><img src="face/face<%=i%>.gif" WIDTH="15" HEIGHT="15">&nbsp;&nbsp;
<%next%>
<br>
<%for i=10 to 18%>
	<input type="radio" value="face<%=i%>" name="Expression"><img src="face/face<%=i%>.gif" WIDTH="15" HEIGHT="15">&nbsp;&nbsp;
<%next%>
 </td>
        </tr>
        <tr>
          <td width=20% valign=top class=tablebody1>
<b>����</b><br>
<li>HTML��ǩ�� <%=iif(Forum_Setting(35),"����","������")%>
<li>UBB��ǩ�� <%=iif(Forum_Setting(34),"����","������")%>
<li>��ͼ��ǩ�� <%=iif(Forum_Setting(37),"����","������")%>
<li>Flash��ǩ��<%=iif(Forum_Setting(38),"����","������")%>
<li>�����ַ�ת����<%=iif(Forum_Setting(36),"����","������")%>
<li>�ϴ�ͼƬ��<%=iif(Forum_Setting(3),"����","������")%>
<li>���<%=Forum_Setting(13)\1024%>KB <BR><BR>
<B>��������</B><BR>
<li><a href="javascript:money()" title="ʹ���﷨��[money=������ò���������Ҫ��Ǯ��]����[/money]">��Ǯ��</a>
<li><a href="javascript:point()" title="ʹ���﷨��[point=������ò���������Ҫ��Ǯ��]����[/point]">������</a>
<li><a href="javascript:usercp()" title="ʹ���﷨��[usercp=������ò���������Ҫ��Ǯ��]����[/usercp]">������</a>
<li><a href="javascript:power()" title="ʹ���﷨��[power=������ò���������Ҫ��Ǯ��]����[/power]">������</a>
<li><a href="javascript:article()" title="ʹ���﷨��[post=������ò���������Ҫ��Ǯ��]����[/post]">������</a>
<li><a href="javascript:replyview()" title="ʹ���﷨��[replyview]�ò������ݻظ���ɼ�[/replyview]">�ظ��ɼ���</a>
	  </td>
          <td width=80% class=tablebody1>
<%if Cint(GroupSetting(7))=1 then%>
<iframe name="ad" frameborder=0 width=100% height=30 scrolling=no src="saveannounce_upload.asp?boardid=<%=boardid%>"></iframe>
<%end if%>
<%if Forum_Setting(17)=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<br>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=����ʹ��Ctrl+Enterֱ���ύ���� class=FormClass onkeydown=ctlent()><%=Server.htmlencode(content)%></textarea>
          </td>
        </tr>

		<tr>
                <td class=tablebody1 valign=top colspan=2><b>�������ͼ�����������м�����Ӧ�ı���</B><br>&nbsp;
<%
for i=1 to 28
	if len(i)=1 then i="0" & i
	response.write "<img src=""pic/em"&i&".gif"" border=0 onclick=""insertsmilie('[em"&i&"]')"" style=""CURSOR: hand"">&nbsp;"
next
%>
    		</td>
                </tr>
                <tr>
                <td valign=top class=tablebody1><b>ѡ��</b></td>
                <td valign=middle class=tablebody1><input type=checkbox name=signflag value=yes checked>�Ƿ���ʾ����ǩ����<br>
                <input type=checkbox name=emailflag value=yes>�лظ�ʱʹ���ʼ�֪ͨ����</font>
<BR><BR></td>
                </tr><tr>
                <td valign=middle colspan=2 align=center class=tablebody2>
                <input type=Submit value='�� ��' name=Submit> &nbsp; <input type=button value='Ԥ ��' name=Button onclick=gopreview()>&nbsp;
<input type=reset name=Submit2 value='�� ��'>
                </td></form></tr>
      </table>
</form>
<form name=preview action=preview.asp method=post target=preview_page>
<input type=hidden name=title value=><input type=hidden name=body value=>
</form>
<hr width=<%=Forum_body(12)%> size=1>
<%
	sql="Select top 10 b.UserName,b.Topic,b.dateandtime,b.body,b.announceid,b.isbest,u.lockuser,u.UserGroupID from "&TotalUseTable&" b inner join [user] u on b.postuserid=u.userid where b.boardid="&boardid&" and b.rootid="&Announceid&" and not b.locktopic=2 order by b.announceid desc"
	set rs=conn.execute(sql)
	do while not rs.eof
	PostUserGroup=rs("UserGroupID")
	if bgcolor="tablebody1" then
		bgcolor="tablebody2"
		abgcolor="tablebody1"
	else
		bgcolor="tablebody1"
		abgcolor="tablebody2"
	end if
%>
<TABLE border=0 width=<%=Forum_body(12)%> align=center>
  <TBODY>
  <TR>
    <TD valign=middle align=top class=<%=bgcolor%>>
--&nbsp;&nbsp;���ߣ�<%=htmlencode(rs("username"))%><br>
--&nbsp;&nbsp;����ʱ�䣺<%=rs("dateandtime")%><br><br>
--&nbsp;&nbsp;<%=htmlencode(rs("topic"))%><br>
<%
	if rs("lockuser")=2 then
		response.write "���û������ѱ�����"
	elseif rs("isbest")=1 and Cint(GroupSetting(41))=0 then
		response.write "��û������þ������ӵ�Ȩ��"
	else
		response.write dvbcode(rs("body"),PostUserGroup)
	end if
%>
    </TD></TR></TBODY></TABLE> 
	<hr size=1 class=tableborder1>
<%
	rs.movenext
	loop	
	set rs=nothing
	call activeonline()
end sub
%>