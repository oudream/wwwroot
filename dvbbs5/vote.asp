<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/char_board.asp" -->
<%
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if founderr then
		call nav()
		call head_var(1)
		call dvbbs_error()
	else
		stats=boardtype&"����ͶƱ"
		call nav()
		call head_var(2)
		call main()
		if founderr then call dvbbs_error()
	end if
	call footer()
sub main()
	if Cint(GroupSetting(8))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳����ͶƱ��Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
	end if
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
%>
<script src="inc/ubbcode.js"></script>
<form action=Savevote.asp?boardID=<%=request("boardid")%> method=POST onSubmit=submitonce(this) name=frmAnnounce>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
      <th width=100% colspan=2 align=left>&nbsp;&nbsp;������ͶƱ </th>
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
		            <td width=80% class=tablebody1><input name=subject size=50 maxlength=80 class=FormClass>&nbsp;
<select name=votetimeout size=1>
<option value=0>����ʱ��</option>
<option value=0>��������</option>
<option value=1>һ��</option>
<option value=3>����</option>
<option value=7>һ��</option>
<option value=15>����</option>
<option value=30>һ��</option>
<option value=90>����</option>
<option value=180>����</option>
</select>
&nbsp;<font color=<%=Forum_body(8)%>><b>*</b></font>���ó���25�����ֻ�50��Ӣ���ַ�
	 </td>
        </tr>
        <tr>
          <td width=20% class=tablebody1><b>ͶƱ��Ŀ </b> <br>
        <br>
        <li>ÿ��һ��ͶƱ��Ŀ�����<b><%=Forum_Setting(39)%></b>��</li>
        <li>�����Զ����ϣ������Զ�����</li><br>
        <br>
        <input type=radio name=votetype value=0 checked>
          ��ѡͶƱ<br>
          <input type=radio name=votetype value=1>
          ��ѡͶƱ</font></td>
	  <td width=80% class=tablebody1><textarea name=vote cols=95 rows=8></textarea>
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
<!--��ǩ״̬-->
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
<iframe name="ad" frameborder=0 width=100% height=30 scrolling=no src=saveannounce_upload.asp?boardid=<%=boardid%>></iframe>
<%end if%>
<%if Forum_Setting(17)=1 then%>
<!--#include file="inc/getubb.asp"-->
<%end if%>
<br>
<textarea class=smallarea cols=95 name=Content rows=12 wrap=VIRTUAL title=����ʹ��Ctrl+Enterֱ���ύ���� class=FormClass onkeydown=ctlent()></textarea>
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
                <input type=checkbox name=emailflag value=yes>�лظ�ʱʹ���ʼ�֪ͨ����
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
<%
	call activeonline()
end sub
%>