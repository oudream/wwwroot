<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: friendlist.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================

dim msg
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
stats="�����б�"
call nav()
if founderr=true then
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	select case request("action")
	case "info"
		call info()
	case "addF"
		call addF()
	case "saveF"
		call saveF()
	case "ɾ��"
		call DelFriend()
	case "��պ���"
		call AllDelFriend()
	case else
		call info()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub info()
%>
<form action="friendlist.asp" method=post name=inbox>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
            <tr>
                <th valign=middle width="25%">����</td>
                <th valign=middle width="25%">�ʼ�</td>
                <th valign=middle width="25%">��ҳ</td>
                <th valign=middle width="10%">OICQ</td>
                <th valign=middle width="10%">������</td>
                <th valign=middle width="5%">����</td>
            </tr>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select F.*,U.useremail,U.homepage,U.oicq from Friend F inner join [user] U on F.F_Friend=U.username where F.F_username='"&trim(membername)&"' order by F.f_addtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
                <tr>
                <td class=tablebody1 align=center valign=middle colspan=6>���ĺ����б���û���κ����ݡ�</td>
                </tr>
<%else%>
<%do while not rs.eof%>
                <tr>
                    <td align=center valign=middle class=tablebody1><a href="dispuser.asp?name=<%=htmlencode(rs("F_friend"))%>" target=_blank><%=htmlencode(rs("F_friend"))%></a></td>
                    <td align=center valign=middle class=tablebody1><a href="mailto:<%=htmlencode(rs("useremail"))%>"><%=htmlencode(rs("useremail"))%></a></td>
                    <td align=center class=tablebody1><a href="<%=htmlencode(rs("homepage"))%>" target=_blank><%=htmlencode(rs("homepage"))%></a></td>
                    <td align=center class=tablebody1><%=htmlencode(rs("oicq"))%></td>
                    <td align=center class=tablebody1><a href="usersms.asp?action=new&touser=<%=htmlencode(rs("f_friend"))%>">����</a></td>
                <td align=center class=tablebody1><input type=checkbox name=id value=<%=rs("f_id")%>></td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
                
        <tr> 
          <td align=right valign=middle colspan=6 class=tablebody2><input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ѡ��������ʾ��¼&nbsp;<input type=button name=action onclick="location.href='friendlist.asp?action=addF'" value="��Ӻ���">&nbsp;<input type=submit name=action onclick="{if(confirm('ȷ��ɾ��ѡ���ļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="ɾ��">&nbsp;<input type=submit name=action onclick="{if(confirm('ȷ��������еļ�¼��?')){this.document.inbox.submit();return true;}return false;}" value="��պ���"></td>
                </tr>
                </table>
</form>
<%
end sub

sub delFriend()
dim delid
delid=replace(request.form("id"),"'","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
founderr=true
exit sub
else
	conn.execute("delete from Friend where F_username='"&trim(membername)&"' and F_id in ("&delid&")")
	sucmsg=sucmsg+"<br>"+"<li><b>���Ѿ�ɾ��ѡ���ĺ��Ѽ�¼��"
	call dvbbs_suc()
end if
end sub
sub AllDelFriend()
	conn.execute("delete from Friend where F_username='"&trim(membername)&"'")
	sucmsg=sucmsg+"<br>"+"<li><b>���Ѿ�ɾ�������к����б�"
	call dvbbs_suc()
end sub

sub addF()
%>
<form action="Friendlist.asp" method=post name=messager>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=2> 
              <input type=hidden name="action" value="saveF">
              �������--����������������Ϣ</th>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle width=70><b>���ѣ�</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="touser" size=50 value="<%=request("myFriend")%>">
			  &nbsp;ʹ�ö��ţ�,���ֿ������5λ�û�
            </td>
          </tr>
          <tr> 
            <td valign=middle colspan=2 align=center class=tablebody2> 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="���">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub saveF()
dim incept
if request("touser")="" then
	errmsg=errmsg+"<br>"+"<li>��������д���Ͷ����˰ɡ�"
	founderr=true
	exit sub
else
	incept=checkStr(request("touser"))
	incept=split(incept,",")
end if

for i=0 to ubound(incept)
set rs=server.createobject("adodb.recordset")
sql="select username from [user] where username='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>��̳û������û�������δ�ɹ���"
	founderr=true
	exit sub
end if
set rs=nothing

sql="select F_friend from friend where F_username='"&membername&"' and  F_friend='"&incept(i)&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	sql="insert into friend (F_username,F_friend,F_addtime) values ('"&membername&"','"&incept(i)&"',Now())"
	conn.execute(sql)
end if
if i>4 then
	errmsg=errmsg+"<br>"+"<li>ÿ�����ֻ�����5λ�û�����������5λ�Ժ����������д��"
	founderr=true
	exit sub
	exit for
end if
next

sucmsg=sucmsg+"<br>"+"<li><b>��ϲ����������ӳɹ���"
call dvbbs_suc()
end sub
%>