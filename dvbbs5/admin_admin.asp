<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="md5.asp"-->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<script language="JavaScript">
<!--
function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.name != 'chkall')
       e.checked = form.chkall.checked;
    }
  }
//-->
</script>
<%
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim body,username2,password2,oldpassword,oldusername,oldadduser,username1
		if request("action")="updat" then
			call update()
			response.write body
		elseif request("action")="del" then
			call del()
			response.write body
        elseif request("action")="pasword" then
			call pasword()
        elseif request("action")="newpass" then
			call newpass()
			response.write body
		elseif request("action")="add" then
			call addadmin()
		elseif request("action")="savenew" then
			call savenew()
			response.write body
		else
			call userlist()
		end if
		conn.close
		set conn=nothing
	end if

	sub userlist()
%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
                <tr> 
                  <th height=22 colspan=5>����Ա����(����û������в���)</th>
                </tr>
                <tr align=center> 
                  <td width="30%" height=22 class="forumHeaderBackgroundAlternate"><B>�û���</B></td><td width="25%" class="forumHeaderBackgroundAlternate"><B>�ϴε�½ʱ��</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>�ϴε�½IP</B></td><td width="15%" class="forumHeaderBackgroundAlternate"><B>����</B></td>
                </tr>
<%
	set rs=conn.execute("select * from admin order by LastLogin desc")
	do while not rs.eof
%>
                <tr> 
                  <td><a href="admin_admin.asp?id=<%=rs("id")%>&action=pasword"><font color=<%=Forum_body(7)%>><%=rs("username")%></font></a></td><td><font color=<%=Forum_body(7)%>><%=rs("LastLogin")%></font></td><td><font color=<%=Forum_body(7)%>><%=rs("LastLoginIP")%></font></td><td><a href="admin_admin.asp?action=del&id=<%=rs("id")%>&name=<%=rs("username")%>">ɾ��</a></td>
                </tr>
<%
	rs.movenext
	loop
	rs.close
	set rs=nothing
%>
	       </table>
<%
	end sub

	sub del()
	conn.execute("delete from admin where id="&request("id"))
	conn.execute("update [user] set usergroupid=4 where username='"&replace(request("name"),"'","")&"'")
	body="<li>����Աɾ���ɹ���"
	end sub

sub pasword()
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	oldadduser=rs("adduser")
  %> 
<form action="?action=newpass" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա���Ϲ����������޸�
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��½���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=hidden name="oldusername" value="<%=rs("username")%>">
              <input type=text name="username2" value="<%=rs("username")%>">  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��½���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2" value="<%=oldpassword%>">  (����ע�����벻ͬ,��Ҫ�޸���ֱ������)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><%=oldadduser%>
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 

              <input type=hidden name="oldadduser" value="<%=oldadduser%>">
              <input type=hidden name="adduser" value="<%=oldadduser%>">
              <input type=hidden name=id value="<%=request("id")%>">
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%       rs.close
         set rs=nothing
end sub

sub newpass()
	dim passnw,usernw,aduser
	set rs=conn.execute("select * from admin where id="&request("id"))
	oldpassword=rs("password")
	if request("username2")="" then
		body="<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		usernw=trim(request("username2"))
	end if
	if request("password2")="" then
		body="<li>�������������롣<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	elseif trim(request("password2"))=oldpassword then
		passnw=request("password2")
	else
		passnw=md5(request("password2"))
	end if
	if request("adduser")="" then
		body="<li>���������Ա���֡�<a href=?>�� <font color=red>����</font> ��</a>"
		exit sub
	else 
		aduser=trim(request("adduser"))
	end if

	set rs=server.createobject("adodb.recordset")
	sql="select * from admin where username='"&trim(request("oldusername"))&"'"
	rs.open sql,conn,1,3
	if not rs.eof and not rs.bof then
	rs("username")=usernw
	rs("adduser")=aduser
	rs("password")=passnw
	body="<li>����Ա���ϸ��³ɹ������ס������Ϣ��<br> ����Ա��"&request("username2")&" <BR> ��   �룺"&request("password2")&" <a href=?>�� <font color=red>����</font> ��</a>"
	rs.update
	
	end if
	rs.close
	set rs=nothing
end sub


sub addadmin()
%> 
<form action="?action=savenew" method=post>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
               <tr> 
                  <th colspan=2 height=23>����Ա��������ӹ���Ա
                  </th>
                </tr>
               <tr > 
            <td width="26%" align="right" class=forumrow>��̨��½���ƣ�</td>
            <td width="74%" class=forumrow>
              <input type=text name="username2">  (����ע������ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow>��̨��½���룺</td>
            <td width="74%" class=forumrow>
              <input type="password" name="password2">  (����ע�����벻ͬ)
            </td>
          </tr>
          <tr > 
            <td width="26%" align="right" class=forumrow height=23>ǰ̨�û����ƣ�</td>
            <td width="74%" class=forumrow><input type=text name="username1">  (��ѡ����д�������޸�)
            </td>
          </tr>
          <tr align="center"> 
            <td colspan="2" class=forumrow> 
              <input type="submit" name="Submit" value="�� ��">
            </td>
          </tr>
        </table>
        </form>

<%
end sub

sub savenew()
	if request.form("username2")="" then
	body="�������̨��½�û�����"
	exit sub
	end if
	if request.form("username1")="" then
	body="������ǰ̨��½�û�����"
	exit sub
	end if
	if request.form("password2")="" then
	body="�������̨��½���룡"
	exit sub
	end if
	set rs=conn.execute("select username from [user] where username='"&replace(request.form("username1"),"'","")&"'")
	if rs.eof and rs.bof then
	body="��������û�������һ����Ч��ע���û���"
	exit sub
	end if
	set rs=conn.execute("select username from admin where username='"&replace(request.form("username2"),"'","")&"'")
	if not (rs.eof and rs.bof) then
	body="��������û����Ѿ��ڹ����û��д��ڣ�"
	exit sub
	end if
	conn.execute("update [user] set usergroupid=1 where username='"&replace(request.form("username1"),"'","")&"'")
	conn.execute("insert into admin (username,[password],adduser) values ('"&replace(request.form("username2"),"'","")&"','"&md5(replace(request.form("password2"),"'",""))&"','"&replace(request.form("username1"),"'","")&"')")
	body="��ӳɹ������ס�¹���Ա��̨��½��Ϣ�������޸��뷵�ع���Ա����"
end sub
%>