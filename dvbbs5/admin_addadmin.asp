<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="admin_config.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<%=Eval(""""&gettemplate("Forum_css",4)&"""")%>

<BODY  <%=Forum_body(11)%>>
<%
	if not master or instr(session("flag"),"41")=0 then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		dim body
		call main()
		conn.close
		set conn=nothing
	end if

	sub main()
%>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> bgcolor=<%=Forum_body(0)%> align=center>
  <tr>
    <td>
      <table cellpadding=3 cellspacing=1 border=0 width=100%>
        <tr bgcolor='<%=Forum_body(2)%>'>
        <td><font color=<%=Forum_body(6)%>>��ӭ<b><%=membername%></b>�����������</font>
        </td>
        </tr>
            <tr bgcolor=<%=Forum_body(4)%>>
              <td width="100%" valign=top><p><font color=<%=Forum_body(7)%>>
<%
	if request("action")="updat" then
		call update()
		response.write body
	else
%>
              <table width="95%" border="0" cellspacing="3" cellpadding="0" align=center>
                <tr> 
                  <td bgcolor="<%=Forum_body(3)%>"> <font color=<%=Forum_body(7)%>>
                    <p><b>��ӹ���Ա</b>��<br>
                      ע�⣺��ӹ���Ա���뵽����ԱȨ������ҳ��������Ȩ�޽������á�</p></font>
                  </td>
                </tr>
                <tr> 
                  <td bgcolor="<%=Forum_body(2)%>" height=22><font color=<%=Forum_body(6)%>><b>>>��ӹ���Ա</b>����������д������Ϣ��</font>
                  </td>
                </tr>
<form action="admin_addadmin.asp?action=updat" method=post>
		<tr><td height=25><font color=<%=Forum_body(7)%>>
��ȷ�����û���������̳ע���û�<br>
�û�����<input name=username type=text size=30>��<input type="submit" name="Submit" value="���">
		</font></td>
		</tr>
</form>
	      </table>
<%
	end if
	if founderr then call dvbbs_error()
%></font>
		</td>
            </tr>
        </table>
        </td>
    </tr>
</table>
<%
	end sub

	sub update()
	dim username
	if request("username")="" then
	Founderr=true
	Errmsg=Errmsg+"<br>"+"<li>����д�û����ơ�"
	exit sub
	else
	username=replace(request("username"),"'","")
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select username,userclass from [user] where username='"&username&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
	Founderr=true
	Errmsg=Errmsg+"<br>"+"<li>��̳ע���û���û������ӵ��û�������ȷ�Ϻ���ӡ�"
	exit sub
	else
	rs("userclass")=20
	rs.update
	end if
	rs.close
	sql="select * from admin where username='"&username&"'"
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
	rs.addnew
	rs("username")=username
        rs("adduser")=username
	body="<li>����Ա��ӳɹ����뵽����ԱȨ������ҳ��������Ȩ�޽������ü���<font color=red>�Ĺ���Ա��������</font>��"
	rs.update
	else
	Founderr=true
	Errmsg=Errmsg+"<br>"+"<li>�ù���Ա�Ѿ����ڣ���������ӡ�"
	end if
	rs.close
	set rs=nothing
	end sub
%>