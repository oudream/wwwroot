<%
sub cp_var()
%>
<BR>
<table cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<tr>
<%if not founduser then%>
<td height=25>
>> ��ӭ�� , ������ <a href="login.asp">��½</a> �� <a href="reg.asp">ע��</a>
<%else%>
<td width=65% >
>> ��ӭ <B><%=membername%></B>  [ <a href="JavaScript:openScript('messanger.asp?action=new',500,400)">������</a> :: <a href="topicwithme.asp?s=2">���������</a> :: <a href="topicwithme.asp?s=1">���������</a> :: <a href="dispuser.asp?id=<%=userid%>&boardid=<%=boardid%>&action=permission">������ʲô</a> ]
</td><td width=35% align=right>
<%if Cint(newincept())>Cint(0) then%>
<bgsound src="pic/mail.wav" border=0>
<img src=<%=Forum_info(7)%>msg_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> ��</font></a>)
<%else%>
<img src=<%=Forum_info(7)%>msg_no_new_bar.gif> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<font color=gray>0 ��</font>)
<%end if%>
<%end if%>
</td></tr>
</table>

<table cellspacing=1 cellpadding=3 align=center class=tableBorder2>
<tr><td height=25 valign=middle>
<img src="<%=Forum_info(7)%>Forum_nav.gif" align=absmiddle> <a href=index.asp><%=Forum_info(0)%></a> �� 
<a href=usermanager.asp>�û�<%=Membername%>�Ŀ������</a> �� <%=stats%>
<a name=top></a>
</td></td>
</table>
<br>

<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th width=14% height=25 id=tabletitlelink><a href=usermanager.asp>���������ҳ</a></th>
<th width=14%  id=tabletitlelink><a href=mymodify.asp>���������޸�</a></th>
<th width=14%  id=tabletitlelink><a href=modifypsw.asp>�û������޸�</a></th>
<th width=14%  id=tabletitlelink><a href=modifyadd.asp>��ϵ�����޸�</a></th>
<th width=14%  id=tabletitlelink><a href=usersms.asp>�û����ŷ���</a></th>
<th width=14%  id=tabletitlelink><a href=friendlist.asp>�༭�����б�</a></th>
<th width=14%  id=tabletitlelink><a href=favlist.asp>�û��ղع���</a></th>
</tr>
</table>
<hr size=1 width="<%=Forum_body(12)%>">
<%
end sub
%>