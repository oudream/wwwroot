<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
	stats="��̳����"
	if boardmaster or master then
		guestlist=""
	else
		guestlist=" lockboard<>1 and "
	end if
	call nav()
	if BoardID="" or not isInteger(BoardID) then
		boardid=0
		call head_var(1)
	else
		boardid=cLng(Boardid)
		call head_var(2)
	end if
	call main()
	if founderr then call dvbbs_error()
	call footer()

sub main()
	if Cint(GroupSetting(14))=0 then
		Errmsg=Errmsg+"<br>"+"<li>��û���ڱ���̳������Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
		founderr=true
		exit sub
	end if
%>
<form action=queryresult.asp method=post>
    <table cellpadding=5 cellspacing=1 align=center class=tableborder1>
    	<tr>
	<th valign=middle colspan=2 >������Ҫ�����Ĺؼ���</th></tr>
	<tr>
	<td class=tablebody1 colspan=2 align=center valign=middle>(ͬʱ��ѯ�������ʹ��'<font  color=<%=Forum_body(8)%>>or</font>' �ָ��ؼ��֣���ѯͬʱ����ĳ����ʹ��'<font  color=<%=Forum_body(8)%>>and</font>'�ָ��ؼ���)<br><br><input type=text size=40 name=keyword></td></tr>
           <tr>
	<td class=tablebody2 valign=middle colspan=2 align=center><b>����ѡ��</b></td></tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>��������</b>&nbsp;<input name=sType type=radio value=1>
           </td>
           <td class=tablebody1 align=left valign=middle>
           <select name=nSearch>
                <option value=1>������������
                 <option value=2>�����ظ�����
                  <option value=3>���߶�����
           </select>
	</td>
           </tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>�ؼ�������</b>&nbsp;<input name=sType type=radio value=2 checked>
           </td>
           <td class=tablebody1 align=left valign=middle>
           <select name=pSearch>
	   <option value=1>�������������ؼ���
	   <!--<option value=2>�����������������ؼ���
	       <option value=3>���߶�����-->
	</select>
	</td>
           </tr>
           <tr>
           <td class=tablebody1 align=right valign=middle>
           <b>���ڷ�Χ</b>
           </td>
           <td class=tablebody1 align=left valign=middle>
<select name=SearchDate class=smallsel> <option value=ALL>��������<option value=1>��������<option  selected value=5>5������<option value=10>10������<option value=30>30������</option></select> 
           <select name=Stable size=1>
<%
For i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""""
	if trim(NowUseBBS)=trim(AllPostTable(i)) then response.write " selected "
	response.write ">"&AllPostTablename(i)&"</option>"
next
%>
           </select>
	 </td>
           </tr>
           <tr>
	<td class=tablebody1 valign=middle colspan=2 align=center><b>��ѡ��Ҫ��������̳ (��Ҫѡ��Щ�� >> �� << �������ģ���ֻ���������������̳)</b></td></tr>
	<tr>
	<td class=tablebody1 colspan=2 valign=middle align=center>
           <b>������̳�� &nbsp; 
<select name=boardid size=1>
<option value=0>������̳</option>
<%
	dim rs1,sql1
	dim searchboard,searchboarda,searchclass,searchinfo
	sql="select id,class from class order by orders"
	set rs=conn.execute(sql)
	if not(rs.eof and rs.bof) then
	do while not rs.eof
	response.write "<option value=0>>>&nbsp;"& rs(1) &" <<"
		sql1="select boardid,boardtype from board where "&guestlist&" class="& rs(0)&" order by orders"
		set rs1=conn.execute(sql1)
		if rs1.eof and rs1.bof then
			searchboard="<option value=0>û����̳"
		else
			do while not rs1.eof
				response.write "<option value="& rs1(0) &""
				if boardid<>"" then
				if cint(boardid)=cint(rs1(0)) then response.write " selected "
				end if
				response.write ">"& rs1(1)
			rs1.movenext
			loop
		end if
	rs.movenext
	loop
	end if
	set rs=nothing
	set rs1=nothing
%>
</select>
	</b></td>
	</tr>
	<tr>
	<td class=tablebody2 valign=middle colspan=2 align=center>
	<input type=submit value=��ʼ����>
	</td></form></tr></table>
<%
	call activeonline()
end sub
%>