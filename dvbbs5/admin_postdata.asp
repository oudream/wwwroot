<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 3.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
	Server.ScriptTimeOut=9999999
	if not master or session("flag")="" then
		Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��½</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
		call dvbbs_error()
	else
		if request("action")="Nowused" then
		call nowused()
		elseif request("action")="update" then
		call update()
		elseif request("action")="del" then
		call del()
		elseif request("action")="CreatTable" then
		call creattable()
		elseif request("action")="search" then
		call search()
		elseif request("action")="update2" then
		call update2()
		elseif request("action")="update3" then
		call update3()
		else
		call main()
		end if
		conn.close
		set conn=nothing
	end if

sub main()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="2" class=Forumrow><B>˵��</B>��<BR>������ѡ����������֮һ��ģʽ�������������ڲ�ͬ��֮���ת����</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>ģʽһ������Ҫת�Ƶ�����</th>
</tr>
<FORM METHOD=POST ACTION="?action=search">
<tr> 
<td height="23" width="20%" class=Forumrow><B>��������</B></td>
<td height="23" width="80%" class=Forumrow><input type="text" name="keyword">&nbsp;
<select name="tablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="checkbox" name="username" value="yes" checked>�û�&nbsp;<input type="checkbox" name="topic" value="yes">����&nbsp;<input type="submit" name="submit" value="����"></td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>ע��</B>��������������ڱ������ͷ����û����ݣ������������������в���</td>
</tr>
<tr> 
<th height="23" colspan="2" align=left>ģʽ�����ڲ�ͬ��ת������</th>
</tr>
<FORM METHOD=POST ACTION="?action=update2">
<tr> 
<td height="23" width="100%" class=Forumrow colspan="2">&nbsp;
<select name="OutTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
 <input type="checkbox" name="top1000" value="yes" checked>ǰ <input type="checkbox" name="end1000" value="yes">�� <input type=text name="selnum" value="100" size=3>�� ��¼ת�Ƶ�
<select name="InTablename">
<%
for i=0 to ubound(AllPostTable)
	response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
next
%>
</select>
&nbsp;<input type="submit" name="submit" value="�ύ">
</td>
</tr>
</FORM>
<tr> 
<td height="23" colspan="2" class=Forumrow><B>ע��</B>����ǰN����¼ָ���ݿ������緢������ӣ����ƽ��ÿ��������5���ظ�����ô100������������ĸ���������500����¼������ͨ��Ҫ���ܳ���ʱ�䣬���µ��ٶ�ȡ�������ķ����������Լ��������ݵĶ��١�ִ�б����轫���Ĵ����ķ�������Դ���������ڷ����������ٵ�ʱ����߱��ؽ��и��²�����</td>
</tr>
</table>
<%
end sub

sub nowused()
%>
<form method="POST" action="?action=update">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="5" class=Forumrow><B>˵��</B>��<BR>�������ݱ���ѡ�е�Ϊ��ǰ��̳��ʹ���������������ݵı�һ�������ÿ�����е�����Խ����̳������ʾ�ٶ�Խ�죬�������е����������ݱ��е������г������������ʱ��������һ�����ݱ��������������ݣ�SQL�汾�û�����ÿ�������ݴﵽ20���Ժ������ӱ�����������ᷢ����̳�ٶȿ�ܶ�ܶࡣ<BR>��Ҳ���Խ���ǰ��ʹ�õ����ݱ����������ݱ����л�����ǰ��ʹ�õ��������ݱ���ǰ��̳�û�����ʱĬ�ϵı����������ݱ�</td>
</tr>
<tr> 
<th height="23" colspan="5">��ǰ���ݱ��趨</th>
</tr>
<tr> 
<td width="20%" class=forumHeaderBackgroundAlternate><b>����<B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>˵��</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>��ǰ����</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>��ǰĬ��</B></td>
<td width="20%" class=forumHeaderBackgroundAlternate><B>ɾ��</B></td>
</tr>
<%
for i=0 to ubound(AllPostTable)
%>
<tr> 
<td width="20%" class=Forumrow><%=AllPostTable(i)%></td>
<td width="20%" class=Forumrow><%=AllPostTableName(i)%></td>
<td width="20%" class=Forumrow>
<%
set rs=conn.execute("select count(*) from "&AllPostTable(i)&"")
response.write rs(0)
%>
</td>
<td width="20%" class=Forumrow><input value="<%=AllPostTable(i)%>" name="TableName" type="radio" <%if NowUseBBS=AllPostTable(i) then%>checked<%end if%>></td>
<td width="20%" class=Forumrow><a href="?action=del&tablename=<%=AllPostTable(i)%>"  onclick="{if(confirm('ɾ�������������ݱ��������ӣ���������ɾ�������ݽ����ɻָ���ȷ��ɾ����?')){return true;}return false;}">ɾ��</a></td>
</tr>
<%
next
%>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</form>
<FORM METHOD=POST ACTION="?action=CreatTable">
<tr> 
<th height="23" colspan="5">������ݱ�</th>
</tr>
<tr> 
<td width="20%" class=Forumrow>��ӵı���</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablename" value="bbs<%=ubound(AllPostTable)+2%>">&nbsp;ֻ����bbs+���ֱ�ʾ����bbs5������������಻�ܳ���9</td>
</tr>
<tr> 
<td width="20%" class=Forumrow>��ӱ��˵��</td>
<td width="80%" class=Forumrow colspan=4><input type=text name="tablereadme">&nbsp;�������ñ����;�����������Ӻ�������ز���������ʾ</td>
</tr>
<tr> 
<td width="100%" colspan=5 class=Forumrow> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</FORM>
</table>
<%
end sub

sub update()
	conn.execute("update config set NowUseBBS='"&request.form("TableName")&"'")
	response.write "���³ɹ���"
end sub

sub del()
	dim nAllPostTable,nAllPostTableName
	if request("tablename")=nowusebbs then
		errmsg="<br><li>��ǰ����ʹ�õı���ɾ����"
		call dvbbs_error()
		exit sub
	end if
	For i=0 to ubound(AllPostTable)
		if trim(request("tablename"))=trim(allposttable(i)) then
		nAllPostTableName=AllPostTableName(i)
		exit for
		end if
	Next
	set rs=conn.execute("select * from config where active=1")
	nAllPostTable=replace(rs("AllPostTable"),"|"&request("tablename"),"")
	nAllPostTableName=replace(rs("AllPostTableName"),"|"&nAllPostTableName,"")
	conn.execute("update config set AllPostTable='"&nAllPostTable&"',AllPostTableName='"&nAllPostTableName&"'")
	conn.execute("drop table "&request("tablename")&"")
	response.write "ɾ���ɹ���"
end sub

sub CreatTable()
if request.form("tablename")="" then
	errmsg="<br><li>�����������"
	call dvbbs_error()
	exit sub
elseif len(request.form("tablename"))<>4 then
	errmsg="<br><li>����ı������Ϸ���"
	call dvbbs_error()
	exit sub
elseif not isnumeric(right(request.form("tablename"),1)) then
	errmsg="<br><li>����ı������Ϸ���"
	call dvbbs_error()
	exit sub
elseif cint(right(request.form("tablename"),1))>9 or cint(right(request.form("tablename"),1))<0 then
	errmsg="<br><li>����ı������Ϸ���"
	call dvbbs_error()
	exit sub
end if
if request.form("tablereadme")="" then
	errmsg="<br><li>��������˵����"
	call dvbbs_error()
	exit sub
end if
for i=0 to ubound(AllPostTable)
	if AllPostTable(i)=request.form("tablename") then
		errmsg="<br><li>������ı����Ѿ����ڣ����������롣"
		call dvbbs_error()
		exit sub
	end if
next

Dim NewAllPostTable,NewAllPostTableName
set rs=conn.execute("select allposttable,allposttablename from config where active=1")
NewAllPostTable=rs(0) & "|" & request.form("tablename")
NewAllPostTableName=rs(1) & "|" & request.form("tablereadme")

'Set conn = Server.CreateObject("ADODB.Connection")
'connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("dvbbs5.mdb")
'conn.open connstr
'�����±�
sql="CREATE TABLE "&request.form("tablename")&" (AnnounceID int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,"&_
		"ParentID int default 0,"&_
		"BoardID int default 0,"&_
		"UserName varchar(50),"&_
		"PostUserID int default 0,"&_
		"Topic varchar(250),"&_
		"Body text,"&_
		"DateAndTime datetime default Now(),"&_
		"length int Default 0,"&_
		"RootID int Default 0,"&_
		"layer int Default 0,"&_
		"orders int Default 0,"&_
		"isbest int Default 0,"&_
		"ip varchar(20) NULL,"&_
		"Expression varchar(20) NULL,"&_
		"locktopic int Default 0,"&_
		"signflag int Default 0,"&_
		"emailflag int Default 0,"&_
		"isagree varchar(50) NULL)"
conn.execute(sql)

'�������
conn.execute("create index dispbbs on "&request.form("tablename")&" (postuserid,boardid,rootid,locktopic,announceid)")
conn.execute("create index save_2 on "&request.form("tablename")&" (rootid,orders)")
conn.execute("create index search on "&request.form("tablename")&" (boardid,dateandtime,topic,announceid)")

conn.execute("update config set AllPostTable='"&NewAllPostTable&"',AllPostTableName='"&NewAllPostTableName&"'")
response.write "��ӱ�ɹ����뷵�ء�"
end sub

sub update2()
dim trs
dim ForNum,TopNum
Dim orderby
if request.form("outtablename")=request.form("intablename") then
	Errmsg="<br><li>��������ͬ���ݱ���ת�����ݡ�"
	call dvbbs_error()
	exit sub
end if
if (not isnumeric(request.form("selnum"))) or request.form("selnum")="" then
	Errmsg="<br><li>����д��ȷ�ĸ���������"
	call dvbbs_error()
	exit sub
end if
if request.form("top1000")="yes" then
orderby=""
else
orderby=" desc"
end if
TopNum=Clng(request.form("selnum"))
if TopNum>100 then
	ForNum=int(TopNum/100)+1
	TopNum=100
else
	ForNum=1
end if

for i=1 to ForNum
set rs=conn.execute("select top "&TopNum&" topicid from topic where PostTable='"&request.form("outtablename")&"' order by topicid "&orderby&"")
if rs.eof and rs.bof then
	Errmsg="<br><li>����ѡ�񵼳������ݱ��Ѿ�û���κ�����"
	call dvbbs_error()
	exit sub
else
	do while not rs.eof
	'��ȡ�����������ݱ�
	set trs=conn.execute("select * from "&request.form("outtablename")&" where rootid="&rs("topicid"))
	if not (trs.eof and trs.bof) then
	do while not trs.eof
	'���뵼���������ݱ�
	conn.execute("insert into "&request("intablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&trs("boardid")&","&trs("parentid")&",'"&checkstr(trs("username"))&"','"&checkstr(trs("topic"))&"','"&checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&trs("postuserid")&")")
	'��������ָ�����ӱ�
	conn.execute("update topic set PostTable='"&request.form("inTableName")&"' where TopicID="&rs("topicid"))
	'ɾ�������������ݱ��Ӧ����
	conn.execute("delete from "&request.form("outTableName")&" where AnnounceID="&trs("Announceid"))
	trs.movenext
	loop
	end if
	rs.movenext
	loop
end if
next
set trs=nothing
set rs=nothing
response.write "���³ɹ���"
end sub

sub search()
dim keyword
dim totalrec
dim n
dim currentpage,page_count,Pcount
currentPage=request("page")
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
end if
if request("keyword")="" then
	Errmsg="<br><li>��������Ҫ��ѯ�Ĺؼ��֡�"
	call dvbbs_error()
	exit sub
else
	keyword=replace(request("keyword"),"'","")
end if
if request("username")="yes" then
sql="select * from topic where PostTable='"&request("tablename")&"' and postusername='"&keyword&"' order by LastPostTime desc"
elseif request("topic")="yes" then
sql="select * from topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
else
	Errmsg="<br><li>��ѡ������ѯ�ķ�ʽ��"
	call dvbbs_error()
	exit sub
end if
%>
<form method="POST" action="?action=update3">
<input type=hidden name="topic" value="<%=request("topic")%>">
<input type=hidden name="username" value="<%=request("username")%>">
<input type=hidden name="keyword" value="<%=keyword%>">
<input type=hidden name="tablename" value="<%=request("tablename")%>">
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
<tr> 
<td height="23" colspan="6" class=Forumrow><B>˵��</B>��<BR>�����Զ����е������������ת�����ݱ�Ĳ�������������ͬ���ڽ���ת��������</td>
</tr>
<tr> 
<th height="23" colspan="6">����<%=request("tablename")%>���</th>
</tr>
<tr> 
<td width="6%" class=forumHeaderBackgroundAlternate align=center><b>״̬<B></td>
<td width="45%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
<td width="15%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>�ظ�</B></td>
<td width="22%" class=forumHeaderBackgroundAlternate align=center><B>ʱ��</B></td>
<td width="6%" class=forumHeaderBackgroundAlternate align=center><B>����</B></td>
</tr>
<%
set rs=server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
	response.write "<tr> <td class=Forumrow colspan=6 height=25>û��������������ݡ�</td></tr>"
else
	rs.PageSize = Forum_Setting(11)
	rs.AbsolutePage=currentpage
	page_count=0
	totalrec=rs.recordcount
	while (not rs.eof) and (not page_count = rs.PageSize)
%>
<tr> 
<td width="6%" class=Forumrow align=center>
<%
if rs("locktopic")=1 then
	response.write "����"
elseif rs("isvote")=1 then
	response.write "ͶƱ"
elseif rs("isbest")=1 then
	response.write "����"
else
	response.write "����"
end if
%>
</td>
<td width="45%" class=Forumrow><%=htmlencode(rs("title"))%></td>
<td width="15%" class=Forumrow align=center><a href="admin_user.asp?action=modify&userid=<%=rs("postuserid")%>"><%=htmlencode(rs("postusername"))%></a></td>
<td width="6%" class=Forumrow align=center><%=rs("child")%></td>
<td width="22%" class=Forumrow><%=rs("dateandtime")%></td>
<td width="6%" class=Forumrow align=center><input type="checkbox" name="topicid" value="<%=rs("topicid")%>"></td>
</tr>
<%
  	page_count = page_count + 1
	rs.movenext
	wend
	dim endpage
	Pcount=rs.PageCount
	response.write "<tr><td valign=middle nowrap colspan=2 class=forumrow height=25>&nbsp;&nbsp;��ҳ�� "

	if currentpage > 4 then
	response.write "<a href=""?page=1&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">[1]</a> ..."
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
        response.write " <a href=""?page="&i&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&i&"]</a>"
		end if
	end if
	next
	if currentpage+3 < Pcount then 
	response.write "... <a href=""?page="&Pcount&"&action=search&keyword="&keyword&"&topic="&request("topic")&"&username="&request("username")&"&tablename="&request("tablename")&""">["&Pcount&"]</a>"
	end if
	response.write "</td>"
	response.write "<td colspan=3 class=forumrow>���н��<input type=checkbox name=allsearch value=yes>"
	response.write "&nbsp;<select name=toTablename>"

	for i=0 to ubound(AllPostTable)
		response.write "<option value="""&AllPostTable(i)&""">"&AllPostTableName(i)& "--" &AllPostTable(i)&"</option>"
	next

	response.write "</select>&nbsp;<input type=submit name=submit value=ת��>"
	response.write "</td>"
	response.write "<td class=forumrow align=center><input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">"
	response.write "</td></tr>"
end if
rs.close
set rs=nothing
response.write "</table></form><BR><BR>"
end sub

sub update3()
dim keyword,trs
if request.form("tablename")=request.form("totablename") then
	Errmsg="<br><li>��������ͬ���ݱ��ڽ�������ת����"
	call dvbbs_error()
	exit sub
end if
if request.form("allsearch")="yes" then
	if request("keyword")="" then
		Errmsg="<br><li>��������Ҫ��ѯ�Ĺؼ��֡�"
		call dvbbs_error()
		exit sub
	else
		keyword=replace(request("keyword"),"'","")
	end if
	if request("username")="yes" then
		sql="select topicid from topic where PostTable='"&request("tablename")&"' and postusername='"&keyword&"' order by LastPostTime desc"
	elseif request("topic")="yes" then
		sql="select topicid from topic where PostTable='"&request("tablename")&"' and title like '%"&keyword&"%' order by LastPostTime desc"
	else
		Errmsg="<br><li>��ѡ������ѯ�ķ�ʽ��"
		call dvbbs_error()
		exit sub
	end if
else
	if request.form("topicid")="" then
		Errmsg="<br><li>��ѡ��Ҫת�Ƶ����ӡ�"
		call dvbbs_error()
		exit sub
	end if
	sql="select topicid from topic where PostTable='"&request("tablename")&"' and TopicID in ("&request.form("TopicID")&")"
end if

set rs=conn.execute(sql)
if rs.eof and rs.bof then
	Errmsg="<br><li>û���κμ�¼��ת����"
	call dvbbs_error()
	exit sub
else
	do while not rs.eof
	'ȡ��ԭ������
	set trs=conn.execute("select * from "&request("tablename")&" where rootid="&rs("topicid"))
	if not (trs.eof and trs.bof) then
	'�����±�
	do while not trs.eof
	conn.execute("insert into "&request("totablename")&"(Boardid,ParentID,username,topic,body,DateAndTime,length,rootid,layer,orders,ip,Expression,locktopic,signflag,emailflag,isbest,PostUserID) values ("&trs("boardid")&","&trs("parentid")&",'"&checkstr(trs("username"))&"','"&checkstr(trs("topic"))&"','"&checkstr(trs("body"))&"','"&trs("dateandtime")&"',"&trs("length")&","&trs("rootid")&","&trs("layer")&","&trs("orders")&",'"&trs("ip")&"','"&trs("Expression")&"',"&trs("locktopic")&","&trs("signflag")&","&trs("emailflag")&","&trs("isbest")&","&trs("postuserid")&")")
	trs.movenext
	loop
	'���¸��������
	conn.execute("update topic set PostTable='"&request("totablename")&"' where topicid="&rs("topicid"))
	'ɾ��ԭ�����������
	conn.execute("delete from "&request("tablename")&" where rootid="&rs("topicid"))
	end if
	rs.movenext
	loop
end if
set trs=nothing
set rs=nothing
response.write "���³ɹ���"
end sub
%>
<script language="javascript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>