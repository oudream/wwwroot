<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&Request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
if rs.bof or rs.eof then
	response.redirect "login.asp"
	response.end
end if
jingyong=rs("jingyong")
rs.close
set rs=nothing

if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="check" or Request.cookies(Forcast_SN)("KEY")="bigmaster" or Request.cookies(Forcast_SN)("KEY")="smallmaster" or Request.cookies(Forcast_SN)("KEY")="typemaster" or (Request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) then
	if request.QueryString("action")="del" then
		conn.execute "delete from news where newsid in ("&request.form("selectdel")&")"
		response.Redirect  "mynews.asp"
	end if 
	PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
	MyPageSize   = 16           'ÿҳ��ʾ����������
		
	If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
		MyPage=1
	Else
		MyPage=Int(Abs(Request("page")))
	End if
	%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="site.css" rel=stylesheet>
<title><%=copyright%><%=version%><%=ver%></title>
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
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr> 
<td colspan=8 class="TDtop" height=25> 
	<div align="center" >�� <B>�ҵ�����</B> ��</div>
</td>
</tr>
<tr>
<td colspan=8 align=center height=25><a href="newsaddd1.asp">�������</a></td>
</tr>
<%set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from "& db_News_Table &" where editor='" & CheckStr(request.cookies(Forcast_SN)("UserName")) & "' order by NewsID desc"
rs.open rs.Source,conn,1,3
if not rs.eof then
	rs.PageSize     = MyPageSize
	MaxPages         = rs.PageCount
	rs.absolutepage = MyPage
	total            = rs.RecordCount
	i = 1
	%>
	<tr align="center" class="TDtop1">
	<td width="7%">ID��</td>
	<td width="39%">�� &nbsp;&nbsp;&nbsp;��</td>
	<td width="16%">��&nbsp;&nbsp;&nbsp;��</td>
	<td width="10%">״̬</td>
	<td width="8%">�༭</td>
	<td width="4%">ѡ��</td>
	</tr>
	<form action="MyNews.asp?action=del" method=post name=check>
	<%do while not rs.eof and (i <= MyPageSize)%>
	<tr bgcolor=ffffff height="20">
	<td width=7% align=center ><%=rs("NewsID")%></td>
	<td width=39% ><a href='ReadNews.asp?NewsID=<%=rs("NewsID")%>' target=_blank title="�鿴���¡�<%=rs("Title")%>��"><font color="<%=rs("titlecolor")%>"><%=left(rs("Title"),40)%></font></a><%if rs("picname")<>"" then%><font color=red>[ͼ]<%end if%></td>
	<td width=16% align=center ><%=rs("UpdateTime")%>��</td>
	<td width=10% align=center ><%if rs("checkked")<>1 then%>δ��<%else%>����<%end if%>��<%if rs("istop")<>1 then%>δ��<%else%>�ѹ�<%end if%></td>
	<td width=8% align=center ><a href="NewsEdit.asp?NewsID=<%=rs("NewsID")%>">�༭</a></td>
	<td width="4%" align=center><input name="selectdel" type="checkbox" id="selectdel" value=<%=rs("NewsID")%>>	</td>
	</tr>
	<%rs.MoveNext
	i = i + 1
	loop%>
	<tr>
	<td colspan="8" align=right height="25" width="100%">
	<input type=checkbox name=chkall value=on onclick="CheckAll(this.form)">ѡ��������ʾ����&nbsp;
	<input type=submit name=action onclick="{if(confirm('ɾ��ѡ����������')){this.document.check.submit();return true;}return false;}" value="ɾ������" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
	</tr>
	</form>
<tr>
    <td colspan=8  align=center height=25>
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> ҳ��ÿҳ <%=MyPageSize%> �� 
<%
url="mynews.asp?"
							
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
	PagePrev=PageShowSize*(PageNextSize-1)
	Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
	Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
	Prev_Page = MyPage - 1
	Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
	PageSizeShow = PageShowSize
Else
	PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
	PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
	if PageLink <> MyPage Then
		Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
	else
		Response.Write "<B>["& PageLink &"]</B> "
	end if
	If PageLink = MaxPages Then Exit for
	Next

if Mypage+1 <=Pagetpage  then
	Next_Page = MyPage + 1
	Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
	PageNext = PageShowSize * PageNextSize + 1
	Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
	Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if
%>
</td>

<%else
	Response.write "<td width='100%' align=center><B>�㻹û�з��������!����������¡�</B></td>"
end if
rs.close
set rs=nothing%>
</tr>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%else%>
	<script language=javascript>
		history.back()
		alert("�Բ��𣬹���Ա��δ��ͨ�����ʺţ����Ժ�")
	</script>
<%end if%>