<%@ Language=VBScript %>
<%
dim curraction,IID,cate
if request("curraction")="" then
	response.write "<script language=JavaScript>" & chr(13) & "alert('非法登录！');" & "history.back()" & "</script>" 
else
	if session("IID")="" then
	IID=1
	cate="办公室"
	else
		IID=session("IID")
		cate=session("cate")
	end if
	curraction=request("curraction")
end if
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY bgcolor=#99cccc>
<script language="javascript">
function confirmDel(id){
if ( confirm("你确定要删除吗？")) {
window.location.href = "del.asp?id=" + id
  }
}
</script>
<div align=center>
<%
if curraction= "INS" then
%>
<TABLE border=1 cellPadding=1 cellSpacing=1 width="300">
	<TR>
    <TD align=center colspan=2><%Response.Write Cate&"管理"%></TD></TR>
    <TR>
    <TD align=center colspan=2>添加记录</TD></TR>
<FORM action="add.asp" method=post name=FORM1>
  <TR>
    <TD>姓&nbsp;&nbsp;名</TD>
    <TD><INPUT name=text1></TD></TR>
  <TR>
    <TD>性&nbsp;&nbsp;别</TD>
    <TD><SELECT name=sel1>
				<OPTION selected>男</OPTION>
				<OPTION>女</OPTION></SELECT></TD></TR>
  <TR>
    <TD>出生年月</TD>
    <TD><INPUT name=text2></TD></TR>
  <TR>
    <TD>职&nbsp;&nbsp;务</TD>
    <TD><INPUT name=text3></TD></TR>
  <TR>
    <TD>学&nbsp;&nbsp;历</TD>
    <TD><INPUT name=text4></TD></TR>
  <TR>
    <TD align=center colspan=2><INPUT type="submit" value="提交" id=submit1 name=submit1></TD></TR></TABLE>
</FORM>
<%
else
%>
<TABLE WIDTH=75% BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TD>姓名</TD>
		<TD>性别</TD>
		<TD>出生年月</TD>
		<TD>WORKING</TD>
	</TR>
<% 
set rs=server.createobject("adodb.recordset") 
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
sql="select * from mytable where cateid="+IID
rs.open sql,conn,1,1 
if rs.EOF or rs.BOF then
	Response.Write "<tr><td colspan=4>数据库暂无信息</td></tr>"
else
do while not rs.EOF or rs.BOF
%>
<tr><td><%=rs("name")%></td><td><%=rs("sex")%></td><td><%=rs("borndate")%></td>
<td><a href='javascript:confirmDel(<%= rs("id") %>)'>删除</a></td></tr>
<%
rs.movenext
loop
end if
set rs=nothing 
set conn=nothing
%></TABLE>
<%
end if%>
</div>
</BODY>
</HTML>
