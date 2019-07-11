<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp" -->

<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

NewsID=Request.QueryString("NewsID")
typeid=Request.Form("typeid")

set rs2=server.CreateObject ("ADODB.RecordSet")
if request.cookies(Forcast_SN)("key")="bigmaster" then
	rs2.Source="select * from "& db_BigClass_Table &""
else
	rs2.Source="select * from "& db_BigClass_Table &" where typeid="&typeid
end if
rs2.Open rs2.source,conn,1,1

dim rs
dim sql
dim count
set rs=server.createobject("adodb.recordset")
sql = "select * from "& db_SmallClass_Table &" order by smallclassorder"
rs.open sql,conn,1,1
%>
<script language="JavaScript">
var onecount;
onecount=0;
subcat = new Array();
<%
count = 0
do while not rs.eof
%>
subcat[<%=count%>] = new Array("<%= trim(rs("smallclassname"))%>","<%= trim(rs("bigclassid"))%>","<%= trim(rs("smallclassid"))%>");
<%
	count = count + 1
	rs.movenext
loop
rs.close
%>
onecount=<%=count%>;
function changelocation(locationid)
{
document.form1.smallclassid.length = 0;
document.form1.smallclassid.options[document.form1.smallclassid.length] = new Option("请选择小类", "");
var locationid=locationid;
var i;
for (i=0;i < onecount; i++)
{
if (subcat[i][1] == locationid)
{
document.form1.smallclassid.options[document.form1.smallclassid.length] = new Option(subcat[i][0], subcat[i][2]);
}
}
}
</script>
<script language="JavaScript">
<!--
function na_select_form (fname, type_name) 
{
  document.forms[fname].elements[type_name].select()
  document.forms[fname].elements[type_name].focus()
}
// -->
</script>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 添加文章</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" width="100%" align=center cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>">
<form name=form1 method="post" action="<%if eWebEditor="1" then%>NewsAdd1_eWebEditor.asp<%else%>NewsAdd1.asp<%end if%>">
<tr align="center">
<td colspan="1" class="TDtop" height=25>
	<div align="center" >┊ <B>添加文章--类别选择</B> ┊</div>
</td>
</tr>
<input type=hidden name="typeid" value=<%=typeid%>>
<tr bgcolor="#FFFFFF">
<td align="center">
<p>　</p>
<p>所属小类：
<select name="bigclassid" onChange="changelocation(document.form1.bigclassid.options[document.form1.bigclassid.selectedIndex].value)" size="1">
<option selected value="">请选择大类</option>
<%dim bigclassid,master,bigmaster,bigclassmaster
do while not rs2.eof
bigclassid=rs2("bigclassid")
master=rs2("bigclassmaster")
if master<>"" then
bigmaster=split(master,"|")
 for i=0 to ubound(bigmaster)
	if request.cookies(Forcast_SN)("username")=trim(bigmaster(i)) then
	bigclassmaster=true
	exit for
	else
	bigclassmaster=false
	end if
 next
end if
%>
<option value="<%=trim(rs2("bigclassid"))%>"><%=trim(rs2("bigclassname"))%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<select name="smallclassid" size="1">
<option selected value="">请选择小类</option>
</select>
</p>
<p>　</p>
</td>
</tr>
<tr>
<td align="center" height=25 class="TDtop">
<input type="button" value=" 返 回 " name="B1" onclick=javascript:history.go(-1) style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
<input type="submit" value=" 继 续 " name="B1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->