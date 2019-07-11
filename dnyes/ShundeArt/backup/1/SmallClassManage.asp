<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.ASP"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->

<%
if request.cookies(Forcast_SN)("key")<>"smallmaster" then
	dim bigclassid,bigclassname,typename,typeid
	bigclassid=request("bigclassid")             
             
	if bigclassid="" then
		Show_Err("未指定参数！<br><br><a href='javascript:history.back()'>返回</a>")
		response.end
	else
		if not IsNumeric(bigclassid) then
			Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		end if
	end if

	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="SELECT * From "& db_BigClass_Table &" where bigclassid="&bigclassid
	RS.open sql,Conn,3,3
	bigclassname=rs("bigclassname")
	typeid=rs("typeid")
	master5=rs("bigclassmaster")
	if master5<>"" then
		bigmaster5=split(master5,"|")
		for i=0 to ubound(bigmaster5)
			if request.cookies(Forcast_SN)("username")=trim(bigmaster5(i)) then
				bigclassmaster5=true
				exit for
			else
				bigclassmaster5=false
			end if
		next
	end if
	
	Set rs1 = Server.CreateObject("ADODB.Recordset")
	sql1="SELECT * From "& db_Type_Table &" where typeid="&typeid
	RS1.open sql1,Conn,3,3
	typename=rs1("typename")
	master4=rs1("typemaster")
	if master4<>"" then
		tmaster4=split(master4,"|")
		for i=0 to ubound(tmaster4)
			if request.cookies(Forcast_SN)("username")=trim(tmaster4(i)) then
				typemaster4=true
				exit for
			else
				typemaster4=false
			end if
		next
	end if
	rs1.close
	set rs1=nothing
	rs.close
set rs=nothing
end if
if typemaster4=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or bigclassmaster5=true or request.cookies(Forcast_SN)("key")="smallmaster" or request.cookies(Forcast_SN)("key")="selfreg" then
	%>

<html>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 文章管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->

<center>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="smallclassset.asp?action=update">
<%if request.cookies(Forcast_SN)("key")<>"smallmaster" then%>
<tr class="TDtop">
<td colspan="8">
<a href=typemanage.asp >全部总栏</a>
</td>
</tr>
<tr >
<td colspan="8"class="TDtop1">
&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href="BigClassManage.asp?TypeID=<%=typeid%>"><font color=red><%=typename%></font></a>
</b>
&nbsp;&nbsp;&nbsp;&nbsp;其他总栏：
	<%
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql ="SELECT  * From "& db_Type_Table &" where typeid<>"&typeid&" order by typeorder"
	RS.open sql,Conn,3,3
	do while not rs.EOF
		master1=rs("typemaster")
		if master1<>"" then
			tmaster1=split(master1,"|")
			for i=0 to ubound(tmaster1)
				if request.cookies(Forcast_SN)("username")=trim(tmaster1(i)) then
					typemaster1=true
					exit for
				else
					typemaster1=false
				end if
			next
		end if
		if (request.cookies(Forcast_SN)("key")="typemaster" and typemaster1=true) or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg" then%>
<a href=BigClassManage.asp?typeid=<%=rs("typeid")%>><%=rs("typename")%></a>
		<%end if
		rs.movenext
	loop
	rs.close
	set rs=nothing%>
</td>
</tr>
<tr >
<td colspan="8" class="TDtop1">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href="SmallClassManage.asp?BigClassID=<%=bigclassid%>"><font color=red><%=bigclassname%></font></a>
</b>
&nbsp;&nbsp;&nbsp;&nbsp;其他大类：
	<%if request.cookies(Forcast_SN)("key")<>"bigmaster" then
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_BigClass_Table &" where typeid="&typeid&" and bigclassid<>"&bigclassid&" order by bigclassorder"
		RS.open sql,Conn,3,3
		do while not rs.EOF%>
<a href=SmallClassManage.asp?bigclassid=<%=rs("bigclassid")%>><%=rs("bigclassname")%></a>
			<%rs.movenext
		loop
		rs.close
		set rs=nothing
	else
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_BigClass_Table &" where bigclassid<>"&bigclassid&" order by bigclassorder"
		RS.open sql,Conn,3,3
		do while not rs.EOF
			master=rs("bigclassmaster")
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
			if (bigclassmaster=true and request.cookies(Forcast_SN)("key")="bigmaster") then%>
<a href=SmallClassManage.asp?bigclassid=<%=rs("bigclassid")%>><%=rs("bigclassname")%></a>
			<%end if
			rs.movenext
		loop
		rs.close
		set rs=nothing
	end if%>
</td>
</tr>
<%end if
Set rs6 = Server.CreateObject("ADODB.Recordset")
if request.cookies(Forcast_SN)("key")<>"smallmaster" then
	sql6 ="SELECT * From "& db_SmallClass_Table &" where bigclassid="&bigclassid&" order by SmallClassorder"
else
	sql6 ="SELECT * From "& db_SmallClass_Table &" order by SmallClassorder"
end if
RS6.open sql6,Conn,3,3
%>
<tr align=center >
<td width="25%"></td>
<td width="12%">小类名称</td>
<td width="12%">小类注释</td>
<td width="9%">小类显示</td>
<td width="9%">属于大类</td>
<td width="12%">小类排序<br>(必填项，数字)</td>
<td width="16%">小类管理员<br>(可以多个，用|分隔)</td>
<td width="5%">删除</td>
</tr>
<%
do while not rs6.eof
	if request.cookies(Forcast_SN)("key")="smallmaster" then
		dim smallclassmaster,smallmaster,master2
		master2=rs6("smallclassma")
		if master2<>"" then
			smallmaster=split(master2,"|")
			for i=0 to ubound(smallmaster)
				if request.cookies(Forcast_SN)("username")=trim(smallmaster(i)) then
					smallclassmaster=true
					exit for
				else
					smallclassmaster=false
				end if
			next
		end if
	end if
	set rstype=createobject("adodb.recordset")
	if request.cookies(Forcast_SN)("key")<>"smallmaster" then
		sql="select * from "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder "
	else
		sql="select * from "& db_BigClass_Table &" where bigclassid="&rs6("bigclassid")&" order by bigclassorder "
	end if
	rstype.Open sql,conn,1,3
	if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster4=true) or (request.cookies(Forcast_SN)("key")="bigmaster") or (smallclassmaster=true and request.cookies(Forcast_SN)("key")="smallmaster") or request.cookies(Forcast_SN)("key")="selfreg" then%>
<tr>
<td>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<a href="listnews.asp?smallclassid=<%=rs6("SmallClassID")%>" title="<%if rs6("smallClasszs")<>"" then%><%=rs6("smallClasszs")%><%else%>无<%end if%>"><%=rs6("smallClassName")%></a></td>
<td align=center>
	<input type="hidden" name="smallclassid" value="<%=rs6("smallclassid")%>">
	<input class=text type="text" name="smallclassname" size="10" value="<%=rs6("smallclassname")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<input class=text type="text" name="smallclasszs" size="10" value="<%=rs6("smallclasszs")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<select size="1" name="smallclassview" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>>
	<option <%if rs6("smallclassview")="1" then%> selected <%end if%> value="1">显示</option>
	<option <%if rs6("smallclassview")="0" then%> selected <%end if%> value="0">不显示</option>                             
</select>
</td>
<td align=center>
	<select name="bigclassid" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>>
		<%if (request.cookies(Forcast_SN)("key")="typemaster" and typemaster1=true) or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="selfreg" then
			do while not rstype.EOF%>
	<option value="<%=rstype("bigclassid")%>" <%if rs6("bigclassid")=cint(rstype("bigclassid")) then%>selected<%end if%>><%=rstype("bigclassname")%></option>
				<%
				rstype.MoveNext
			loop
		end if
		if request.cookies(Forcast_SN)("key")="bigmaster" then%>
	<option value="<%=bigclassid%>"><%=bigclassname%></option>
		<%end if
		if request.cookies(Forcast_SN)("key")="smallmaster" and smallclassmaster=true then%>
	<option value="<%=rstype("bigclassid")%>"><%=rstype("bigclassname")%></option>
		<%end if%>
	</select>
</td>
<td align=center><input class=text type="text" name="smallclassorder" size="10" value="<%=rs6("smallclassorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input class=text type="text" name="smallclassma" size="10" value="<%=rs6("smallclassma")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" then%><a href="smallclasskill.asp?smallclassid=<%=rs6("smallclassid")%>">删除</a><%end if%></td>
</tr>
	<%end if
	RS6.MoveNext
Loop
set rstype=nothing
rs6.close
set rs6=nothing
%>
<tr> 
<td colspan="1" height="25" align="center" bgcolor="#FFFFFF">
	<a href="javascript:window.location.reload()" target=mainmenu title="添加栏目类别后请更新左栏菜单项" style="font-family: 宋体; font-size: 9pt">刷新左栏</a>
</td>
<td colspan="7"  align="center" bgcolor="#FFFFFF">
	<input type="submit" name="Submit2" value="保存" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
<%if (request.cookies(Forcast_SN)("key")="bigmaster") or request.cookies(Forcast_SN)("key")="super" or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster1=true) then%>
<form method="post" action="smallclassset.asp?action=add">
<tr >
<td align="center">
添加小类
</td>
<td align=center><input class=text type="text" name="smallclassname" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input class=text type="text" name="smallclasszs" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center>
	<select size="1" name="smallclassview" style="font-family: 宋体; font-size: 9pt">
	<option selected value="1">显示</option>
	<option value="0">不显示</option>                             
	</select></td>
<td align=center>
	<% 
	set rstype=createobject("adodb.recordset")
	if request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="super" then
		sql="select * from "& db_BigClass_Table &" where typeid="&typeid&" order by bigclassorder "
	else
		sql="select * from "& db_BigClass_Table &" order by bigclassorder "
	end if
	rstype.Open sql,conn,1,3%>
	<select name="bigclassid" style="font-family: 宋体; font-size: 9pt">
	<%if request.cookies(Forcast_SN)("key")<>"bigmaster" then
		do while not rstype.EOF%>
	<option <%if bigclassid=trim(rstype("bigclassid")) then%>selected<%end if%> value="<%=rstype("bigclassid")%>"><%=rstype("bigclassname")%></option>
			<%rstype.MoveNext
		loop
	else%>
	<option value="<%=bigclassid%>"><%=bigclassname%></option>
	<%end if%>
	</select>
</td>
<td align=center>
	<input class=text type="text" name="smallclassorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<input class=text type="text" name="smallclassma" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<input type="hidden" name="typeid" value="<%=typeid%>"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<input type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
	<%
	set rstype=nothing
end if
if request.cookies(Forcast_SN)("key")<>"smallmaster" then
	%>
<tr >
<td colspan="8" align=center>
	<input type=button value="在此大类发表新文章" class=button onclick="javascript:window.open('newsadd1.asp?bigclassid=<%=bigclassid%>','_self','')"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<input type=button value="本大类中的无小类的文章" class=button onclick="javascript:window.open('smallno.asp?bigclassid=<%=bigclassid%>','_self','')"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
<%end if%>
</table>
</center>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
<%else
	Show_Err("对不起，你无权操作！<br><br><a href='javascript:history.back()'>返回</a>")
	Response.End 
end if
%>