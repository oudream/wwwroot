<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
if request.cookies(Forcast_SN)("KEY")="smallmaster" then
	response.redirect "SmallClassManage.asp"
	response.end
else
	if request.cookies(Forcast_SN)("key")<>"bigmaster" then
		dim typeid,typename
		typeid=request("typeid")
	
		if typeid="" then
			Show_Err("未指定参数！<br><br><a href='javascript:history.back()'>返回</a>")
			response.end
		else
			if not IsNumeric(typeid) then
				Show_Err("非法参数！<br><br><a href='javascript:history.back()'>返回</a>")
				response.end
			end if
		end if
	
		Set rs = Server.CreateObject("ADODB.Recordset")
		sql ="SELECT  * From "& db_Type_Table &" where typeid="&typeid
		RS.open sql,Conn,3,3
		typename=rs("typename")
		master=rs("typemaster")
		if master<>"" then
			tmaster=split(master,"|")
			for i=0 to ubound(tmaster)
				if request.cookies(Forcast_SN)("username")=trim(tmaster(i)) then
					typemaster=true
					exit for
				else
					typemaster=false
				end if
			next
		end if
		rs.close
		set rs=nothing
	end if
	if typemaster=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="selfreg" then%>
<html>
<head>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 大类管理</title>
<script>
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<LINK href=site.css rel=stylesheet>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
	<%Set rs6 = Server.CreateObject("ADODB.Recordset")
	if request.cookies(Forcast_SN)("key")<>"bigmaster" then
		sql6 ="SELECT  * From "& db_BigClass_Table &" where typeid="&typeid&" order by BigClassorder"
	else
		sql6 ="SELECT  * From "& db_BigClass_Table &" order by BigClassorder"
	end if
	RS6.open sql6,Conn,3,3
	%>
<form method="post" action="BigClassset.asp?action=update">
	<%if request.cookies(Forcast_SN)("key")<>"bigmaster" then%>
<tr class="TDtop">
<td colspan="8"><a href=TypeManage.asp >全部总栏</a></td>
</tr>
<tr class="TDtop1">
<td colspan="8">&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
<b>
<a href=BigClassManage.asp?typeid=<%=typeid%>><font color=red><%=typename%></font></a>
</b>&nbsp;&nbsp;&nbsp;&nbsp;其他总栏：
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
			if (typemaster1=true and request.cookies(Forcast_SN)("key")="typemaster") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg" then%>
<a href=BigClassManage.asp?typeid=<%=rs("typeid")%>><%=rs("typename")%></a>
			<%end if
			rs.movenext
		loop
		rs.close
		set rs=nothing%>
	</td></tr>
	<%end if%>
<tr align=center >
<td width="25%"></td>
<td width="12%">大类名称</td>
<td width="12%">大类注释</td>
<td width="9%">大类显示</td>
<td width="9%">属于总栏</td>
<td width="13%">大类排序<br>(必填项，数字)</td>
<td width="16%">大类管理员<br>(可以多个，用|分隔)</td>
<td width="5%">删除</td>
</tr>
	<%	
	do while not rs6.eof
		dim bigclassmaster,bigmaster,master
		master=rs6("bigclassmaster")
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
	
		set rstype=createobject("adodb.recordset")
		sql="select * from "& db_Type_Table &" order by typeorder"
		rstype.Open sql,conn,1,3
		if bigclassmaster=true or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or (typemaster=true and request.cookies(Forcast_SN)("key")="typemaster") or request.cookies(Forcast_SN)("key")="selfreg" then
			%>
<tr >
<td>
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src=images/menu1.gif>
	<a href="SmallClassManage.asp?BigClassID=<%=rs6("BigClassID")%>" title="<%if rs6("BigClasszs")<>"" then%><%=rs6("bigClasszs")%><%else%>无<%end if%>"><%=rs6("BigClassName")%></a></td>
<td align=center>
	<input type="hidden" name="bigclassid" value="<%=rs6("bigclassid")%>"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<input class=text type="text" name="bigclassname" size="10" value="<%=rs6("bigclassname")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<input class=text type="text" name="bigclasszs" size="10" value="<%=rs6("bigclasszs")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align=center>
	<select size="1" name="bigclassview" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>>
	<option <%if rs6("bigclassview")="1" then%> selected <%end if%>value="1">显示</option>
	<option <%if rs6("bigclassview")="0" then%> selected <%end if%>value="0">不显示</option>                             
	</select>
</td>
<td align=center>
	<select name="typeid" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>>
			<%if request.cookies(Forcast_SN)("key")<>"bigmaster" then
				do while not rstype.EOF
					master2=rstype("typemaster")
					if master2<>"" then
						tmaster2=split(master2,"|")
						for i=0 to ubound(tmaster2)
							if request.cookies(Forcast_SN)("username")=trim(tmaster2(i)) then
								typemaster2=true
								exit for
							else
								typemaster2=false
							end if
						next
					end if
					if (typemaster2=true and request.cookies(Forcast_SN)("key")="typemaster") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg" then%>
	<option value="<%=rstype("typeid")%>" <%if rs6("typeid")=cint(rstype("typeid")) then%>selected<%end if%>><%=rstype("typename")%></option>
					<%end if
					rstype.MoveNext
				loop
			else
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql ="SELECT  * From "& db_Type_Table &" where typeid="&rs6("typeid")
				RS.open sql,Conn,3,3%>
	<option value="<%=rs6("typeid")%>"><%=rs("typename")%></option>
				<%
				rs.close
				set rs=nothing
			end if%>
	</select>
</td>
<td align=center><input class=text type="text" name="bigclassorder" size="10" value="<%=rs6("bigclassorder")%>" ONKEYPRESS="event.returnValue=IsDigit();" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input class=text type="text" name="bigclassmaster" size="10" value="<%=rs6("bigclassmaster")%>" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center>
			<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" then%>
	<a href="bigclasskill.asp?bigclassid=<%=rs6("bigclassid")%>">删除</a>
			<%end if%>
</td>
</tr>
		<%end if
		RS6.MoveNext
	Loop
	set rstype=nothing
	rs6.close
	set rs6=nothing
	%>
<tr>
<td height="25" colspan="1" align="center" >
<a href="javascript:window.location.reload()" target=mainmenu title="添加栏目类别后请更新左栏菜单项" style="font-family: 宋体; font-size: 9pt">刷新左栏</a>
</td> 
<td colspan="7"  align="center" >
	<input type="submit" name="Submit2" value="保存" style="font-family: 宋体; font-size: 9pt" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
	<input type="button" value="添加文章" style="font-family: 宋体; font-size: 9pt" onclick="javascript:window.open('NewsAddd1.asp','_self','')" <%if not(request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="selfreg") then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
	<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" then%>
<form method="post" action="BigClassset.asp?action=add">
<tr >
<td align="center">添加大类</td>
<td align=center><input class=text type="text" name="bigclassname" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input class=text type="text" name="bigclasszs" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center>
	<select size="1" name="bigclassview" style="font-family: 宋体; font-size: 9pt">
	<option selected value="1">显示</option>
	<option value="0">不显示</option>                             
	</select>
</td>
<td align=center>
		<%set rstype=createobject("adodb.recordset")
		sql="select * from "& db_Type_Table &" order by typeorder"
		rstype.Open sql,conn,1,3%>
	<select name="typeid" style="font-family: 宋体; font-size: 9pt">
		<%do while not rstype.EOF
			master3=rstype("typemaster")
			if master3<>"" then
				tmaster3=split(master3,"|")
				for i=0 to ubound(tmaster3)
					if request.cookies(Forcast_SN)("username")=trim(tmaster3(i)) then
						typemaster3=true
						exit for
					else
						typemaster3=false
					end if
				next
			end if
			if request.cookies(Forcast_SN)("key")="super" or (typemaster3=true and request.cookies(Forcast_SN)("key")="typemaster") then%>
	<option <%if typeid=trim(rstype("typeid")) then%>selected<%end if%> value="<%=rstype("typeid")%>"><%=rstype("typename")%></option>
			<%end if
			rstype.MoveNext
		loop
		set rstype=nothing
		%>
	</select>
</td>
<td align=center><input class=text type="text" name="bigclassorder" size="10" ONKEYPRESS="event.returnValue=IsDigit();"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input class=text type="text" name="bigclassmaster" size="10"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td align=center><input type="submit" name="Submit2" value="添加" style="font-family: 宋体; font-size: 9pt"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
</tr>
</form>
	<%end if%>
</table>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		response.write "<script>alert('对不起，你无权操作！');history.go(-1);</Script>"
		Response.End 
	end if
end if
conn.close
set conn=nothing%>