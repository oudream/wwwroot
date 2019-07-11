<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%
IF request.cookies(Forcast_SN)("KEY")="bigmaster" THEN
	response.redirect "BigClassManage.asp"
	response.end
else
	IF request.cookies(Forcast_SN)("KEY")="smallmaster" THEN
		response.redirect "SmallClassManage.asp"
		response.end
	else
		%>
<html>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<head>
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
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form action ="typeset.asp?action=update" method=post>
<tr height="20"> 
	<td width="100%" height="20"  colspan="8" align="left" valign="middle" class="TDtop">
		<a href=typemanage.asp>全部总栏</a>（如果要修改总栏资料，只要修改完毕后按“保存”即可；此为批量修改）
	</td>
</tr>
		<%
		Set rs6 = Server.CreateObject("ADODB.Recordset")
		sql6 ="select * from "& db_Type_Table &" order by typeorder"
		RS6.open sql6,Conn,3,3
		%>
<tr align="center"  height="25" class="TDtop1"> 
<td width="6%"  >ID号</td>
<td width="13%" >总栏名称</td>
<td width="25%"  >总栏注释（前为总栏名，后为注释）</td>
<td width="8%"  >总栏显示</td>
<td width="10%"  >总栏模板</td>
<td width="8%"  >排列顺序<br>
  (必填项，数字)</td>
<td width="16%"  >管理员<br>
  (可以多个，用|分隔)</td>
<td width="6%"  >删除</td>
</tr>
		<%
		do while not rs6.eof
			dim typemaster,tmaster,master
			master=rs6("typemaster")
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
			if (typemaster=true and request.cookies(Forcast_SN)("key")="typemaster") or request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="selfreg" or request.cookies(Forcast_SN)("key")="check" then%>
<tr>
<td align="center" bgcolor="#FFFFFF"><%=rs6("typeid")%>
	<input type=hidden name="typeid" value="<%=rs6("typeid")%>">
</td>
<td align="center" bgcolor="#FFFFFF">
	<a href="BigClassManage.asp?typeid=<%=rs6("typeid")%>" title="<%if rs6("typecontent")<>"" then%><%=rs6("typecontent")%><%else%>无<%end if%>"><%=rs6("typeName")%></a>
</td>
<td align="center" bgcolor="#FFFFFF">
	<input class=text type="text" name="typename" size="13" value="<%=rs6("typename")%>" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<input class=text type="text" name="typecontent" size="14" value="<%=rs6("typecontent")%>" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td align="center" bgcolor="#FFFFFF">
	<select size="1" name="typeview" style="font-family: 宋体; font-size: 9pt" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>>
	<option <%if rs6("typeview")=1 then%>selected<%end if%> value="1">显示</option>
	<option <%if rs6("typeview")=0 then%>selected<%end if%> value="0">不显示</option>
	</select>
</td>
<td align="center" bgcolor="#FFFFFF">
	<select size="1" name="mode" style="font-family: 宋体; font-size: 9pt" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>>
	<option <%if rs6("mode")=1 then%>selected<%end if%> value="1">图片产品模板</option>
	<option <%if rs6("mode")=2 then%>selected<%end if%> value="2">新闻媒体模板</option>
	<option <%if rs6("mode")=3 then%>selected<%end if%> value="3">网址推荐模板</option>
	<option <%if rs6("mode")=4 then%>selected<%end if%> value="4">软件文章模板</option>
	</select>
</td>
<td align="center" bgcolor="#FFFFFF">
	<input class=text type="text" name="typeorder" size="5" style="font-family: 宋体; font-size: 9pt" value="<%=rs6("typeorder")%>" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%> ONKEYPRESS="event.returnValue=IsDigit();"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td bgcolor="#FFFFFF" align="center">
	<input class=text type="text" name="typemaster" size="18" value="<%=rs6("typemaster")%>" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
<td bgcolor="#FFFFFF" align="center"><%if request.cookies(Forcast_SN)("key")="super" then%><a href="typekill.asp?typeID=<%=rs6("typeID")%>">删除</a><%end if%></td>
</tr>
			<%end if
			RS6.MoveNext
		Loop
		rs6.close
		set rs6=nothing
		%>
<tr> 
<td colspan="2" height="40" align="center" width="100%" bgcolor="#FFFFFF">
	<a href="javascript:window.location.reload()" target=mainmenu title="添加栏目类别后请更新左栏菜单项" style="font-family: 宋体; font-size: 9pt">刷新左栏</a>
</td>
<td colspan="6" height="40" align="center" width="100%" bgcolor="#FFFFFF">
	<input type="submit" name="Submit2" value="保存" style="font-family: 宋体; font-size: 9pt" <%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;&nbsp;&nbsp;
	<input type="button" value="添加文章" style="font-family: 宋体; font-size: 9pt" onclick="javascript:window.open('newsaddd1.asp','_self','')"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
</td>
</tr>
</form>
		<%if request.cookies(Forcast_SN)("key")="super" then%>
<form method="post" action="typeset.asp?action=add" name="type">
<tr>
<td align="center" bgcolor="#FFFFFF"></td>
<td align="center" bgcolor="#FFFFFF">添加总栏</td>
<td align="center" bgcolor="#FFFFFF"><input class=text type="text" name="typename" size="13"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
        <input class=text type="text" name="typecontent" size="14" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">          </td>
<td align="center" bgcolor="#FFFFFF">
	<select size="1" name="typeview" style="font-family: 宋体; font-size: 9pt" 
			<%if request.cookies(Forcast_SN)("key")<>"super" then%>disabled<%end if%>>
	<option value="1">显示</option>
	<option value="0">不显示</option>
	</select>
</td>
<td align="center" bgcolor="#FFFFFF">
	<select size="1" name="mode" style="font-family: 宋体; font-size: 9pt">
        <option value="1">图片产品模板</option>
        <option value="2">新闻媒体模板</option>
        <option value="3">网址推荐模板</option>
        <option value="4">软件文章模板</option>
      	</select>
</td>
<td align="center" bgcolor="#FFFFFF"><input class=text type="text" name="typeorder" size="5" style="font-family: 宋体; font-size: 9pt" ONKEYPRESS="event.returnValue=IsDigit();"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td bgcolor="#FFFFFF" align="center"><input class=text type="text" name="typemaster" size="18"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
<td bgcolor="#FFFFFF" align="center"><input type="submit" name="Submit" value="添加" style="font-family: 宋体; font-size: 9pt"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"></td>
</tr>
</form>
</table>
</center>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%
		end if
	end if
	conn.close 
	set conn=nothing
end if%>