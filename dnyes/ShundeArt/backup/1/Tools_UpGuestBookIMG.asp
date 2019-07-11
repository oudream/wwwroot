<!--#include file=conn.asp -->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->
<!--#include file="ChkManage.asp"-->
<%
if not(request.cookies(Forcast_SN)("ManageKEY")="super" or request.cookies(Forcast_SN)("ManageKEY")="typemaster" or request.cookies(Forcast_SN)("ManageKEY")="bigmaster") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td  colspan="4" width="100%" height=25 align="center" valign="middle" height="24">
		┊ <strong>更新旧版留言头像问题</strong> ┊
	</td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" height=25 bgcolor="#FFFFFF">注意，运行此程序前请先备份数据库。并且此程序只能运行一次。切记！</td>
</tr>
<form method="post" action="Tools_UpGuestBookIMG.asp" name="type">
<tr> 
	<td align="center" colspan="4" width="100%"  height="34">
	<%action=request("action")%>
		<input name="action" type="hidden" value="OK">
		<input type="submit" name="Submit" value="更新" title="运行更新程序" onMouseOver="window.status='运行更新程序';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;
		<input type="reset" value="重置" name="B1" onMouseOver="window.status='重置填写内容';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" bgcolor="#FFFFFF">
	<%if action="OK" then
		i=0
		set rs=server.createobject("adodb.recordset")
		sql="select * from "& db_Review_Table &" where NewsID < 1 or (NewsID is null)" 
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "<DIV align=center><font color=red>本系统还没有留言！</font></DIV><BR>"
		else
			do while not rs.eof
				face1=rs("face")
				response.write "<DIV align=center>原:["& face1 &"]</DIV>"
				face1=replace(face1,"images/","",1,-1,1)
				rs("face")="images/"& face1
				rs.movenext
				i=i+1
				response.write "<DIV align=center>现:[images/"& face1 &"]</DIV><BR>"
			loop
			rs.close
			set rs=nothing
		end if
		response.write "<DIV align=center><font color=green>共"& i &"条旧版雁过留声头象修正完毕！</font></DIV><BR>"
	end if%>
	</td>
</tr>
</form>
</table>
<%end if%>