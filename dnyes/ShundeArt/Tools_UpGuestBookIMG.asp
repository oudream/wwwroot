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
	Show_Err("�Բ������ĺ�̨����Ȩ�޲���������")
	response.end
else
	%>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr class="TDtop"> 
	<td  colspan="4" width="100%" height=25 align="center" valign="middle" height="24">
		�� <strong>���¾ɰ�����ͷ������</strong> ��
	</td>
</tr>
<tr> 
	<td colspan="4" align="center" width="100%" height=25 bgcolor="#FFFFFF">ע�⣬���д˳���ǰ���ȱ������ݿ⡣���Ҵ˳���ֻ������һ�Ρ��мǣ�</td>
</tr>
<form method="post" action="Tools_UpGuestBookIMG.asp" name="type">
<tr> 
	<td align="center" colspan="4" width="100%"  height="34">
	<%action=request("action")%>
		<input name="action" type="hidden" value="OK">
		<input type="submit" name="Submit" value="����" title="���и��³���" onMouseOver="window.status='���и��³���';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;&nbsp;
		<input type="reset" value="����" name="B1" onMouseOver="window.status='������д����';return true;" onMouseOut="window.status='';return true;" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
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
			response.write "<DIV align=center><font color=red>��ϵͳ��û�����ԣ�</font></DIV><BR>"
		else
			do while not rs.eof
				face1=rs("face")
				response.write "<DIV align=center>ԭ:["& face1 &"]</DIV>"
				face1=replace(face1,"images/","",1,-1,1)
				rs("face")="images/"& face1
				rs.movenext
				i=i+1
				response.write "<DIV align=center>��:[images/"& face1 &"]</DIV><BR>"
			loop
			rs.close
			set rs=nothing
		end if
		response.write "<DIV align=center><font color=green>��"& i &"���ɰ��������ͷ��������ϣ�</font></DIV><BR>"
	end if%>
	</td>
</tr>
</form>
</table>
<%end if%>