<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkUser.asp" -->
<!--#include file="ChkManage.asp" -->
<%
IF request.cookies(Forcast_SN)("ManageKEY")<>"super" or request.cookies(Forcast_SN)("ManagePurview")<>"99999" then
	Show_Err("�Բ������ĺ�̨����Ȩ�޲�����")
	response.end
else
	response.buffer=true
	Response.Expires=0

	Set rs9 = Server.CreateObject("ADODB.Recordset")
	sql9 ="SELECT * From "& db_System_Table &" Order By id DESC"
	RS9.open sql9,Conn,1,1
	if rs9("system")="1" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
		%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - ϵͳ����</title>
<style type="text/css">
<!--
.style1 {color: #FF0000}
-->
</style>
<SCRIPT language="JavaScript" type="text/javascript">
// Begin color
	function color(color)
	{
	url = 'color.htm';
	window.open(url,color,"width=430,height=440,status=no,toolbar=no,menubar=no,resizable=no");
	}
// End color-->
</script>
</head>
<body topmargin="0">
<!--#include file="Admin_Top.asp"-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="SystemSave.asp" name="system">
<tr class="TDtop"> 
	<td colspan="3" height=25 width="100%" > 
		<div align="center">�� <strong>��վ��������</strong> ��</div>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>��վ���ƣ�</td>
	<td  colspan="2" align="left" >
		<input type="text" name="name" size="50" value="<%=rs9("name")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"> 
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>��վ��ַ��</td>
	<td  colspan="2" align="left" width="593">
		http://<input type="text" name="xpurl" size="40" value="<%=rs9("xpurl")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">/ (ָϵͳ���ڵ�ַ����http://�������/news)
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" ><font color=red>*</font>�������䣺</td>
	<td  colspan="2" align="left" >
		<input type="text" name="email" size="50" value="<%=rs9("email")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"> 
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >��Ȩ˵����</td>
	<td  colspan="2" align="left" > <font color="#000000"><b>[<%=rs9("Copyright")%>]</b></font></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >�汾��Ϣ��</td>
	<td  colspan="2" align="left" > <font color="#000000"><b>[<%=rs9("version")%>&nbsp;<%=rs9("ver")%>] </b></font></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">TOP�˵�һ������ѡ��</td>
	<td colspan="2" align="left">
		<select name="B_BG" size="1" id="gd1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("B_BG")<>"0" then%>selected<%end if%> value="1">һ��</option>
		<option <%if rs9("B_BG")="0" then%>selected<%end if%> value="0">����</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">LOGOͼ�꣺</td>
	<td colspan="2" align="left">
		<input name="logo" type="text" id="logo" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("logo")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">LOGOͼ��URL��</td>
	<td colspan="2" align="left">
		<input name="logourl" type="text" id="logourl" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("logourl")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">LOGO���ͣ�</td>
	<td colspan="2" align="left">
		<select name="gd1" size="1" id="gd1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("gd1")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd1")="0" then%>selected<%end if%> value="0">flash</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">Banner����</td>
	<td colspan="2" align="left">
		<input name="banner" type="text" id="banner" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("banner")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">Banner��URL��</td>
	<td colspan="2" align="left">
		<input name="bannerurl" type="text" id="bannerurl" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=rs9("bannerurl")%>" size="50">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">Banner�����ͣ�</td>
	<td colspan="2" align="left">
		<select name="gd2" size="1" id="gd2" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("gd2")<>"0" then%>selected<%end if%> value="1">photo</option>
		<option <%if rs9("gd2")="0" then%>selected<%end if%> value="0">flash</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >�ϴ��ļ����ͣ�</td>
	<td  colspan="2" align="left" >
		<input type="text" name="UpFileType" size="50" value="<%=UpFileType%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><font color="#FF0000">��Сд��|���ֿ�</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >�ϴ��ļ���С��</td>
	<td  colspan="2" align="left" >
		<input type="text" name="MaxFileSize" size="50" value="<%=MaxFileSize%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><font color="#FF0000">K</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >���Ա����δ��</td>
	<td height="17" colspan="2" align="left" > 
		<input name="byteType" type="text" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" value="<%=byteType%>" size="50"><font color="#FF0000">��Сд��|���ֿ�</font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="25%"  align="right" bgcolor="#FFFFFF">
		�Զ���TOP�˵���<br><br>
		<font color="#FF0000">HTML��ʽ��д���粻֧��FSO��<br>
		�༭config.asp</font><br>
	</td>
	<td colspan="2" align="left">
		<textarea name="basemenu" cols="80" rows="6" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=basemenu%></textarea><font color="#FF0000"></font>
	</td>
</tr>
<tr bgcolor="#FFFFFF">
	<td width="25%"  align="right" bgcolor="#FFFFFF">
		�Զ���BOTTOM�˵���<br><br>
		<font color="#FF0000">HTML��ʽ��д���粻֧��FSO��<br>
		�༭config.asp</font>
	</td>
	<td colspan="2" align="left">
		<textarea name="BOTTOMmenu" cols="80" rows="6" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'"><%=BOTTOMmenu%></textarea>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">������棺</td>
	<td colspan="2" align="left">
		<select name="R_BG" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("R_BG")<>"0" then%>selected<%end if%> value="1">����</option>
		<option <%if rs9("R_BG")="0" then%>selected<%end if%> value="0">����</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" height="18" align="right" bgcolor="#FFFFFF">���ͼƬ��ַ��</td>
	<td colspan="2" align="left">
		<input type="text" name="R_TOP" size="50" value="<%=rs9("R_TOP")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">������ӣ�</td>
	<td colspan="2" align="left">
		<input type="text" name="L_MAIN" size="50" value="<%=rs9("L_MAIN")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">���˵����</td>
	<td colspan="2" align="left">
		<input type="text" name="R_MAIN" size="50" value="<%=rs9("R_MAIN")%>" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF">����ҳ����أ�</td>
	<td colspan="2" align="left">
		����ҳ���棺<select name="M_BG" size="1" id="M_BG" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<option <%if rs9("M_BG")<>"0" then%>selected<%end if%> value="1">����</option>
		<option <%if rs9("M_BG")="0" then%>selected<%end if%> value="0">����</option>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" bgcolor="#FFFFFF"> <p align="right">����ƹ��棺</p></td>
	<td colspan="2" align="left" >
		<textarea name="gg1" cols="80" rows="4" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'" title=���������뱾վ�Ĺ������� onMouseOver="window.status='���������뱾վ�Ĺ�������';return true;" onMouseOut="window.status='';return true;" ><%=rs9("gg1")%></textarea> 
	</td>
</tr>
<tr class="TDtop"> 
	<td colspan="3"><div align="center">��ʾ��������</div></td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >������ʾ������</td>
	<td colspan="2" align="left">
		<select size="1" name="top_news" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_news")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >����ҳ��������ʾ������</td>
	<td  colspan="2" align="left" >
		<select size="1" name="bigclassshownum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("bigclassshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >ר����ʾ������</td>
	<td  colspan="2" align="left" >
		<select size="1" name="top_sp" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_sp")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >����������ʾ������</td>
	<td  colspan="2" align="left" > <select size="1" name="top_txt" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_txt")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >����ͼ����ʾ������</td>
	<td  colspan="2" align="left" > <select size="1" name="top_img" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("top_img")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >�Ķ�����ҳ������������</td>
	<td  colspan="2" align="left" >
		<select size="1" name="reviewnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("reviewnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >ͼƬJS���µ�������</td>
	<td  colspan="2" align="left" >
		<select size="1" name="picnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("picnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >һ��JS���µ�������</td>
	<td  colspan="2" align="left" >
		<select size="1" name="newsnum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("newsnum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >�û����а���ʾ����</td>
	<td  colspan="2" align="left" >
		<select size="1" name="topuser" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("topuser")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td width="25%" align="right" >��ҳLOGO������ʾ����</td>
	<td  colspan="2" align="left" >
		<select size="1" name="linkshownum" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<%for i=1 to 30%>
			<option <%if rs9("linkshownum")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
		<%next%>
		</select>
	</td>
</tr>
<tr bgcolor="#FFFFFF"> 
	<td colspan="3"  width="611">
		<div align="center"> 
		<input type="submit" name="Submit" value="�ύ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		<input type="reset" name="Submit2" value="����" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
		</div>
	</td>
</tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
		<%rs9.close
		set rs9=nothing
		conn.close
		set conn=nothing
	else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>