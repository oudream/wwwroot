<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
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
	if rs9("powermana")="1" or request.cookies(Forcast_SN)("ManagePurview")="99999" then
	%>
<html>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - Ȩ������</title>
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div align="center">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<form method="post" action="SystemSave1.asp" name="system">
<tr class="TDtop"> 
          <td colspan="3"  width="100%" height=25> 
            <div align="center">�� <strong>��վ��������</strong> ��</div></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%"  align="right">����ҳ����Ϣ��Ȩ������</td>
          <td  colspan="2" align="left"><select name="topbg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("topbg")<>"0" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("topbg")="0" then%>selected<%end if%> value="0">����</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >������Ŀ������</td>
          <td  colspan="2" align="left" > <select name="t_bg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("t_bg")<>"0" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("t_bg")="0" then%>selected<%end if%> value="0">����</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ʹ�÷ּ�������ƣ�</td>
          <td  colspan="2" align="left" > <select name="uselevel" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("uselevel")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("uselevel")="0" then%>selected<%end if%> value="0">����</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >����Ĭ��״̬��</td>
          <td  colspan="2" align="left" > <select name="usecheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("usecheck")="1" then%>selected<%end if%> value="1">�����</option>
              <option <%if rs9("usecheck")="0" then%>selected<%end if%> value="0">δ���</option>
            </select>
            �����<a title=����Ĭ��״̬��ϸ˵�� onMouseOver="window.status='����Ĭ��״̬��ϸ˵��';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("�����ָС�����Ա������º����²�\n�ؾ���ϵͳ����Ա���������Ա����˾�\n����ʾ��������δ���ָС�����Ա���\n���º����²�ֱ����ʾ���������뾭��\nϵͳ����Ա���������Ա����˺�Ż���\nʾ��ϵͳ����Ա���������Ա���ܴ��ޡ�")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >¼��Ա���޸���������£�</td>
          <td  colspan="2" align="left" > <select name="modnewsable" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("modnewsable")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("modnewsable")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select>
            �����<a title=¼��Ա���޸������������ϸ˵�� onMouseOver="window.status='¼��Ա���޸������������ϸ˵��';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("������Ĭ��״̬Ϊδ���ʱ����ѡ���\n��ʾ�����޸ģ����޸ĵ����½���Ϊδ��\n��״̬�����ڰ�ȫ�Ŀ��ǣ����رձ�ʾ��\n���޸ģ�������Ĭ��״̬Ϊ�����ʱ����\nѡ����Ч��")>����</a>���鿴��ϸ˵����</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >����ע��״̬��</td>
          <td  colspan="2" align="left" > <select name="zs" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("zs")="0" then%>selected<%end if%> value="0">δ��</option>
            </select>
            �����<a title=����ע��״̬��ϸ˵�� onMouseOver="window.status='����ע��״̬��ϸ˵��';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("����ע�����ú󣬽����ڲ鿴��������ʱ\n���ִ������ʾ��Ϣ��δ�����򲻳��֡�\n��Ȼ��ǰ���Ǹô���Ҫ����ʾ��Ϣ��")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С��ע��״̬��</td>
          <td  colspan="2" align="left" > <select name="zs1" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs1")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("zs1")="0" then%>selected<%end if%> value="0">δ��</option>
            </select>
            �����<a title=С��ע��״̬��ϸ˵�� onMouseOver="window.status='С��ע��״̬��ϸ˵��';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("С��ע�����ú󣬽����ڲ鿴С������ʱ\n����С�����ʾ��Ϣ��δ�����򲻳��֡�\n��Ȼ��ǰ���Ǹ�С��Ҫ����ʾ��Ϣ��")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ר��ע��״̬��</td>
          <td  colspan="2" align="left" > <select name="zs2" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("zs2")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("zs2")="0" then%>selected<%end if%> value="0">δ��</option>
            </select>
            �����<a title=ר��ע��״̬��ϸ˵�� onMouseOver="window.status='ר��ע��״̬��ϸ˵��';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("ר��ע�����ú󣬽����ڲ鿴ר������ʱ\n����ר�����ʾ��Ϣ��δ�����򲻳��֡�\n��Ȼ��ǰ���Ǹ�ר��Ҫ����ʾ��Ϣ��")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ������۹��ܣ�</td>
          <td  colspan="2" align="left" > <select name="reviewable" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewable")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("reviewable")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select></td>
        </tr>
		<tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ע���û�ǩ�չ��ܣ�</td>
          <td  colspan="2" align="left" > <select name="M_MAIN" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("M_MAIN")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("M_MAIN")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select><font color=red> ��ȷ�����۹����Ѿ�����</font></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾ������IP��ַ��</td>
          <td  colspan="2" align="left" > <select name="showip" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showip")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("showip")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >����Ĭ��״̬��</td>
          <td  colspan="2" align="left" > <select name="reviewcheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewcheck")="1" then%>selected<%end if%> value="1">�����</option>
              <option <%if rs9("reviewcheck")="0" then%>selected<%end if%> value="0">δ���</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >������۲˵��У�</td>
          <td  colspan="2" align="left" > <select name="reviewcheckshow" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reviewcheckshow")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("reviewcheckshow")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select>
            ���������</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ��������������빦�ܣ�</td>
          <td  colspan="2" align="left" > <select name="linkreg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkreg")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("linkreg")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�������ӹ���˵��У�</td>
          <td  colspan="2" align="left" > <select name="linkpass" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkpass")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("linkpass")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select>
            �������վ����</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >������²˵��У�</td>
          <td  colspan="2" align="left" > <select name="checkshow" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkshow")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("checkshow")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select>
            ���������</td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ�Ƿ���ʾ��Ա���а�</td>
          <td  colspan="2" align="left" > <select name="showuser" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showuser")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showuser")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ�Ƿ���ʾר�����£�</td>
          <td  colspan="2" align="left" > <select name="showspecial" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showspecial")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showspecial")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ�Ƿ���ʾϵͳ���ݣ�</td>
          <td  colspan="2" align="left" > <select name="showdata" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showdata")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showdata")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾ��������</td>
          <td  colspan="2" align="left" > <select name="showsearch" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showsearch")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showsearch")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >����������ʽ��</td>
          <td  colspan="2" align="left" > <select name="search" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("search")="1" then%>selected<%end if%> value="1">��Լ��</option>
              <option <%if rs9("search")="0" then%>selected<%end if%> value="0">��ϸ��</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�ڱ�����Ƿ���ʾ���ߣ�</td>
          <td  colspan="2" align="left" > <select name="showauthor" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showauthor")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showauthor")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td width="25%"  align="right">�Ƿ���ʾ��̳��ڼ���̳���£�</td>
          <td  colspan="2" align="left"><select name="topfont" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("topfont")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("topfont")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾͶƱϵͳ��</td>
          <td  colspan="2" align="left" > <select name="showvote" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showvote")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showvote")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾ�������ӣ�</td>
          <td  colspan="2" align="left" > <select name="showlink" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showlink")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showlink")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾLOGO���ӣ�</td>
          <td  colspan="2" align="left" > <select name="showlinkmap" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showlinkmap")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showlinkmap")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾ��Ա��¼��Ŀ��</td>
          <td  colspan="2" align="left" > <select name="showclub" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showclub")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showclub")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ���ʾ������������������</td>
          <td  colspan="2" align="left" > <select name="showcount" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showcount")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showcount")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ�Ƿ���ʾ���·���ʱ�䣺</td>
          <td  colspan="2" align="left" > <select name="showtime" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showtime")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showtime")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ���·���ʱ����ʾ��ʽ��</td>
          <td  colspan="2" align="left" > <select name="showyear" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showyear")="1" then%>selected<%end if%> value="1">������</option>
              <option <%if rs9("showyear")="0" then%>selected<%end if%> value="0">����</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >��ҳ�����Ƿ���ʾ�������</td>
          <td  colspan="2" align="left" > <select name="showclick" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("showclick")="1" then%>selected<%end if%> value="1">��ʾ</option>
              <option <%if rs9("showclick")="0" then%>selected<%end if%> value="0">����ʾ</option>
            </select></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ִ�в�����ȴ�ʱ�䣺</td>
          <td  colspan="2" align="left" > <select size="1" name="freetime" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <%for i=0 to 30%>
              <option <%if rs9("freetime")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
              <%next%>
            </select>
            �� </td>
        </tr>
        <tr class="TDtop"> 
          <td colspan="3"><div align="center">����Ȩ������</div></td>
        </tr>
        <%if request.cookies(Forcast_SN)("ManageKEY")="super" and request.cookies(Forcast_SN)("ManagePurview")="99999" then%>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����Աϵͳ���ù��ܣ�</td>
          <td  colspan="2" align="left" > <select name="system" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("system")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("system")="0" then%>selected<%end if%> value="0">����</option>
            </select>
            �����<a title=ϵͳ����Աϵͳ���ù�����ϸ˵�� onMouseOver="window.status='ϵͳ����Աϵͳ���ù�����ϸ˵��';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("��ϵͳ����Աϵͳ���ù�������Ϊ����ʱ��ϵͳ����Ա�ܽ���ϵͳ���ã�����ѡ��������в��������ڣ���ϵͳ����Աϵͳ���ù�������Ϊ����ʱ��ϵͳ����Ա���ܽ������á�")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����ԱCSS�༭���ܣ�</td>
          <td  colspan="2" align="left" > <select name="css" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("css")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("css")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����Աϵͳ��ʼ�����ܣ�</td>
          <td  colspan="2" align="left" > <select name="init" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("init")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("init")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����Ա����Ȩ�����ã�</td>
          <td  colspan="2" align="left" > <select name="powermana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("powermana")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("powermana")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����ԱͶƱ�����ܣ�</td>
          <td  colspan="2" align="left" > <select name="votemana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("votemana")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("votemana")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ϵͳ����Ա�������ӹ���</td>
          <td  colspan="2" align="left" > <select name="linkmana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("linkmana")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("linkmana")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <%end if%>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�������Ա����С��ѡ�</td>
          <td  colspan="2" align="left" > <select name="shsmallgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shsmallgl")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("shsmallgl")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�������Ա����ר��ѡ�</td>
          <td  colspan="2" align="left" > <select name="shspecialgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shspecialgl")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("shspecialgl")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�������Ա��������ѡ�</td>
          <td  colspan="2" align="left" > <select name="shdelreview" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("shdelreview")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("shdelreview")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >���Ա�޸���������Ȩ�ޣ�</td>
          <td  colspan="2" align="left" > <select name="checkmod" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkmod")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("checkmod")="0" then%>selected<%end if%> value="0">��ֹ</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >���Աɾ����������Ȩ�ޣ�</td>
          <td  colspan="2" align="left" > <select name="checkdel" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("checkdel")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("checkdel")="0" then%>selected<%end if%> value="0">��ֹ</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С�����Ա����С��ѡ�</td>
          <td  colspan="2" align="left" > <select name="smallgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("smallgl")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("smallgl")="0" then%>selected<%end if%> value="0">����</option>
            </select>
            ���ú�¼��Աֻ�ܹ����Լ���ӵ�С�࣬�����û���ӵ�С���޷����� </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С�����Ա����ר��ѡ�</td>
          <td  colspan="2" align="left" > <select name="specialgl" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("specialgl")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("specialgl")="0" then%>selected<%end if%> value="0">����</option>
            </select>
            ���ú�¼��Աֻ�ܹ����Լ���ӵ�ר�⣬�����û���ӵ�ר���޷����� </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С�����Ա��������ѡ�</td>
          <td  colspan="2" align="left" > <select name="delreview" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("delreview")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("delreview")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С�����Ա�޸�����ѡ�</td>
          <td  colspan="2" align="left" > <select name="inputmodpwd" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("inputmodpwd")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("inputmodpwd")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >С�����Ա����С�����£�</td>
          <td  colspan="2" align="left" > <select name="smallmana" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("smallmana")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("smallmana")="0" then%>selected<%end if%> value="0">����</option>
            </select>
            ���ú�С�����Ա�ɹ���С����������£����ú�С�����Ա���ܹ��������� </td>
        </tr>
        <tr class="TDtop"> 
          <td colspan="3"> <div align="center" >�û�ע������</div></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >�Ƿ������û���ע�᣺</td>
          <td  colspan="2" align="left" > <select name="reg" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("reg")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("reg")="0" then%>selected<%end if%> value="0">����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ע����û��Ƿ��ܷ������£�</td>
          <td  colspan="2" align="left" > <select name="fabiao" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiao")="1" then%>selected<%end if%> value="1">��������</option>
              <option <%if rs9("fabiao")="0" then%>selected<%end if%> value="0">�ȴ����</option>
            </select> </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ע���û���������Ĭ��״̬��</td>
          <td  colspan="2" align="left" > <select name="fabiaocheck" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiaocheck")="1" then%>selected<%end if%> value="1">�����</option>
              <option <%if rs9("fabiaocheck")="0" then%>selected<%end if%> value="0">δ���</option>
            </select>
            �����<a title=ע���û���������Ĭ��״̬��ϸ˵�� onMouseOver="window.status='ע���û���������Ĭ��״̬��ϸ˵��';return true;" onMouseOut="window.status='';return true;"  href=javascript:alert("�����ָע���û�������º����²���\n����ϵͳ����Ա���������Ա����˾Ϳ�\n��ʾ��������δ���ָע���û��������\n�����²�ֱ����ʾ���������뾭��ϵͳ\n����Ա���������Ա����˺�Ż���ʾ��")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td width="25%" align="right" >ע���û����޸���������£�</td>
          <td  colspan="2" align="left" > <select name="fabiaomod" size="1" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <option <%if rs9("fabiaomod")="1" then%>selected<%end if%> value="1">����</option>
              <option <%if rs9("fabiaomod")="0" then%>selected<%end if%> value="0">�ر�</option>
            </select>
            �����<a title=ע���û����޸������������ϸ˵�� onMouseOver="window.status='ע���û����޸������������ϸ˵��';return true;" onMouseOut="window.status='';return true;" href=javascript:alert("��ע���û��������µ�Ĭ��״̬Ϊδ���\nʱ����ѡ�����ʾ�����޸ģ����޸ĵ�\n���½���Ϊδ���״̬�����ڰ�ȫ�Ŀ���\n�����رձ�ʾ�����޸ģ���ע���û�����\n����Ĭ��״̬Ϊ�����ʱ����ѡ����Ч��")>����</a>���鿴��ϸ˵���� 
          </td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td colspan="3"  width="611"> <div align="center"> 
              <%if rs9("powermana")="1" and request.cookies(Forcast_SN)("ManagePurview")<>"99999" then%>
              <input type="hidden" name="system" value="<%=rs9("system")%>">
              <input type="hidden" name="css" value="<%=rs9("css")%>">
              <input type="hidden" name="init" value="<%=rs9("init")%>">
              <input type="hidden" name="powermana" value="<%=rs9("powermana")%>">
              <input type="hidden" name="linkmana" value="<%=rs9("linkmana")%>">
              <input type="hidden" name="votemana" value="<%=rs9("votemana")%>">
              <%end if
              rs9.close
              set rs9=nothing
              conn.close
              set conn=nothing
              ConnUser.close
              set ConnUser=nothing%>
              <input type="submit" name="Submit" value="�ύ" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
              <input type="reset" name="Submit2" value="����" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
            </div></td>
        </tr>
</form>
</table>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("�Բ��𣬸ù����Ѿ�������ϵͳ����Ա�رգ���û��Ȩ�޲�����")
		response.end
	end if
end if%>