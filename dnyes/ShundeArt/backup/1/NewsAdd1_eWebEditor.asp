<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="chkuser.asp"-->
<!--#include file = "./eWebEditor/Include/DeCode.asp"-->

<%dim jingyong
smallclassid=request("smallclassid")
bigclassid=request("bigclassid")
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
jingyong=rs("jingyong")
rs.close
set rs=nothing

set rs8=server.createobject("adodb.recordset")
sql8="select * from "& db_System_Table &" where name"
rs8.open sql8,conn,1,3
name=rs8("name")

randomize
rannum=int(90000*rnd)+10000
Response.cookies(Forcast_SN)("newsrelated")=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&rannum

if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("KEY")="bigmaster" or request.cookies(Forcast_SN)("KEY")="smallmaster" or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) or request.cookies(Forcast_SN)("KEY")="typemaster" then%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - �������</title>
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.title.value=="")
	{
	  alert("�Բ������������±��⣡")
	  document.form1.title.focus()
	  return false
	 }
}
//-->
</script>
<script language="JavaScript">
function title_color_onchange() {
	var title_colorIndex;
	title_colorIndex = form1.title_color.selectedIndex;
	form1.title.select ();
	switch (title_colorIndex)
	{
		case 1: form1.title.style.color = "#000000";
				break;
		case 2:	form1.title.style.color = "#000000";
				break;
		case 3:	form1.title.style.color = "#FFFFFF";
				break;
		case 4:	form1.title.style.color = "#008000";
				break;
		case 5:	form1.title.style.color = "#800000";
				break;
		case 6:	form1.title.style.color = "#808000";
				break;
		case 7:	form1.title.style.color = "#000080";
				break;
		case 8:	form1.title.style.color = "#800080";
				break;
		case 9:	form1.title.style.color = "#808080";
				break;
		case 10:form1.title.style.color = "#FFFF00";
				break;
		case 11:form1.title.style.color = "#00FF00";
				break;
		case 12:form1.title.style.color = "#00FFFF";
				break;
		case 13:form1.title.style.color = "#FF00FF";
				break;
		case 14:form1.title.style.color = "#FF0000";
				break;
		case 15:form1.title.style.color = "#0000FF";
				break;
		case 16:form1.title.style.color = "#008080";
				break;
	}
	form1.title.focus ();
}

function title_size_onchange() {
	var title_sizeIndex;
	title_sizeIndex = form1.title_size.selectedIndex;
	form1.title.select ();
	switch (title_sizeIndex)
	{
		case 1: form1.title.style.fontSize = "9pt";
				break;
	}
	form1.title.focus ();
}

function title_type_onchange() {
	var title_typeIndex;
	title_typeIndex = form1.title_type.selectedIndex;
	form1.title.select ();
	switch (title_typeIndex)
	{
		case 1: form1.title.style.removeAttribute ("fontWeight");
				form1.title.style.removeAttribute ("fontStyle");
				form1.title.style.removeAttribute ("textDecoration");
				break;
		case 2:	form1.title.style.removeAttribute ("fontStyle");
				form1.title.style.removeAttribute ("textDecoration");
				form1.title.style.fontWeight = "bolder";
				break;
		case 3:	form1.title.style.removeAttribute ("fontWeight");
				form1.title.style.removeAttribute ("textDecoration");
				form1.title.style.fontStyle = "italic";
				break;
		case 4:	form1.title.style.removeAttribute ("fontStyle");
				form1.title.style.removeAttribute ("fontWeight");
				form1.title.style.textDecoration = "underline";
				break;
		case 5: form1.title.style.removeAttribute ("textDecoration");
				form1.title.style.fontWeight = "bolder";
				form1.title.style.fontStyle = "italic";
				break;
	}
	form1.title.focus ();
}

function title_face_onchange() {
	var title_faceIndex;
	title_faceIndex = form1.title_face.selectedIndex;
	form1.title.select ();
	switch (title_faceIndex)
	{
		case 1: form1.title.style.family = "����";
				break;
		case 2:	form1.title.style.family = "����_GB2312";
				break;
		case 3:	form1.title.style.family = "������";
				break;
		case 4:	form1.title.style.family = "����";
				break;
		case 5:	form1.title.style.family = "����";
				break;
		case 6:	form1.title.style.family = "��Բ";
				break;
	}
	form1.title.focus ();
}
</script>
<link rel="stylesheet" type="text/css" href="site.css">
<link rel="STYLESHEET" type="text/css" href="editor.css">
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%set rs5=server.CreateObject ("ADODB.RecordSet")
rs5.Source="select * from "& db_System_Table &""
rs5.Open rs5.source,conn,1,1%>
<form name="form1" method="GET" action="newsadd2.asp" OnSubmit="return checkdata()" >
  <div align="center"><center>
      <table border="0" cellspacing="0" width="750" cellpadding="0">
        <tr align="center"> 
          <td width="100%"> 
            <table border="1" cellspacing="0" width="100%" cellpadding="0"  bordercolor="<%=border%>" style="border-collapse: collapse" height="463">
              <tr align="center">
<%if smallclassid<>"" then
dim typename,bigclassname,smallclassname
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_SmallClass_Table &" where smallclassid="&smallclassid
rs1.open sql1,conn,1,3
smallclassname=rs1("smallclassname")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_BigClass_Table &" where bigclassid="&rs1("bigclassid")
rs2.open sql2,conn,1,3
bigclassname=rs2("bigclassname")
set rs3=server.createobject("adodb.recordset")
sql3="select * from "& db_Type_Table &" where typeid="&rs2("typeid")
rs3.open sql3,conn,1,3
typename=rs3("typename")
%>
<input name="typeid" type="hidden" value="<%=rs2("typeid")%>">
<input name="bigclassid" type="hidden" value="<%=rs1("bigclassid")%>">
<input name="smallclassid" type="hidden" value="<%=smallclassid%>">
<td colspan="2" height="29" class="unnamed2" valign="middle" bgcolor="<%=m_top%>" width="760">��<b class="unnamed2"><a href="BigClassManage.asp?typeid=<%=rs2("typeid")%>"><%=typename%></a>-<a href="SmallClassManage.asp?bigclassid=<%=rs1("bigclassid")%>"><%=bigclassname%></a>-<a href="listnews.asp?smallclassid=<%=smallclassid%>"><%=smallclassname%></a></b>���������</td>
<%rs3.close
set rs3=nothing
rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
else%>
<%if bigclassid="" then
response.write "<script>alert('������ѡ����࣡');history.go(-1);</Script>"
response.end
else
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_BigClass_Table &" where bigclassid="&bigclassid
rs1.open sql1,conn,1,3
bigclassname=rs1("bigclassname")
set rs2=server.createobject("adodb.recordset")
sql2="select * from "& db_Type_Table &" where typeid="&rs1("typeid")
rs2.open sql2,conn,1,3
typename=rs2("typename")
%>
<input name="typeid" type="hidden" value="<%=rs1("typeid")%>">
<input name="bigclassid" type="hidden" value="<%=bigclassid%>">
                <td colspan="2"  height="29" class="unnamed2" valign="middle" bgcolor="<%=m_top%>" width="760">��<b class="unnamed2"><a href="BigClassManage.asp?typeid=<%=rs2("typeid")%>"><%=typename%></a>-<a href="SmallClassManage.asp?bigclassid=<%=rs1("bigclassid")%>"><%=bigclassname%></a></b>���� �� �� ��</td>
<%rs2.close
set rs2=nothing
rs1.close
set rs1=nothing
end if
end if
%>
              </tr>
              <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >����ר�⣺</td>
                <td width="680" bgcolor="#FFFFFF" height="24"> &nbsp; 
                  <select name="SpecialID" size="1" onMouseOver="window.status='�������ѡ������ר��';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
                    <option value=0>�������κ�ר��</option>
                    <%if rs3.EOF then %>
                    <option value=0>�����κ�ר��</option>
                    <%else
while not rs3.EOF
%>
                    <option value="<%=rs3("SpecialID")%>"><%=trim(rs3("SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
                  </select>
&nbsp;�ڶ�ר�⣺<select name="SpecialID2" size="1" onMouseOver="window.status='�������ѡ������ר��';return true;" onMouseOut="window.status='';return true;">
                    <%			
set rs3=server.CreateObject ("ADODB.RecordSet")
rs3.Source="select * from "& db_Special_Table &""
rs3.Open rs3.source,conn,1,1			
%>
                    <option value=0>�������κ�ר��</option>
                    <%if rs3.EOF then %>
                    <option value=0>�����κ�ר��</option>
                    <%else
while not rs3.EOF
%>
                    <option value="<%=rs3("SpecialID")%>"><%=trim(rs3("SpecialName"))%></option>
                    <%
rs3.MoveNext
wend
end if
rs3.close	
%>
                  </select>
                </td>
              </tr>
              <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">���±��⣺</td>
                <td width="680" bgcolor="#FFFFFF"> <span style='cursor:hand' title='���̶Ի���' onClick='if (me.size>10)me.size=me.size-2'>-</span> 
                  <input name="title" id=me type="TEXT" size=60 maxlength=100 style="background-color:ffffff;border: 1 double"  onMouseOver="window.status='����������Ҫ��ӵ����±��⣬����';return true;" onMouseOut="window.status='';return true;" title='����������Ҫ��ӵ����±��⣬����'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (me.size<102)me.size=me.size+2'>+</span> <font color="#FF0000">*</font><SELECT size=1 id=title_color name=title_color LANGUAGE=javascript onchange="return title_color_onchange()" style="width:55">
<OPTION value="0" selected>��ɫ</OPTION>
<OPTION value="0">ȱʡ</OPTION>
<OPTION value="#000000" style="background-color:#000000"></OPTION>
<OPTION value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
<OPTION value="#008000" style="background-color:#008000"></OPTION>
<OPTION value="#800000" style="background-color:#800000"></OPTION>
<OPTION value="#808000" style="background-color:#808000"></OPTION>
<OPTION value="#000080" style="background-color:#000080"></OPTION>
<OPTION value="#800080" style="background-color:#800080"></OPTION>
<OPTION value="#808080" style="background-color:#808080"></OPTION>
<OPTION value="#FFFF00" style="background-color:#FFFF00"></OPTION>
<OPTION value="#00FF00" style="background-color:#00FF00"></OPTION>
<OPTION value="#00FFFF" style="background-color:#00FFFF"></OPTION>
<OPTION value="#FF00FF" style="background-color:#FF00FF"></OPTION>
<OPTION value="#FF0000" style="background-color:#FF0000"></OPTION>
<OPTION value="#0000FF" style="background-color:#0000FF"></OPTION>
<OPTION value="#008080" style="background-color:#008080"></OPTION>
</SELECT>

<SELECT size=1 id=title_type name=title_type LANGUAGE=javascript onchange="return title_type_onchange()" style="width:55">
<OPTION value="0" selected>����</OPTION>
<OPTION value="0">��ͨ</OPTION>
<OPTION value="b">�Ӵ�</OPTION>
<OPTION value="i">��б</OPTION>
<OPTION value="u">�»���</OPTION>
</SELECT>

</SELECT>
<SELECT size=1 id=title_size name=title_size LANGUAGE=javascript onchange="return title_size_onchange()" style="width:55">
<OPTION value="0" selected>����</OPTION>
<OPTION value="1">Ҫ��</OPTION>
</SELECT>

</td>
              </tr>
              <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >�������£�</td>
                <td width="680" bgcolor="#FFFFFF" > <span style='cursor:hand' title='���̶Ի���' onClick='if (title_face.size>10)title_face.size=title_face.size-2'>-</span> 
                  <input name="title_face" id=title_face type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" onMouseOver="window.status='����Ϊֱ����������ʱ������д��ַ';return true;" onMouseOut="window.status='';return true;" title='����Ϊֱ����������ʱ������д��ַ'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (title_face.size<102)title_face.size=title_face.size+2'>+</span> 
                (����Ϊֱ����������ʱ������д��ַ)</td>
              </tr>
			  <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >������Դ��</td>
                <td width="680" bgcolor="#FFFFFF" > <span style='cursor:hand' title='���̶Ի���' onClick='if (message.size>10)message.size=message.size-2'>-</span> 
                  <input name="Original" id=message type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" onMouseOver="window.status='����У�����������������Դ';return true;" onMouseOut="window.status='';return true;" title='����У�����������������Դ'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (message.size<102)message.size=message.size+2'>+</span> 
                (��ҳģ��ʱ������д��ַ)</td>
              </tr>
              <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >�������ߣ�</td>
                <td width="680" bgcolor="#FFFFFF" ><span style='cursor:hand' title='���̶Ի���' onClick='if (mess.size>10)mess.size=mess.size-2'>-</span> 
                  <input value="" name="Author" id=mess type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" onMouseOver="window.status='����У�������������������';return true;" onMouseOut="window.status='';return true;" title='����У�������������������'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onClick='if (mess.size<102)mess.size=mess.size+2'>+</span> 
                </td>
              </tr>
              <tr> 
                <td width="79" align="right" valign="top" class="unnamed2" bgcolor="#FFFFFF" height="145">�������ݣ�<br><font color="#FF0000">*</font></td>
                <td width="680" valign="top" height="145">
<%
Response.Write eWebEditor_DeCode(sContent, "OBJECT, CLASS, STYLE, NAMESPACE")
%>	
<input type="hidden" name="hdnBigfield" value="����������д���ݡ�">
<iframe ID="message" src="./eWebEditor/ewebeditor.asp?id=hdnBigfield&style=s_mini" frameborder="0" scrolling="no" width="100%" HEIGHT="350"></iframe> 
		</td>
              </tr>
              <tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF" >������£�</td>
                <td width="680" bgcolor="#FFFFFF" ><span style='cursor:hand' title='���̶Ի���' onclick='if (ss.size>10)ss.size=ss.size-2'>-</span> 
                  <INPUT NAME="about" id=ss TYPE="TEXT" SIZE=30 maxlength=100 style="background-color:ffffff;border: 1 double" onMouseOver="window.status='����У�����������������µĹؼ��ֻ���������';return true;" onMouseOut="window.status='';return true;" title='����У�����������������µĹؼ��ֻ���������'>
                  <span style='cursor:hand' title='�ӳ��Ի���' onclick='if (ss.size<102)ss.size=ss.size+2'>+</span> 
                  ����ؼ��ֻ��������� </td>
              </tr><tr> 
                <td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">��ҳͼƬ��</td>
                <td width="680" bgcolor="#FFFFFF">&nbsp;&nbsp;<input name="PicUrl" type="text" id="PicUrl" size="30" maxlength="200" style="background-color:ffffff;border: 1 double">
              ��������ҳ��ͼƬ���´���ʾ����ֱ�Ӵ��ϴ�ͼƬ��ѡ��<select name="PicList" id="PicList" onChange="PicUrl.value=this.value;">
                <option selected>��ָ����ҳͼƬ</option>
              </select>
              </tr>
			  <tr> 
                <td width="79" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >����ͼƬ��</td>
                <td width="680" valign="middle" bgcolor="#FFFFFF" >&nbsp;
				<input type="radio" value="1" name="getphoto" onMouseOver="window.status='������������Զ��ͼƬ';return true;" onMouseOut="window.status='';return true;" title='������������Զ��ͼƬ'>
                  �� 
                  <input type="radio" value="0" checked name="getphoto">
                  ��&nbsp;&nbsp;����Զ��ͼƬ&nbsp;&nbsp;(��������վ�ϸ������ݣ����������а�����ͼƬ)</td>
              </tr>

              <%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="typemaster" then%>
              <tr> 
                <td width="79" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >�Ƽ����£�</td>
                <td width="680" valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="goodnews" onMouseOver="window.status='�������������Ƽ�����';return true;" onMouseOut="window.status='';return true;" title='�������������Ƽ�����'>
                  �� 
                  <input type="radio" value="0" checked name="goodnews">
                  ��&nbsp;&nbsp;�����Ƽ�����</td>
              </tr>
<tr> 
                <td width="79" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >���¹̶���</td>
                <td width="680" valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" name="istop" onMouseOver="window.status='�����������ɹ̶�����';return true;" onMouseOut="window.status='';return true;" title='�����������ɹ̶�����'>
                  �� 
                  <input type="radio" value="0" checked name="istop">
                  ��&nbsp;&nbsp;���ɹ̶�����</td>
              </tr><%if uselevel=1 then%>
<tr> 
                <td width="79" align="right" bgcolor="#FFFFFF" height="30">���¼���</td>
                <td width="680" bgcolor="#FFFFFF" height="30"> &nbsp; <select size="1" name="newslevel" onMouseOver="window.status='������ѡ����Ҫ��ӵ����·ּ��������';return true;" onMouseOut="window.status='';return true;">
<%for i=0 to 3%>
                    <option value="<%=i%>"><%=i%></option>
<%next%>
                    </select>���¼���ָ����������»�Ա�ļ���0Ϊ�οͣ�1Ϊ��ͨ��Ա��2Ϊ�߼���Ա��3Ϊ�ؼ���Ա��
                  </td>
              </tr>
              <%end if%><%end if%>
              <tr  bgcolor="<%=m_top%>" align="center"> 
                <td colspan="2" height="30" width="760">
<input name="newsrelated" type="hidden" value="<%=Request.cookies(Forcast_SN)("newsrelated")%>">
<input name="checkked" type="hidden" value="<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="bigmaster" or request.cookies(Forcast_SN)("key")="typemaster" then%>1<%else%><%if request.cookies(Forcast_SN)("key")="selfreg" and fabiaocheck="1" then%>1<%else%><%if request.cookies(Forcast_SN)("key")="smallmaster" and rs5("usecheck")="1" then%>1<%else%>0<%end if%><%end if%><%end if%>"><%rs5.close%>
<input name="encode" type="hidden" value="html">
<input name="editor" type="hidden" value="<%=request.cookies(Forcast_SN)("username")%>">
<input type="button" value=" �� �� " onclick="javascript:history.go(-1)" class="unnamed5" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">&nbsp;&nbsp;
<input type="submit" value=" �� �� " name="Submit" class="unnamed5" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">
                  &nbsp; 
                  <input type="reset" value=" �� �� "
 name="Reset" class="unnamed5" style="border-style: solid; border-width: 1; padding-left: 4; padding-right: 4; padding-top: 1; padding-bottom: 1">


 <%rs8.close
set rs8=nothing%>

                </td>
              </tr>
            </table>
            
          </td>
        </tr>
      </table>
  </center>
  </div>
</form>
</body>
</html><%end if%>