<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/email.asp"-->
<!--#include file="md5.asp"-->
<%
'=========================================================
' File: reg.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
	response.buffer=true
	stats="�û�ע��"
	call nav()
	call head_var(1)
	if request("action")="apply" then
		call reg_2()
	elseif request("action")="save" then
		call reg_3()
		if founderr then call dvbbs_error()
	else
		call reg_1()
	end if
	call activeonline()
	call footer()

sub reg_1()
%>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr><th align=center><form action="reg.asp?action=apply" method=post>�������������</td></tr>
    <tr><td class=tablebody1 align=left>
<b>����ע��ǰ�����Ķ�<%=Forum_info(0)%>��̳Э��</b>
<BR><BR>
��ӭ������<%=Forum_info(0)%>�μӽ��������ۣ�<%=Forum_info(0)%>Ϊ������̳��Ϊά�����Ϲ������������ȶ��������Ծ������������
<BR><BR>
һ���������ñ�վΣ�����Ұ�ȫ��й¶�������ܣ������ַ�������Ἧ��ĺ͹���ĺϷ�Ȩ�棬�������ñ�վ���������ƺʹ���������Ϣ�� 
<BR><BR>
��һ��ɿ�����ܡ��ƻ��ܷ��ͷ��ɡ���������ʵʩ�ģ�<BR>
������ɿ���߸�������Ȩ���Ʒ���������ƶȵģ�<BR>
������ɿ�����ѹ��ҡ��ƻ�����ͳһ�ģ�<BR>
���ģ�ɿ�������ޡ��������ӣ��ƻ������Ž�ģ�<BR>
���壩�������������ʵ��ɢ��ҥ�ԣ������������ģ�<BR>
����������⽨���š����ࡢɫ�顢�Ĳ�����������ɱ���ֲ�����������ģ�<BR>
���ߣ���Ȼ�������˻���������ʵ�̰����˵ģ����߽����������⹥���ģ�<BR>
���ˣ��𺦹��һ��������ģ�<BR>
���ţ�����Υ���ܷ��ͷ�����������ģ�<BR>
��ʮ��������ҵ�����Ϊ�ġ�
<BR><BR>
�����������أ����Լ������ۺ���Ϊ���� 
	</td></tr>
    <tr><td align=center class=tablebody2><input type=submit value=��ͬ��></td></form></tr>
</table>
<%
end sub

sub reg_2()
%>
<FORM name=theForm action=reg.asp?action=save method=post>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<TBODY> 
<TR align=middle> 
<Th colSpan=2 height=24>���û�ע��</TD>
</TR>
<TR> 
<TD width=40% class=tablebody1><B>�û���</B>��<BR>ע���û������ܳ���12���ַ���6�����֣�</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=12 size=30 name=name></TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>�Ա�</B>��<BR>��ѡ�������Ա�</font></TD>
<TD width=60%  class=tablebody1> <INPUT type=radio CHECKED value=1 name=sex>
<IMG  src=pic/Male.gif align=absMiddle>�к� &nbsp;&nbsp;&nbsp;&nbsp; 
<INPUT type=radio value=0 name=sex>
<IMG  src=pic/Female.gif align=absMiddle>Ů��</font></TD>
</TR>
<%if cint(Forum_Setting(23))<>1 then%>
<TR> 
<TD width=40% class=tablebody1><B>����(����6λ)</B>��<BR>
���������룬���ִ�Сд��<BR>
�벻Ҫʹ���κ����� '*'��' ' �� HTML �ַ�
</TD>
<TD width=60% class=tablebody1>
<INPUT type=password maxLength=16 size=30 name=psw>
</TD>
</TR>
<TR> 
<TD width=40% class=tablebody1><B>����(����6λ)</B>��<BR>������һ��ȷ��</TD>
<TD class=tablebody1> 
<INPUT type=password maxLength=16 size=30 name=pswc>
</TD>
</TR>
<%end if%>
<TR> 
<TD width=40%  class=tablebody1><B>��������</B>��<BR>�����������ʾ����</TD>
<TD class=tablebody1> 
<INPUT type=text size=30 name=quesion>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>�����</B>��<BR>�����������ʾ����𰸣�����ȡ����̳����</TD>
<TD class=tablebody1> 
<INPUT type=text size=30 name=answer>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>Email��ַ</B>��<BR>��������Ч���ʼ���ַ���⽫ʹ�����õ���̳�е����й���</font></TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=50 size=30 name=e_mail>����<input type=button value='����ʺ�' name=Button onclick=gopreview()></TD>
</TR>
</tbody>
</table>
</td></tr></tbody></table>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1 id=adv style="DISPLAY: none">
<TBODY> 
<TR align=middle> 
<Th colSpan=2 height=24 align=left>��д��ϸ����</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>ͷ��</B>��<BR>ѡ���ͷ�񽫳������������Ϻͷ���������У���Ҳ����ѡ��������Զ���ͷ��</TD>
<TD width=60%  class=tablebody1> 
<select name=face size=1 onChange="document.images['face'].src=options[selectedIndex].value;" style="BACKGROUND-COLOR: #99CCFF; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #000000">
<%for i=1 to 60%>
<option value='pic/Image<%=i%>.gif'>Image<%=i%></option>
<%next%>
</select>
&nbsp;<img id=face src=<%=Forum_info(7)%>Image1.gif>&nbsp;<a href=allface.asp target=_blank>�鿴����ͷ��</a>
</TR>
<TR> 
<TD width=40% valign=top class=tablebody1><B>�Զ���ͷ��</B>��<br>���ͼ��λ����������ͼƬ�����Զ����Ϊ��</TD>
<TD width=60%  class=tablebody1>
<%if cint(Forum_Setting(7))=1 then%>
<iframe name=ad frameborder=0 width=300 height=40 scrolling=no src=reg_upload.asp></iframe> 
<br>
<%end if%>
ͼ��λ�ã� 
<input type=TEXT name=myface size=20 maxlength=100>
&nbsp;����Url��ַ<br>
��&nbsp;&nbsp;&nbsp;&nbsp;�ȣ� 
<input type=TEXT name=width size=3 maxlength=3 value=32>
20---120������<br>
��&nbsp;&nbsp;&nbsp;&nbsp;�ȣ� 
<input type=TEXT name=height size=3 maxlength=3 value=32>
20---120������<br>
</TD>
</TR>
      <tr bgcolor="&Forum_body(4)&">    
        <td width=40%  class=tablebody1><B>����</B><BR>�粻����д����ȫ������</td>   
        <td width=60%  class=tablebody1>
<select name=birthyear>
<option value=></option>
<%for i=1901 to 2002%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
                   �� 
<select name=birthmonth>
<option value=></option>
<%for i=1 to 12%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
                   �� 
<select name=birthday>
<option value=></option>
<%for i=1 to 31%>
<option value="<%=i%>"><%=i%></option>
<%next%>
</select>
��
        </td>   
      </tr>
<tr> 
<td width=40%  class=tablebody1><B>�ظ���ʾ</B>��<BR>����������������˻ظ�ʱ��ʹ����̳��Ϣ֪ͨ����</td>
<td width=60%  class=tablebody1>
<input type=radio name=showRe value=1 checked>
��ʾ��
<input type=radio name=showRe value=0>
����ʾ
</tr>
<%if cint(Forum_Setting(32))=1 then%>
<TR> 
<TD width=40% class=tablebody1><B>����</B>��<BR>����������ѡ��Ҫ���������</TD>
<TD width=60% class=tablebody1> 
<select name=groupname>
<%
set rs=conn.execute("select * from GroupName")
if rs.eof and rs.bof then
	response.write "<option value=��������>��������</option>"
else
	do while not rs.eof
	response.write "<option value="&rs("Groupname")&">"&rs("GroupName")&"</option>"
	rs.movenext
	loop
end if
rs.close
set rs=nothing
%>
</select>
</TD>
</TR>
<%end if%>
<TR> 
<TD width=40%  class=tablebody1><B>OICQ����</B>��<BR>��д����QQ��ַ�����������˵���ϵ</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=20 size=44 name=OICQ>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>ICQ����</B>��<BR>��д����ICQ��ַ�����������˵���ϵ</font></TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=20 size=44 name=ICQ>
</TD>
</TR>
<TR > 
<TD width=40%  class=tablebody1><B>MSN</B>��<BR>��д����MSN��ַ�����������˵���ϵ</TD>
<TD width=60%  class=tablebody1> 
<INPUT maxLength=70 size=44 name=msn>
</TD>
</TR>
<TR> 
<TD width=40%  class=tablebody1><B>ǩ��</B>��<BR>���300�ֽ�<BR>
���ֽ�����������������µĽ�β�����������ĸ��ԡ� </TD>
<TD width=60%  class=tablebody1> 
<TEXTAREA name=Signature rows=5 wrap=PHYSICAL cols=60></TEXTAREA>
</TD>
</TR>
<tr>    
<td width=40%  class=tablebody1><B>ѡ��Cookie�ı���ʱ��</B>��<BR>��½��̳��Ϣ����ʱ�䣬�����ʱ�����ظ���½��̳����Ҫ���µ�½</font></td>  
<td width=60%  class=tablebody1>    
			              <input type=radio name=usercookies value=1 checked>
              <font color=red>1��</font> 
              <input type=radio name=usercookies value=2>
    1����
    <input type=radio name=usercookies value=3>
    1��
              <input type=radio name=usercookies value=0>
    ������ </td>   
      </tr>
<tr>    
<td width=40%  class=tablebody1><B>�Ƿ񿪷����Ļ�������</B>��<BR>���ź���˿��Կ��������Ա�Email��QQ����Ϣ</td>  
<td width=60%  class=tablebody1>    
			              <input type=radio name=setuserinfo value=1 checked>
              ����              <input type=radio name=setuserinfo value=0>
    ������ </td>   
      </tr>
<tr>    
<td width=40%  class=tablebody1><B>�Ƿ񿪷�������ʵ����</B>��<BR>���ź���˿��Կ���������ʵ��������ϵ��ʽ����Ϣ</td>
  <td width=60%  class=tablebody1>    
			
              <input type=radio name=setusertrue value=1>
              ����              <input type=radio name=setusertrue value=0 checked>
    ������</td>   
      </tr>
<tr>
<th height=25 align=left valign=middle colspan=2><b>&nbsp;������ʵ��Ϣ</b>���������ݽ�����д��</th>
</tr>
<tr>
<td height=2 valign=top  class=tablebody1 width=40% > ��<b>��ʵ������</b>
<input type=text name=realname size=18>
</td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>�ԡ��� &nbsp; </b>
<br>
<input type=checkbox value= �������� name=character> �������ء�
<input type=checkbox value= ���ɵ�Ƥ name=character> ���ɵ�Ƥ��
<input type=checkbox value= �������� name=character> ����������
<input type=checkbox value= ���ÿɰ� name=character> ���ÿɰ���
<input type=checkbox value= Ƣ������ name=character> Ƣ������<br>

<input type=checkbox value= ������ name=character> �����ߡ�
<input type=checkbox value= ������ name=character> �����ʡ�
<input type=checkbox value= �ĵ����� name=character> �ĵ�������
<input type=checkbox value= ��Ȥ��Ĭ name=character> ��Ȥ��Ĭ��
<input type=checkbox value= ����˹�� name=character> ����˹��<br>

<input type=checkbox value= ������ȡ name=character> ������ȡ��
<input type=checkbox value= �����ѻ� name=character> �����ѻ���
<input type=checkbox value= �������� name=character> �������ѡ�
<input type=checkbox value= Բ������ name=character> Բ��������
<input type=checkbox value= ���Ų�� name=character> ���Ų��<br>

<input type=checkbox value= �����쿪 name=character> �����쿪��
<input type=checkbox value= ����Ƹ� name=character> ����ƸС�
<input type=checkbox value= �������� name=character> ����������
<input type=checkbox value= ������� name=character> ������䡡
<input type=checkbox value= ��С���� name=character> ��С����<br>

<input type=checkbox value= ѭ�浸�� name=character> ѭ�浸�ء�
<input type=checkbox value= �������� name=character> �������ˡ�
<input type=checkbox value= ���Կ��� name=character> ���Կ��
<input type=checkbox value= �������� name=character> �������¡�
<input type=checkbox value= ׷��̼� name=character> ׷��̼�</td>
</tr>
<tr>
<td class=tablebody1><b>���˼�飺 &nbsp;
<textarea name=personal rows=6 cols=65></textarea>
</b></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>�������ң�</b>
<b>
<input type=text name=country size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>��ϵ�绰��</b>
<b>
<input type=text name=userphone size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>ͨ�ŵ�ַ��</b>
<b>
<input type=text name=address size=18>
</b> </td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>ʡ�����ݣ�</b>
<input type=text name=province size=18>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>�ǡ����У�
</b>
<input type=text name=city size=18>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>������Ф��
</b>
<select size=1 name=shengxiao>
<option></option>
<option value=��>��</option>
<option value=ţ>ţ</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
<option value=��>��</option>
</select>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>Ѫ�����ͣ�</b>
<select size=1 name=blood>
<option selected></option>
<option value=A>A</option>
<option value=B>B</option>
<option value=AB>AB</option>
<option value=O>O</option>
<option value=����>����</option>
</select>
</td>
</tr>
<tr>
<td valign=top  class=tablebody1 width=40% >��<b>�š�������</b>
<select size=1 name=belief>
<option selected></option>
<option value=���>���</option>
<option value=����>����</option>
<option value=������>������</option>
<option value=������>������</option>
<option value=�ؽ�>�ؽ�</option>
<option value=��������>��������</option>
<option value=����������>����������</option>
<option value=����>����</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >��<b>ְ����ҵ�� </b>
<select name=occupation>
<option selected> </option>
<option value=�ƻ�/����>�ƻ�/����</option>
<option value=����ʦ>����ʦ</option>
<option value=����>����</option>
<option value=����������ҵ>����������ҵ</option>
<option value=��ͥ����>��ͥ����</option>
<option value=����/��ѵ>����/��ѵ</option>
<option value=�ͻ�����/֧��>�ͻ�����/֧��</option>
<option value=������/�ֹ�����>������/�ֹ�����</option>
<option value=����>����</option>
<option value=��ְҵ>��ְҵ</option>
<option value=����/�г�/���>����/�г�/���</option>
<option value=ѧ��>ѧ��</option>
<option value=�о��Ϳ���>�о��Ϳ���</option>
<option value=һ�����/�ල>һ�����/�ල</option>
<option value=����/����>����/����</option>
<option value=ִ�й�/�߼�����>ִ�й�/�߼�����</option>
<option value=����/����/����>����/����/����</option>
<option value=רҵ��Ա>רҵ��Ա</option>
<option value=�Թ�/ҵ��>�Թ�/ҵ��</option>
<option value=����>����</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >��<b>����״����</b>
<select size=1 name=marital>
<option selected></option>
<option value=δ��>δ��</option>
<option value=�ѻ�>�ѻ�</option>
<option value=����>����</option>
<option value=ɥż>ɥż</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >��<b>���ѧ����</b>
<select size=1 name=education>
<option selected></option>
<option value=Сѧ>Сѧ</option>
<option value=����>����</option>
<option value=����>����</option>
<option value=��ѧ>��ѧ</option>
<option value=˶ʿ>˶ʿ</option>
<option value=��ʿ>��ʿ</option>
</select></td>
</tr>
<tr>
<td valign=top class=tablebody1 width=40% >��<b>��ҵԺУ��</b>
<input type=text name=college size=18></td>
</tr>
</TBODY> 
</TABLE>
</td></tr></table>
<table cellpadding=0 cellspacing=0 border=0 width=<%=Forum_body(12)%> align=center>
<tr>
<td width=50% height=24><INPUT id=advcheck name=advshow type=checkbox value=1 onclick=showadv()><span id=advance>��ʾ�߼��û�����ѡ��</a></span> </td>
<td width=50% ><input type=submit value="ע ��" name=Submit>����<input type=reset value="�� ��" name=Submit2></td>
</tr></table>
</form>
<form name=preview action=chkreg.asp method=post target=preview_page>
<input type=hidden name=username value=><input type=hidden name=email value=>
</form>
<script>
function gopreview()
{
document.forms[1].username.value=document.forms[0].name.value;
document.forms[1].email.value=document.forms[0].e_mail.value;
var popupWin = window.open('preview.asp', 'preview_page', 'scrollbars=yes,width=500,height=300');
document.forms[1].submit()
}
function showadv(){
if (document.theForm.advshow.checked == true) {
	adv.style.display = "";
	advance.innerText="�رո߼��û�����ѡ��"
}else{
	adv.style.display = "none";
	advance.innerText="��ʾ�߼��û�����ѡ��"
}
}
</script>
<%
end sub

sub reg_3()
dim username,sex,pass1,pass2,password
dim useremail,face,width,height
dim oicq,sign,showRe,birthday
dim mailbody,sendmsg,rndnum,num1
dim quesion,answer,topic
dim userinfo,usersetting
dim userclass
if not isnull(session("regtime")) or cint(Forum_Setting(22))>0 then
	if DateDiff("s",session("regtime"),Now())<cint(Forum_Setting(22)) then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����̳����ÿ��ע�����ʱ��Ϊ"&Forum_Setting(22)&"�룬���Ժ�ע�ᡣ"
   		FoundErr=True
	end if
end if

if chkpost=false then
	ErrMsg=ErrMsg+"<Br>"+"<li>���ύ�����ݲ��Ϸ����벻Ҫ���ⲿ�ύ���ԡ�"
   	FoundErr=True
end if
if request("name")="" or strLength(request("name"))>20 then
	errmsg=errmsg+"<br>"+"<li>�����������û���(���Ȳ��ܴ���20)��"
	founderr=true
else
	username=trim(request("name"))
end if
if Instr(request("name"),"=")>0 or Instr(request("name"),"%")>0 or Instr(request("name"),chr(32))>0 or Instr(request("name"),"?")>0 or Instr(request("name"),"&")>0 or Instr(request("name"),";")>0 or Instr(request("name"),",")>0 or Instr(request("name"),"'")>0 or Instr(request("name"),",")>0 or Instr(request("name"),chr(34))>0 or Instr(request("name"),chr(9))>0 or Instr(request("name"),"��")>0 or Instr(request("name"),"$")>0 then
	errmsg=errmsg+"<br>"+"<li>�û����к��зǷ��ַ���"
	founderr=true
else
	username=trim(request("name"))
end if

splitwords=split(splitwords,",")
for i = 0 to ubound(splitwords)
	if instr(username,splitwords(i))>0 then
		errmsg=errmsg+"<br>"+"<li>��������û�������ϵͳ��ֹע���ַ���"
		founderr=true
		exit for
	end if
next

if request("sex")=0 or request("sex")=1 then
	sex=request("sex")
else
	sex=1
end if
	
if request("showRe")=0 or request("showRe")=1 then
	showRe=request("showRe")
else
	showRe=1
end if	
if cint(Forum_Setting(23))=1 then
	Randomize
	Do While Len(rndnum)<8
	num1=CStr(Chr((57-48)*rnd+48))
	rndnum=rndnum&num1
	loop
	password=md5(rndnum)
else
	if request("psw")="" or len(request("psw"))>10 or len(request("psw"))<6 then
		errmsg=errmsg+"<br>"+"<li>��������������(���Ȳ��ܴ���10С��6)��"
		founderr=true
	else
	pass1=request("psw")
	end if
	if request("pswc")="" or strLength(request("pswc"))>10 or len(request("pswc"))<6 then
		errmsg=errmsg+"<br>"+"<li>������ȷ������(���Ȳ��ܴ���10С��6)��"
		founderr=true
	else
		pass2=request("pswc")
	end if
	if pass1<>pass2 then
		errmsg=errmsg+"<br>"+"<li>������������ȷ�����벻һ�¡�"
		founderr=true
	else
		password=md5(pass2)
	end if
end if

if request("quesion")="" then
  	errmsg=errmsg+"<br>"+"<li>������������ʾ���⡣"
	founderr=true
else
	quesion=request("quesion")
end if
if request("answer")="" then
  	errmsg=errmsg+"<br>"+"<li>������������ʾ����𰸡�"
	founderr=true
elseif request("answer")=request("oldanswer") then
	answer=request("answer")
else
	answer=md5(request("answer"))
end if

if IsValidEmail(trim(request("e_mail")))=false then
	errmsg=errmsg+"<br>"+"<li>����Email�д���"
   	founderr=true
else
	useremail=checkStr((request("e_mail")))
end if
if request.form("myface")<>"" then
	if request("width")="" or request("height")="" then
		errmsg=errmsg+"<br>"+"<li>������ͼƬ�Ŀ�Ⱥ͸߶ȡ�"
		founderr=true
	elseif not isInteger(request("width")) or not isInteger(request("height")) then
		errmsg=errmsg+"<br>"+"<li>��������ַ����Ϸ���"
		founderr=true
	elseif request("width")<20 or request("width")>120 then
		errmsg=errmsg+"<br>"+"<li>�������ͼƬ��Ȳ����ϱ�׼��"
		founderr=true
	elseif request("height")<20 or request("height")>120 then
		errmsg=errmsg+"<br>"+"<li>�������ͼƬ�߶Ȳ����ϱ�׼��"
		founderr=true
	else
		face=request("myface")
		width=request("width")
		height=request("height")
	end if
else
	if request("face")<>"" then
		if Instr(request("face"),Forum_info(7))>0 then
		face=request("face")
		width=32
		height=32
		end if
	end if
end if
if request("oicq")<>"" then
	if not isnumeric(request("oicq")) or len(request("oicq"))>10 then
		errmsg=errmsg+"<br>"+"<li>Oicq����ֻ����4-10λ���֣�������ѡ�����롣"
		founderr=true
	end if
end if
if request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" then
	birthday=""
else
	birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
	if not isdate(birthday) then birthday=""
end if
userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")
if founderr then exit sub
dim titlepic
set rs=conn.execute("select usertitle,titlepic from usertitle where not minarticle=-1 order by minarticle")
userclass=rs(0)
titlepic=rs(1)
set rs=server.createobject("adodb.recordset")
sql="select * from [user] where username='"&username&"'"
rs.open sql,conn,1,3
if not rs.eof and not rs.bof then
	if cint(Forum_Setting(24))=1 and rs("useremail")=useremail then
		errmsg=errmsg+"<br>"+"<li>�Բ��𣬱���̳�Ѿ�������һ��Emailֻ��ע��һ���ʺţ�������ѡ������Email��"
		founderr=true
		exit sub
	else
		errmsg=errmsg+"<br>"+"<li>�Բ�����������û����Ѿ���ע�ᡣ"
		founderr=true
		exit sub
	end if
else
	rs.addnew
	rs("username")=username
	rs("userpassword")=password
	rs("useremail")=useremail
	rs("userclass")=userclass
	rs("titlepic")=titlepic
	rs("quesion")=quesion
	rs("answer")=answer
	
	if request("Signature")<>"" then
		rs("sign")=trim(request("Signature"))
	end if
	if request("oicq")<>"" then
		rs("oicq")=request("oicq")
	end if
	if request("icq")<>"" then
		rs("icq")=request("icq")
	end if
	if request("msn")<>"" then
		rs("msn")=request("msn")
	end if
	Rs("article")=0
	if cint(Forum_Setting(25))=1 then
		rs("usergroupid")=5
	else
       	Rs("usergroupid")=4
	end if
	rs("lockuser")=0
	Rs("sex")=sex
	Rs("showRe")=showRe
	if birthday<>"" then
		rs("birthday")=birthday
	end if
	rs("UserGroup")=request("GroupName")
	Rs("addDate")=NOW()
	if face<>"" then
		rs("face")=face
       	Rs("width")=width
       	Rs("height")=height
	end if
	rs("logins")=1
	Rs("lastlogin")=NOW()
	rs("userWealth")=Forum_user(0)
	rs("userEP")=Forum_user(5)
	rs("usercP")=Forum_user(10)
	rs("userinfo")=userinfo
	rs("usersetting")=usersetting
	rs("bbstype")=tempid
	rs.update
	conn.execute("update config set usernum=usernum+1,lastuser='"&username&"'")
end if
rs.close
set rs=nothing
set rs=conn.execute("select top 1 userid from [user] order by userid desc")
userid=rs(0)
if Forum_Setting(47)=1 then
	on error resume next
	'����ע���ʼ�
	dim getpass
	topic="����" & Forum_info(0) & "��ע������"
	if cint(Forum_Setting(23))=1 then
		getpass=htmlencode(rndnum)
	else
		getpass=htmlencode(request("psw"))
	end if
	mailbody="<html>"
	mailbody=mailbody & "<title>ע����Ϣ</title>"
	mailbody=mailbody & "<body>"
	mailbody=mailbody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
	mailbody=mailbody & "<TD valign=middle align=top>"
	mailbody=mailbody & htmlencode(username)&"�����ã�<br><br>"
	mailbody=mailbody & "��ӭ��ע�᱾��̳�����ǽ��ṩ������õ���̳����<br>"
	mailbody=mailbody & "����������ע����Ϣ��<br>"
	mailbody=mailbody & "ע������"&htmlencode(username)&"<br>"
	mailbody=mailbody & "��  �룺"&getpass&"<br>"
	mailbody=mailbody & "<br><br>"
	mailbody=mailbody & "<center><font color=red>�ٴθ�л��ע�᱾ϵͳ��������һ��������������ϼ�԰��</font>"
	mailbody=mailbody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	mailbody=mailbody & "<p align=center>" & Copyright & " &nbsp;&nbsp; " & Version & "</p>"
	mailbody=mailbody & "</body>"
	mailbody=mailbody & "</html>"
	select case Forum_Setting(2)
	case 0
		sendmsg="<li>ϵͳδ�����ʼ����ܣ����ס����ע����Ϣ��</li>"
	case 1
		call jmail(useremail,topic,mailbody)
	case 2
		call Cdonts(useremail,topic,mailbody)
	case 3
		call aspemail(useremail,topic,mailbody)
	case else
		sendmsg="<li>ϵͳδ�����ʼ����ܣ����ס����ע����Ϣ��</li>"
	end select
	if SendMail="OK" then
		if cint(Forum_Setting(23))=1 then
		sendmsg="<li>����ע����Ϣ�������Ѿ������������䣬��ʹ��ϵͳ�����������½��</li>"
		else
		sendmsg="<li>����ע����Ϣ�Ѿ������������䣬��ע����ա�</li>"
		end if
	else
		sendmsg="<li>����ϵͳ���󣬸������͵�ע������δ�ɹ���</li>"
	end if
end if

if Forum_Setting(46)=1 then
	'����ע�����
	dim sender,title,body
	sender=Forum_info(0)
	title=Forum_info(0)&"��ӭ���ĵ���"
	body=Forum_info(0)&"ȫ�������Ա��ӭ���ĵ���"&chr(10)&"�����κ������뼰ʱ��ϵϵͳ����Ա��"&chr(10)&"�����κ�ʹ���ϵ�������鿴��̳������"&chr(10)&"��л��ע�᱾ϵͳ��������һ��������������ϼ�԰��"
	sql="insert into message(incept,sender,title,content,sendtime,flag,issend) values('"&username&"','"&sender&"','"&title&"','"&body&"',Now(),0,1)"
	conn.Execute(sql)
end if

if cint(Forum_Setting(23))=1 or cint(Forum_Setting(25))=1 then

else
	if founduser then
		Response.Cookies("aspsky").path=cookiepath
		Response.Cookies("aspsky")("username")=""
		Response.Cookies("aspsky")("password")=""
		Response.Cookies("aspsky")("userclass")=""
		Response.Cookies("aspsky")("userid")=""
		Response.Cookies("aspsky")("userhidden")=""
		Response.Cookies("aspsky")("usercookies")=""
		conn.execute("delete from online where username='"&membername&"'")
	end if
	select case request("usercookies")
 	case 0
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 1
	    Response.Cookies("aspsky").Expires=Date+1
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 2
		Response.Cookies("aspsky").Expires=Date+31
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
    case 3
	    Response.Cookies("aspsky").Expires=Date+365
 		Response.Cookies("aspsky")("usercookies") = request("usercookies")
	end select
	Response.Cookies("aspsky")("username") = username
	Response.Cookies("aspsky")("password") = password
	Response.Cookies("aspsky")("userclass") = userclass
	Response.Cookies("aspsky")("userid") = userid
	Response.Cookies("aspsky")("userhidden") = 2
	Response.Cookies("aspsky").path=cookiepath
end if
session("regtime")=now()
%>
<meta HTTP-EQUIV=REFRESH CONTENT='2; URL=index.asp'>
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
<tr>
<th height=24>ע��ɹ���<%=Forum_info(0)%>��ӭ���ĵ���</th>
</tr>
<tr><td class=tablebody1><br>
<ul><%=sendmsg%><li><a href="index.asp">����������</a></li></ul>
</td></tr>
</table>
<%
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","����")
	checkreal=w
end if
end function
%>