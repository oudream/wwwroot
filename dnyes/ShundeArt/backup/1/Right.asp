<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<%dim jingyong
set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where  "& db_User_Name &"='"&Request.cookies(Forcast_SN)("username")&"'"
rs.open sql,ConnUser,1,3
if rs.bof or rs.eof then
	response.redirect "login.asp"
	response.end
end if
jingyong=rs("jingyong")
rs.close
set rs=nothing

if Request.cookies(Forcast_SN)("key")="super" or Request.cookies(Forcast_SN)("key")="check" or Request.cookies(Forcast_SN)("KEY")="bigmaster" or Request.cookies(Forcast_SN)("KEY")="smallmaster" or Request.cookies(Forcast_SN)("KEY")="typemaster" or (Request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0) then
%>
<HTML><HEAD><TITLE><%=copyright%><%=version%><%=ver%> - ������ҳ</TITLE>
<META http-equiv=Content-Language content=zh-cn>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="site.css" rel=stylesheet>
<META content="MSHTML 6.00.3790.118" name=GENERATOR>
</HEAD>
<SCRIPT src="inc/menu.js" type=text/javascript></SCRIPT>
<SCRIPT src="inc/js.js" type=text/javascript></SCRIPT>
<SCRIPT src="inc/exit.js" type=text/javascript></SCRIPT>
<SCRIPT language=JavaScript>
<!--
/*for ie and ns*/
document.onclick=function(evt){
var evt=evt?evt:(window.event)?window.event:""
var e=evt.target?evt.target:evt.srcElement
evt.cancelBubble=true
if(e.tagName=="A"&&evt.shiftKey)return false
}
//-->
</SCRIPT>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<TBODY>
<TR class="TDtop" height=25>
<TD height="25" ><div align="center"> �� <B>ϵͳ����</B> �� </div></TD>
</TR>
<TR height=25>
<TD width="100%" height="84" ><p align="left">ϵͳ���� ��<%=copyright%> <br>
<br>
��Ȩ���� �� <a href="http://feitium.yeah.net/">���ڹ����� </a> AND <a href="http://www.5share.com/">��Ե�ž�ͼ��ϵͳ </a><br>
<br>
  �������� �� </p>
<p>��Ҫ����Ա��aaas�������� ����<br>
  ��Ҫ��ȫ���ʣ�TUPUNCO��ˮ�� <br>
  ��Ҫ������ƣ�base <br>
  ��Ҫ����Ա��СͿ��ʦ <br>
  ��������Ա�޸�Ա��base��TUPUNCO <br>
  ��Ҫ���������hjsxah <br>
<br>
  [����3AS���˳�Ե����ϵͳ]�����ɳ�Ե�ž�ͼ��ϵͳ�޸Ķ��ɣ��ر��л��Ե�ž�ͼ��ϵͳ���������ף�jjgn�� <br>
  ���⸽��ϵͳ������ 1.1.0 ������ϵͳ</p>
<p>�����ر��л��li3m��r-rocky ��10h10s��no1��wst��С�ꡢ��ɽ�ɶɡ����ݡ�Сٻ��A09��eline��windy2000��ray��wangy3��tldswgh��roc123��gswei��1.1.1.1.1.1.��hjsxah��ˮ���������ѵĴ���֧�� <br>
<br>
  �ٷ���ʾ��ַ: <a href="http://feitium.yeah.net/">http://feitium.yeah.net <br>
  </a>�ٷ���̳��ַ �� <a href="http://feitiumbbs.yeah.net/">http://feitiumbbs.yeah.net/ </a><br>
  <br>
  �����ƻ����� <br>
  Ϊ�˷�������ϵͳ�ܹ����õķ�չ��������������Ѳ�������޸ģ������Գ����޸���΢С�Ľ�����baseʵ�ڹ��ⲻȥ��������޸Ĺ����У��ö����Ѹ����˴����İ��������ž�������������ģ� </p>
	  <p>1�������ƻ�������Ը�����ҷ����޸ĵ�����ϵͳ���֣���������Ե����ϵͳ��Ѱ桢��ҵ�棩Ҳ������ѣ������Ժ����Ƴ���SQL��Ⱥ���������������ϵͳ�� </p>
	  <p>2��������(���������Ա)������������Ա���ر��л��Ա����̳�����Ա���������������²��԰汾���Լ������ڲ�������ģ�棬������԰����޸�TOPͼ(��Ȼ������񲻱�)�� </p>
	  <p>3�����������70%������Ҫ����������Ա��15%�����ر��л��Ա��һ������С��Ʒ��ʽ����15%�����������õȷ��á� </p>
	  <p>4���������Ϊ10~30Ԫ�� </p>
	  <p>5�������û�������ͬ�ȵļ���֧�֣�ȷ�е�˵�Ǵ�ҹ�ͬ̽��)�� </p>
	  <p>�������ڹ��������⻹���Ƴ����Ʒ��񣨳������ģ�棩��������г����� <br>
          <br>
  ������������ �ٷ���̳ �� <a href="http://feitiumbbs.yeah.net/">http://feitiumbbs.yeah.net </a></p>
	  <p>��Ե�ž�ͼ��ϵͳ��Ȩ�� <br>
          <br>
  �����ס�ѩ��һ���죨jjgn�� <br>
  <a href="http://www.5share.com/">http://www.5share.com/ </a><br>
  E-Mail:jjgn@etang.com <br>
  QQ��82597231 <br>
  2003��1��23�� </p>
  </TD>
	</TR>
</TBODY>
</TABLE>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("�Բ��𣬹���Ա��δ��ͨ�����ʺţ����Ժ�")
	</script>
<%end if%>
<!--#include file=Admin_Bottom.asp-->