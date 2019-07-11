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
<HTML><HEAD><TITLE><%=copyright%><%=version%><%=ver%> - 管理首页</TITLE>
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
<TD height="25" ><div align="center"> ┊ <B>系统介绍</B> ┊ </div></TD>
</TR>
<TR height=25>
<TD width="100%" height="84" ><p align="left">系统名称 ：<%=copyright%> <br>
<br>
版权所有 ： <a href="http://feitium.yeah.net/">沸腾工作室 </a> AND <a href="http://www.5share.com/">尘缘雅境图文系统 </a><br>
<br>
  程序制作 ： </p>
<p>主要程序员：aaas、流浪者 、大奔<br>
  主要安全顾问：TUPUNCO、水晶 <br>
  主要美工设计：base <br>
  主要测试员：小涂老师 <br>
  其他程序员修改员：base、TUPUNCO <br>
  主要插件制作：hjsxah <br>
<br>
  [沸腾3AS流浪尘缘新闻系统]程序由尘缘雅境图文系统修改而成，特别感谢尘缘雅境图文系统作者杨正炎（jjgn） <br>
  另外附属系统：忠网 1.1.0 广告管理系统</p>
<p>另外特别感谢：li3m、r-rocky 、10h10s、no1、wst、小雨、关山飞渡、泡泡、小倩、A09、eline、windy2000、ray、wangy3、tldswgh、roc123、gswei、1.1.1.1.1.1.、hjsxah、水晶、等网友的大力支持 <br>
<br>
  官方演示网址: <a href="http://feitium.yeah.net/">http://feitium.yeah.net <br>
  </a>官方论坛地址 ： <a href="http://feitiumbbs.yeah.net/">http://feitiumbbs.yeah.net/ </a><br>
  <br>
  捐助计划方案 <br>
  为了沸腾新闻系统能够更好的发展，鼓励更多的网友参与程序修改，并给以程序修改者微小的奖励（base实在过意不去，在这个修改工作中，好多网友给以了大量的帮助），才决定出这个方案的： </p>
	  <p>1、捐助计划绝对自愿，而且沸腾修改的新闻系统部分（不包含尘缘新闻系统免费版、商业版）也绝对免费，包括以后考虑推出的SQL版等和自主开发的新闻系统。 </p>
	  <p>2、捐助者(沸腾阳光会员)、程序制作成员、特别感谢成员、论坛贵宾成员，可以提早获得最新测试版本，以及几套内部发布的模版，另外可以帮助修改TOP图(当然基本风格不变)。 </p>
	  <p>3、捐款款项公开，70%发放主要程序制作成员，15%发放特别感谢成员（一般是以小礼品方式），15%用于域名购置等费用。 </p>
	  <p>4、捐款数定为10~30元。 </p>
	  <p>5、所有用户都享受同等的技术支持（确切的说是大家共同探讨)。 </p>
	  <p>附：沸腾工作室另外还将推出定制服务（程序或者模版），这个是有偿服务。 <br>
          <br>
  捐助方法见： 官方论坛 ： <a href="http://feitiumbbs.yeah.net/">http://feitiumbbs.yeah.net </a></p>
	  <p>尘缘雅境图文系统版权： <br>
          <br>
  杨正炎、雪域一线天（jjgn） <br>
  <a href="http://www.5share.com/">http://www.5share.com/ </a><br>
  E-Mail:jjgn@etang.com <br>
  QQ：82597231 <br>
  2003年1月23日 </p>
  </TD>
	</TR>
</TBODY>
</TABLE>
</BODY>
</HTML>
<%else%>
	<script language=javascript>
		history.back()
		alert("对不起，管理员尚未开通您的帐号，请稍候！")
	</script>
<%end if%>
<!--#include file=Admin_Bottom.asp-->