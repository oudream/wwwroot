<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="Char.inc"-->
<%
reviewid=request("reviewid")

if reviewid="" then
	Show_Err("未指定参数<br><br><a href='javascript:history.back()'>返回</a>")
	response.end
else
	if not IsNumeric(reviewid) then
		Show_Err("非法参数<br><br><a href='javascript:history.back()'>返回</a>")
 		response.end
	else 
		if request.cookies(Forcast_SN)("key")="super" then
			aaas=1
			set urs=server.createobject("adodb.recordset")
			sql="select * from "& db_User_Table &" where "& db_User_Name &"='"&Request.cookies(Forcast_SN)("username")&"'"
			urs.open sql,ConnUser,1,3
			if urs.bof or urs.eof then
				aaas=0
			else
				if Request.cookies(Forcast_SN)("passwd")<>urs(db_User_Password) then
					aaas=0
				end if
			end if
			urs.close
			set urs=nothing
		
			set rs=server.createobject("adodb.recordset")
			sql="select * from "& db_Review_Table &" where reviewid="&reviewid 
			rs.open sql,conn,1,3
			if not rs.eof then
				if request.cookies(Forcast_SN)("username")=rs("author") or aaas=1 then
					dim qqhflag
					qqhflag=rs("NewsID")
					%>
<html>
<head>
<title>编辑留言__<%=jjgn%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="news.css" rel=stylesheet type=text/css>
<script language="JavaScript">
<!--
function na_select_form (fname, type_name) 
{
  document.forms[fname].elements[type_name].select()
  document.forms[fname].elements[type_name].focus()
}
// -->
</script>
<script language="javascript">
<!--
function checkdata()
{
if (document.form1.guestname.value=="")
	{
	  alert("对不起，请输入您的大名！")
	  document.form1.guestname.focus()
	  return false
	 }
if (document.form1.content.value=="")
	{
	  alert("对不起，请输入留言内容！")
	  return false
	 }
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<script  language="JavaScript">
<!-- Hide from older browsers...

//Function to format text in the text box
function FormatText(command, option){
	
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}

//Function to clear form
function ResetForm(){

	if (window.confirm('确认要清空对话框内容?')){
	 	frames.message.document.body.innerHTML = ''; 
	 	return true;
	 } 
	 return false;		
}

//Function to open pop up window
function openWin(theURL,winName,features) {
  	window.open(theURL,winName,features);
}

function setMode(newMode)
{
  bTextMode = newMode;
  var cont;
  if (bTextMode) {
    cleanHtml();
    cleanHtml();

    cont=message.document.body.innerHTML;
    message.document.body.innerText=cont;
  } else {
    cont=message.document.body.innerText;
    message.document.body.innerHTML=cont;
  }
  message.focus();
}

function cleanHtml()
{
  var fonts = message.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }

}

// -->
</script>
<script language="JavaScript">
function about()
{
window.showModalDialog("about.htm","","dialogWidth:420px;dialogHeight:240px;scroll:no;status:no;help:no");
}

function help()
{
    var helpmess;
    helpmess="---------------填写帮助---------------\r\n\r\n"+
         "1.请不要发表有危险性的脚本。\r\n\r\n"+
         "2.如果要书写源代码，请选中\r\n\r\n"+
         "　查看HTML源代码书写.\r\n\r\n"+
         "3.需要你自己运行,才能看效果.\r\n\r\n"+
         "4.如果书写js，尽量不要在这儿书写.\r\n\r\n";
    alert(helpmess);
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<form action="guestsave.asp?reviewid=<%=request("reviewid")%>"  method=post name=form1 OnSubmit="return checkdata()" onReset="return ResetForm();">
<tr valign="top"> 
	<td>
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr> 
			<td valign="top">
				<table width="102%" border="0" cellspacing="0" cellpadding="0">
				<tr bgcolor="CDCCCC">
					<td height="3" bgcolor="#FFFFFF" ><img src="images/kb.gif" width="9" height="3"></td>
				</tr>
				<tr bgcolor="CDCCCC"> 
					<td width="100%" height="25" background="images/menu-guestbook.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;栏目导航&nbsp;&nbsp;<a class=daohang href="./" >网站首页</a>&gt;&gt;<a href="GuestBook.asp" class="daohang">雁过留声</a>&gt;&gt;<strong>编辑留言</strong></td>
				</tr>
				<tr bgcolor="CDCCCC"> 
					<td height="25" background="images/menu-guest-l.gif" bgcolor="#FFFFFF">
						<div align="center">
							<a class=daohang href="./" ></a>
							<a class=daohang href="./" ><script language=javascript src=./zongg/ad.asp?i=13></script></a>
						</div>
					</td>
				</tr>
				<tr bgcolor="CDCCCC"> 
					<td height="25" background="images/menu-l-guest.gif" bgcolor="#FFFFFF"></td>
				</tr>
				</table>
			</td>
		</tr>
		</table>
	</td>
</tr>
<tr valign="top"> 
	<td background="images/menu-guest-l.gif"> 
	<table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
		<tr> 
			<td>
				<div align="center">
				<table cellpadding=3 cellspacing=1 border=0 width=95%>
					<tr> 
						<td width="15%" height="22" align="right"> 用户名：</td>
						<td width="70%" height="22">
							<input type="text" name="guestname" maxlength="10" size="10" style="font-family: 宋体; font-size: 9pt;" value="<%=rs("author")%>"> 
						</td>
					</tr>
					<tr> 
						<td width="15%" height="22" align="right"> 密码：</td>
						<td width="70%" height="22">
							<input type="password" name="password" maxlength="10" size="10" style="font-family: 宋体; font-size: 9pt;">如果您是本站注册会员，请输入您的密码
						</td>
					</tr>
					<tr> 
						<td width="15%" align="right" height="22"> &nbsp;身份：</td>
						<td width="70%" height="22">
							<font color="#FF0000"> 
							<select size="1" name="shengfen" style="font-family: 宋体; font-size: 9pt;">
								<option value="保密" selected>保密</option>
								<option value="兄弟">兄弟</option>
								<option value="姐妹">姐妹</option>
							</select>
							</font>
						</td>
					</tr>
					<tr> 
						<td width="15%" height="22" align="right"> &nbsp;来自：</td>
						<td width="70%" height="22">
							<input type="text" name="from" maxlength="10" size="10" style="font-family: 宋体; font-size: 9pt;" value="<%=rs("from")%>"> 
						</td>
					</tr>
					<tr> 
						<td width="15%" height="22" align="right"> &nbsp;e-mail地址：</td>
						<td width="70%" height="22"> <input type="text" name="email" style="font-family: 宋体; font-size: 9pt;" size="20" value="<%=rs("email")%>">如果留空，默认为您注册时使用的E-mail地址。</td>
					</tr>
					<tr> 
						<td width="15%" height="22" align="right"> &nbsp;主页地址：</td>
						<td width="70%" height="22">
							<input type="text" name="homepage" style="font-family: 宋体; font-size: 9pt;" size="20" value="<%=rs("homepage")%>"> 
						</td>
					</tr>
					<tr> 
						<td width="15%" height="22" align="right"> &nbsp;oicq：</td>
						<td width="70%" height="22">
							<input type="text" name="oicq" maxlength="10" style="font-family: 宋体; font-size: 9pt;" size="20" value="<%=rs("oicq")%>"> 
						</td>
					</tr>
					<tr> 
						<td width="15%" align="right" height="22" valign="top"> &nbsp;<br>留言内容：</td>
						<td width="70%" height="22"> <br>
							<select name="selectFont" onChange="FormatText('fontname', selectFont.options[selectFont.selectedIndex].value);document.form1.selectFont.options[0].selected = true;"  style="font-family: 宋体; font-size: 9pt" onMouseOver="window.status='选择选定文字的字体。';return true;" onMouseOut="window.status='';return true;">
								<option selected>选择字体</option>
								<option value="宋体">宋体</option>
								<option value="楷体_GB2312">楷体</option>
								<option value="新宋体">新宋体</option>
								<option value="黑体">黑体</option>
								<option value="隶书">隶书</option>
								<option value="幼圆">幼圆</option>
								<OPTION value="Andale Mono">Andale Mono</OPTION>
								<OPTION value=Arial>Arial</OPTION>
								<OPTION value="Arial Black">Arial Black</OPTION>
								<OPTION value="Book Antiqua">Book Antiqua</OPTION>
								<OPTION value="Century Gothic">Century Gothic</OPTION>
								<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
								<OPTION value="Courier New">Courier New</OPTION>
								<OPTION value=Georgia>Georgia</OPTION>
								<OPTION value=Impact>Impact</OPTION>
								<OPTION value=Tahoma>Tahoma</OPTION>
								<OPTION value="Times New Roman" >Times New Roman</OPTION>
								<OPTION value="Trebuchet MS">Trebuchet MS</OPTION>
								<OPTION value="Script MT Bold">Script MT Bold</OPTION>
								<OPTION value=Stencil>Stencil</OPTION>
								<OPTION value=Verdana>Verdana</OPTION>
								<OPTION value="Lucida Console">Lucida Console</OPTION>
							</select>
							<select name="selectColour" onChange="FormatText('ForeColor',selectColour.options[selectColour.selectedIndex].value);document.form1.selectColour.options[0].selected = true;" onMouseOver="window.status='选择选定文字的颜色。';return true;" onMouseOut="window.status='';return true;">
								<option value="0" selected>选择文字颜色</option>
								<option style="background-color:#F0F8FF;color: #F0F8FF" value="#F0F8FF">#F0F8FF</option>
								<option style="background-color:#FAEBD7;color: #FAEBD7" value="#FAEBD7">#FAEBD7</option>
								<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
								<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">#7FFFD4</option>
								<option style="background-color:#F0FFFF;color: #F0FFFF" value="#F0FFFF">#F0FFFF</option>
								<option style="background-color:#F5F5DC;color: #F5F5DC" value="#F5F5DC">#F5F5DC</option>
								<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">#FFE4C4</option>
								<option style="background-color:#000000;color: #000000" value="#000000">#000000</option>
								<option style="background-color:#FFEBCD;color: #FFEBCD" value="#FFEBCD">#FFEBCD</option>
								<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">#0000FF</option>
								<option style="background-color:#8A2BE2;color: #8A2BE2" value="#8A2BE2">#8A2BE2</option>
								<option style="background-color:#A52A2A;color: #A52A2A" value="#A52A2A">#A52A2A</option>
								<option style="background-color:#DEB887;color: #DEB887" value="#DEB887">#DEB887</option>
								<option style="background-color:#5F9EA0;color: #5F9EA0" value="#5F9EA0">#5F9EA0</option>
								<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>
								<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">#D2691E</option>
								<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">#FF7F50</option>
								<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">#6495ED</option>
								<option style="background-color:#FFF8DC;color: #FFF8DC" value="#FFF8DC">#FFF8DC</option>
								<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
								<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
								<option style="background-color:#00008B;color: #00008B" value="#00008B">#00008B</option>
								<option style="background-color:#008B8B;color: #008B8B" value="#008B8B">#008B8B</option>
								<option style="background-color:#B8860B;color: #B8860B" value="#B8860B">#B8860B</option>
								<option style="background-color:#A9A9A9;color: #A9A9A9" value="#A9A9A9">#A9A9A9</option>
								<option style="background-color:#006400;color: #006400" value="#006400">#006400</option>
								<option style="background-color:#BDB76B;color: #BDB76B" value="#BDB76B">#BDB76B</option>
								<option style="background-color:#8B008B;color: #8B008B" value="#8B008B">#8B008B</option>
								<option style="background-color:#556B2F;color: #556B2F" value="#556B2F">#556B2F</option>
								<option style="background-color:#FF8C00;color: #FF8C00" value="#FF8C00">#FF8C00</option>
								<option style="background-color:#9932CC;color: #9932CC" value="#9932CC">#9932CC</option>
								<option style="background-color:#8B0000;color: #8B0000" value="#8B0000">#8B0000</option>
								<option style="background-color:#E9967A;color: #E9967A" value="#E9967A">#E9967A</option>
								<option style="background-color:#8FBC8F;color: #8FBC8F" value="#8FBC8F">#8FBC8F</option>
								<option style="background-color:#483D8B;color: #483D8B" value="#483D8B">#483D8B</option>
								<option style="background-color:#2F4F4F;color: #2F4F4F" value="#2F4F4F">#2F4F4F</option>
								<option style="background-color:#00CED1;color: #00CED1" value="#00CED1">#00CED1</option>
								<option style="background-color:#9400D3;color: #9400D3" value="#9400D3">#9400D3</option>
								<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">#FF1493</option>
								<option style="background-color:#00BFFF;color: #00BFFF" value="#00BFFF">#00BFFF</option>
								<option style="background-color:#696969;color: #696969" value="#696969">#696969</option>
								<option style="background-color:#1E90FF;color: #1E90FF" value="#1E90FF">#1E90FF</option>
								<option style="background-color:#B22222;color: #B22222" value="#B22222">#B22222</option>
								<option style="background-color:#FFFAF0;color: #FFFAF0" value="#FFFAF0">#FFFAF0</option>
								<option style="background-color:#228B22;color: #228B22" value="#228B22">#228B22</option>
								<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
								<option style="background-color:#DCDCDC;color: #DCDCDC" value="#DCDCDC">#DCDCDC</option>
								<option style="background-color:#F8F8FF;color: #F8F8FF" value="#F8F8FF">#F8F8FF</option>
								<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">#FFD700</option>
								<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">#DAA520</option>
								<option style="background-color:#808080;color: #808080" value="#808080">#808080</option>
								<option style="background-color:#008000;color: #008000" value="#008000">#008000</option>
								<option style="background-color:#ADFF2F;color: #ADFF2F" value="#ADFF2F">#ADFF2F</option>
								<option style="background-color:#F0FFF0;color: #F0FFF0" value="#F0FFF0">#F0FFF0</option>
								<option style="background-color:#FF69B4;color: #FF69B4" value="#FF69B4">#FF69B4</option>
								<option style="background-color:#CD5C5C;color: #CD5C5C" value="#CD5C5C">#CD5C5C</option>
								<option style="background-color:#4B0082;color: #4B0082" value="#4B0082">#4B0082</option>
								<option style="background-color:#FFFFF0;color: #FFFFF0" value="#FFFFF0">#FFFFF0</option>
								<option style="background-color:#F0E68C;color: #F0E68C" value="#F0E68C">#F0E68C</option>
								<option style="background-color:#E6E6FA;color: #E6E6FA" value="#E6E6FA">#E6E6FA</option>
								<option style="background-color:#FFF0F5;color: #FFF0F5" value="#FFF0F5">#FFF0F5</option>
								<option style="background-color:#7CFC00;color: #7CFC00" value="#7CFC00">#7CFC00</option>
								<option style="background-color:#FFFACD;color: #FFFACD" value="#FFFACD">#FFFACD</option>
								<option style="background-color:#ADD8E6;color: #ADD8E6" value="#ADD8E6">#ADD8E6</option>
								<option style="background-color:#F08080;color: #F08080" value="#F08080">#F08080</option>
								<option style="background-color:#E0FFFF;color: #E0FFFF" value="#E0FFFF">#E0FFFF</option>
								<option style="background-color:#FAFAD2;color: #FAFAD2" value="#FAFAD2">#FAFAD2</option>
								<option style="background-color:#90EE90;color: #90EE90" value="#90EE90">#90EE90</option>
								<option style="background-color:#D3D3D3;color: #D3D3D3" value="#D3D3D3">#D3D3D3</option>
								<option style="background-color:#FFB6C1;color: #FFB6C1" value="#FFB6C1">#FFB6C1</option>
								<option style="background-color:#FFA07A;color: #FFA07A" value="#FFA07A">#FFA07A</option>
								<option style="background-color:#20B2AA;color: #20B2AA" value="#20B2AA">#20B2AA</option>
								<option style="background-color:#87CEFA;color: #87CEFA" value="#87CEFA">#87CEFA</option>
								<option style="background-color:#778899;color: #778899" value="#778899">#778899</option>
								<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">#B0C4DE</option>
								<option style="background-color:#FFFFE0;color: #FFFFE0" value="#FFFFE0">#FFFFE0</option>
								<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">#00FF00</option>
								<option style="background-color:#32CD32;color: #32CD32" value="#32CD32">#32CD32</option>
								<option style="background-color:#FAF0E6;color: #FAF0E6" value="#FAF0E6">#FAF0E6</option>
								<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
								<option style="background-color:#800000;color: #800000" value="#800000">#800000</option>
								<option style="background-color:#66CDAA;color: #66CDAA" value="#66CDAA">#66CDAA</option>
								<option style="background-color:#0000CD;color: #0000CD" value="#0000CD">#0000CD</option>
								<option style="background-color:#BA55D3;color: #BA55D3" value="#BA55D3">#BA55D3</option>
								<option style="background-color:#9370DB;color: #9370DB" value="#9370DB">#9370DB</option>
								<option style="background-color:#3CB371;color: #3CB371" value="#3CB371">#3CB371</option>
								<option style="background-color:#7B68EE;color: #7B68EE" value="#7B68EE">#7B68EE</option>
								<option style="background-color:#00FA9A;color: #00FA9A" value="#00FA9A">#00FA9A</option>
								<option style="background-color:#48D1CC;color: #48D1CC" value="#48D1CC">#48D1CC</option>
								<option style="background-color:#C71585;color: #C71585" value="#C71585">#C71585</option>
								<option style="background-color:#191970;color: #191970" value="#191970">#191970</option>
								<option style="background-color:#F5FFFA;color: #F5FFFA" value="#F5FFFA">#F5FFFA</option>
								<option style="background-color:#FFE4E1;color: #FFE4E1" value="#FFE4E1">#FFE4E1</option>
								<option style="background-color:#FFE4B5;color: #FFE4B5" value="#FFE4B5">#FFE4B5</option>
								<option style="background-color:#FFDEAD;color: #FFDEAD" value="#FFDEAD">#FFDEAD</option>
								<option style="background-color:#000080;color: #000080" value="#000080">#000080</option>
								<option style="background-color:#FDF5E6;color: #FDF5E6" value="#FDF5E6">#FDF5E6</option>
								<option style="background-color:#808000;color: #808000" value="#808000">#808000</option>
								<option style="background-color:#6B8E23;color: #6B8E23" value="#6B8E23">#6B8E23</option>
								<option style="background-color:#FFA500;color: #FFA500" value="#FFA500">#FFA500</option>
								<option style="background-color:#FF4500;color: #FF4500" value="#FF4500">#FF4500</option>
								<option style="background-color:#DA70D6;color: #DA70D6" value="#DA70D6">#DA70D6</option>
								<option style="background-color:#EEE8AA;color: #EEE8AA" value="#EEE8AA">#EEE8AA</option>
								<option style="background-color:#98FB98;color: #98FB98" value="#98FB98">#98FB98</option>
								<option style="background-color:#AFEEEE;color: #AFEEEE" value="#AFEEEE">#AFEEEE</option>
								<option style="background-color:#DB7093;color: #DB7093" value="#DB7093">#DB7093</option>
								<option style="background-color:#FFEFD5;color: #FFEFD5" value="#FFEFD5">#FFEFD5</option>
								<option style="background-color:#FFDAB9;color: #FFDAB9" value="#FFDAB9">#FFDAB9</option>
								<option style="background-color:#CD853F;color: #CD853F" value="#CD853F">#CD853F</option>
								<option style="background-color:#FFC0CB;color: #FFC0CB" value="#FFC0CB">#FFC0CB</option>
								<option style="background-color:#DDA0DD;color: #DDA0DD" value="#DDA0DD">#DDA0DD</option>
								<option style="background-color:#B0E0E6;color: #B0E0E6" value="#B0E0E6">#B0E0E6</option>
								<option style="background-color:#800080;color: #800080" value="#800080">#800080</option>
								<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">#FF0000</option>
								<option style="background-color:#BC8F8F;color: #BC8F8F" value="#BC8F8F">#BC8F8F</option>
								<option style="background-color:#4169E1;color: #4169E1" value="#4169E1">#4169E1</option>
								<option style="background-color:#8B4513;color: #8B4513" value="#8B4513">#8B4513</option>
								<option style="background-color:#FA8072;color: #FA8072" value="#FA8072">#FA8072</option>
								<option style="background-color:#F4A460;color: #F4A460" value="#F4A460">#F4A460</option>
								<option style="background-color:#2E8B57;color: #2E8B57" value="#2E8B57">#2E8B57</option>
								<option style="background-color:#FFF5EE;color: #FFF5EE" value="#FFF5EE">#FFF5EE</option>
								<option style="background-color:#A0522D;color: #A0522D" value="#A0522D">#A0522D</option>
								<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">#C0C0C0</option>
								<option style="background-color:#87CEEB;color: #87CEEB" value="#87CEEB">#87CEEB</option>
								<option style="background-color:#6A5ACD;color: #6A5ACD" value="#6A5ACD">#6A5ACD</option>
								<option style="background-color:#708090;color: #708090" value="#708090">#708090</option>
								<option style="background-color:#FFFAFA;color: #FFFAFA" value="#FFFAFA">#FFFAFA</option>
								<option style="background-color:#00FF7F;color: #00FF7F" value="#00FF7F">#00FF7F</option>
								<option style="background-color:#4682B4;color: #4682B4" value="#4682B4">#4682B4</option>
								<option style="background-color:#D2B48C;color: #D2B48C" value="#D2B48C">#D2B48C</option>
								<option style="background-color:#008080;color: #008080" value="#008080">#008080</option>
								<option style="background-color:#D8BFD8;color: #D8BFD8" value="#D8BFD8">#D8BFD8</option>
								<option style="background-color:#FF6347;color: #FF6347" value="#FF6347">#FF6347</option>
								<option style="background-color:#40E0D0;color: #40E0D0" value="#40E0D0">#40E0D0</option>
								<option style="background-color:#EE82EE;color: #EE82EE" value="#EE82EE">#EE82EE</option>
								<option style="background-color:#F5DEB3;color: #F5DEB3" value="#F5DEB3">#F5DEB3</option>
								<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">#FFFFFF</option>
								<option style="background-color:#F5F5F5;color: #F5F5F5" value="#F5F5F5">#F5F5F5</option>
								<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">#FFFF00</option>
								<option style="background-color:#9ACD32;color: #9ACD32" value="#9ACD32">#9ACD32</option>
							</select>
							<select name="selectbgColour" onChange="FormatText('BackColor',selectbgColour.options[selectbgColour.selectedIndex].value);document.form1.selectbgColour.options[0].selected = true;" onMouseOver="window.status='选择选定文字的背景颜色。';return true;" onMouseOut="window.status='';return true;">
								<option value="0" selected>选择文字背景颜色</option>
								<option style="background-color:#F0F8FF;color: #F0F8FF" value="#F0F8FF">#F0F8FF</option>
								<option style="background-color:#FAEBD7;color: #FAEBD7" value="#FAEBD7">#FAEBD7</option>
								<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
								<option style="background-color:#7FFFD4;color: #7FFFD4" value="#7FFFD4">#7FFFD4</option>
								<option style="background-color:#F0FFFF;color: #F0FFFF" value="#F0FFFF">#F0FFFF</option>
								<option style="background-color:#F5F5DC;color: #F5F5DC" value="#F5F5DC">#F5F5DC</option>
								<option style="background-color:#FFE4C4;color: #FFE4C4" value="#FFE4C4">#FFE4C4</option>
								<option style="background-color:#000000;color: #000000" value="#000000">#000000</option>
								<option style="background-color:#FFEBCD;color: #FFEBCD" value="#FFEBCD">#FFEBCD</option>
								<option style="background-color:#0000FF;color: #0000FF" value="#0000FF">#0000FF</option>
								<option style="background-color:#8A2BE2;color: #8A2BE2" value="#8A2BE2">#8A2BE2</option>
								<option style="background-color:#A52A2A;color: #A52A2A" value="#A52A2A">#A52A2A</option>
								<option style="background-color:#DEB887;color: #DEB887" value="#DEB887">#DEB887</option>
								<option style="background-color:#5F9EA0;color: #5F9EA0" value="#5F9EA0">#5F9EA0</option>
								<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>
								<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">#D2691E</option>
								<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">#FF7F50</option>
								<option style="background-color:#6495ED;color: #6495ED" value="#6495ED">#6495ED</option>
								<option style="background-color:#FFF8DC;color: #FFF8DC" value="#FFF8DC">#FFF8DC</option>
								<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
								<option style="background-color:#00FFFF;color: #00FFFF" value="#00FFFF">#00FFFF</option>
								<option style="background-color:#00008B;color: #00008B" value="#00008B">#00008B</option>
								<option style="background-color:#008B8B;color: #008B8B" value="#008B8B">#008B8B</option>
								<option style="background-color:#B8860B;color: #B8860B" value="#B8860B">#B8860B</option>
								<option style="background-color:#A9A9A9;color: #A9A9A9" value="#A9A9A9">#A9A9A9</option>
								<option style="background-color:#006400;color: #006400" value="#006400">#006400</option>
								<option style="background-color:#BDB76B;color: #BDB76B" value="#BDB76B">#BDB76B</option>
								<option style="background-color:#8B008B;color: #8B008B" value="#8B008B">#8B008B</option>
								<option style="background-color:#556B2F;color: #556B2F" value="#556B2F">#556B2F</option>
								<option style="background-color:#FF8C00;color: #FF8C00" value="#FF8C00">#FF8C00</option>
								<option style="background-color:#9932CC;color: #9932CC" value="#9932CC">#9932CC</option>
								<option style="background-color:#8B0000;color: #8B0000" value="#8B0000">#8B0000</option>
								<option style="background-color:#E9967A;color: #E9967A" value="#E9967A">#E9967A</option>
								<option style="background-color:#8FBC8F;color: #8FBC8F" value="#8FBC8F">#8FBC8F</option>
								<option style="background-color:#483D8B;color: #483D8B" value="#483D8B">#483D8B</option>
								<option style="background-color:#2F4F4F;color: #2F4F4F" value="#2F4F4F">#2F4F4F</option>
								<option style="background-color:#00CED1;color: #00CED1" value="#00CED1">#00CED1</option>
								<option style="background-color:#9400D3;color: #9400D3" value="#9400D3">#9400D3</option>
								<option style="background-color:#FF1493;color: #FF1493" value="#FF1493">#FF1493</option>
								<option style="background-color:#00BFFF;color: #00BFFF" value="#00BFFF">#00BFFF</option>
								<option style="background-color:#696969;color: #696969" value="#696969">#696969</option>
								<option style="background-color:#1E90FF;color: #1E90FF" value="#1E90FF">#1E90FF</option>
								<option style="background-color:#B22222;color: #B22222" value="#B22222">#B22222</option>
								<option style="background-color:#FFFAF0;color: #FFFAF0" value="#FFFAF0">#FFFAF0</option>
								<option style="background-color:#228B22;color: #228B22" value="#228B22">#228B22</option>
								<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
								<option style="background-color:#DCDCDC;color: #DCDCDC" value="#DCDCDC">#DCDCDC</option>
								<option style="background-color:#F8F8FF;color: #F8F8FF" value="#F8F8FF">#F8F8FF</option>
								<option style="background-color:#FFD700;color: #FFD700" value="#FFD700">#FFD700</option>
								<option style="background-color:#DAA520;color: #DAA520" value="#DAA520">#DAA520</option>
								<option style="background-color:#808080;color: #808080" value="#808080">#808080</option>
								<option style="background-color:#008000;color: #008000" value="#008000">#008000</option>
								<option style="background-color:#ADFF2F;color: #ADFF2F" value="#ADFF2F">#ADFF2F</option>
								<option style="background-color:#F0FFF0;color: #F0FFF0" value="#F0FFF0">#F0FFF0</option>
								<option style="background-color:#FF69B4;color: #FF69B4" value="#FF69B4">#FF69B4</option>
								<option style="background-color:#CD5C5C;color: #CD5C5C" value="#CD5C5C">#CD5C5C</option>
								<option style="background-color:#4B0082;color: #4B0082" value="#4B0082">#4B0082</option>
								<option style="background-color:#FFFFF0;color: #FFFFF0" value="#FFFFF0">#FFFFF0</option>
								<option style="background-color:#F0E68C;color: #F0E68C" value="#F0E68C">#F0E68C</option>
								<option style="background-color:#E6E6FA;color: #E6E6FA" value="#E6E6FA">#E6E6FA</option>
								<option style="background-color:#FFF0F5;color: #FFF0F5" value="#FFF0F5">#FFF0F5</option>
								<option style="background-color:#7CFC00;color: #7CFC00" value="#7CFC00">#7CFC00</option>
								<option style="background-color:#FFFACD;color: #FFFACD" value="#FFFACD">#FFFACD</option>
								<option style="background-color:#ADD8E6;color: #ADD8E6" value="#ADD8E6">#ADD8E6</option>
								<option style="background-color:#F08080;color: #F08080" value="#F08080">#F08080</option>
								<option style="background-color:#E0FFFF;color: #E0FFFF" value="#E0FFFF">#E0FFFF</option>
								<option style="background-color:#FAFAD2;color: #FAFAD2" value="#FAFAD2">#FAFAD2</option>
								<option style="background-color:#90EE90;color: #90EE90" value="#90EE90">#90EE90</option>
								<option style="background-color:#D3D3D3;color: #D3D3D3" value="#D3D3D3">#D3D3D3</option>
								<option style="background-color:#FFB6C1;color: #FFB6C1" value="#FFB6C1">#FFB6C1</option>
								<option style="background-color:#FFA07A;color: #FFA07A" value="#FFA07A">#FFA07A</option>
								<option style="background-color:#20B2AA;color: #20B2AA" value="#20B2AA">#20B2AA</option>
								<option style="background-color:#87CEFA;color: #87CEFA" value="#87CEFA">#87CEFA</option>
								<option style="background-color:#778899;color: #778899" value="#778899">#778899</option>
								<option style="background-color:#B0C4DE;color: #B0C4DE" value="#B0C4DE">#B0C4DE</option>
								<option style="background-color:#FFFFE0;color: #FFFFE0" value="#FFFFE0">#FFFFE0</option>
								<option style="background-color:#00FF00;color: #00FF00" value="#00FF00">#00FF00</option>
								<option style="background-color:#32CD32;color: #32CD32" value="#32CD32">#32CD32</option>
								<option style="background-color:#FAF0E6;color: #FAF0E6" value="#FAF0E6">#FAF0E6</option>
								<option style="background-color:#FF00FF;color: #FF00FF" value="#FF00FF">#FF00FF</option>
								<option style="background-color:#800000;color: #800000" value="#800000">#800000</option>
								<option style="background-color:#66CDAA;color: #66CDAA" value="#66CDAA">#66CDAA</option>
								<option style="background-color:#0000CD;color: #0000CD" value="#0000CD">#0000CD</option>
								<option style="background-color:#BA55D3;color: #BA55D3" value="#BA55D3">#BA55D3</option>
								<option style="background-color:#9370DB;color: #9370DB" value="#9370DB">#9370DB</option>
								<option style="background-color:#3CB371;color: #3CB371" value="#3CB371">#3CB371</option>
								<option style="background-color:#7B68EE;color: #7B68EE" value="#7B68EE">#7B68EE</option>
								<option style="background-color:#00FA9A;color: #00FA9A" value="#00FA9A">#00FA9A</option>
								<option style="background-color:#48D1CC;color: #48D1CC" value="#48D1CC">#48D1CC</option>
								<option style="background-color:#C71585;color: #C71585" value="#C71585">#C71585</option>
								<option style="background-color:#191970;color: #191970" value="#191970">#191970</option>
								<option style="background-color:#F5FFFA;color: #F5FFFA" value="#F5FFFA">#F5FFFA</option>
								<option style="background-color:#FFE4E1;color: #FFE4E1" value="#FFE4E1">#FFE4E1</option>
								<option style="background-color:#FFE4B5;color: #FFE4B5" value="#FFE4B5">#FFE4B5</option>
								<option style="background-color:#FFDEAD;color: #FFDEAD" value="#FFDEAD">#FFDEAD</option>
								<option style="background-color:#000080;color: #000080" value="#000080">#000080</option>
								<option style="background-color:#FDF5E6;color: #FDF5E6" value="#FDF5E6">#FDF5E6</option>
								<option style="background-color:#808000;color: #808000" value="#808000">#808000</option>
								<option style="background-color:#6B8E23;color: #6B8E23" value="#6B8E23">#6B8E23</option>
								<option style="background-color:#FFA500;color: #FFA500" value="#FFA500">#FFA500</option>
								<option style="background-color:#FF4500;color: #FF4500" value="#FF4500">#FF4500</option>
								<option style="background-color:#DA70D6;color: #DA70D6" value="#DA70D6">#DA70D6</option>
								<option style="background-color:#EEE8AA;color: #EEE8AA" value="#EEE8AA">#EEE8AA</option>
								<option style="background-color:#98FB98;color: #98FB98" value="#98FB98">#98FB98</option>
								<option style="background-color:#AFEEEE;color: #AFEEEE" value="#AFEEEE">#AFEEEE</option>
								<option style="background-color:#DB7093;color: #DB7093" value="#DB7093">#DB7093</option>
								<option style="background-color:#FFEFD5;color: #FFEFD5" value="#FFEFD5">#FFEFD5</option>
								<option style="background-color:#FFDAB9;color: #FFDAB9" value="#FFDAB9">#FFDAB9</option>
								<option style="background-color:#CD853F;color: #CD853F" value="#CD853F">#CD853F</option>
								<option style="background-color:#FFC0CB;color: #FFC0CB" value="#FFC0CB">#FFC0CB</option>
								<option style="background-color:#DDA0DD;color: #DDA0DD" value="#DDA0DD">#DDA0DD</option>
								<option style="background-color:#B0E0E6;color: #B0E0E6" value="#B0E0E6">#B0E0E6</option>
								<option style="background-color:#800080;color: #800080" value="#800080">#800080</option>
								<option style="background-color:#FF0000;color: #FF0000" value="#FF0000">#FF0000</option>
								<option style="background-color:#BC8F8F;color: #BC8F8F" value="#BC8F8F">#BC8F8F</option>
								<option style="background-color:#4169E1;color: #4169E1" value="#4169E1">#4169E1</option>
								<option style="background-color:#8B4513;color: #8B4513" value="#8B4513">#8B4513</option>
								<option style="background-color:#FA8072;color: #FA8072" value="#FA8072">#FA8072</option>
								<option style="background-color:#F4A460;color: #F4A460" value="#F4A460">#F4A460</option>
								<option style="background-color:#2E8B57;color: #2E8B57" value="#2E8B57">#2E8B57</option>
								<option style="background-color:#FFF5EE;color: #FFF5EE" value="#FFF5EE">#FFF5EE</option>
								<option style="background-color:#A0522D;color: #A0522D" value="#A0522D">#A0522D</option>
								<option style="background-color:#C0C0C0;color: #C0C0C0" value="#C0C0C0">#C0C0C0</option>
								<option style="background-color:#87CEEB;color: #87CEEB" value="#87CEEB">#87CEEB</option>
								<option style="background-color:#6A5ACD;color: #6A5ACD" value="#6A5ACD">#6A5ACD</option>
								<option style="background-color:#708090;color: #708090" value="#708090">#708090</option>
								<option style="background-color:#FFFAFA;color: #FFFAFA" value="#FFFAFA">#FFFAFA</option>
								<option style="background-color:#00FF7F;color: #00FF7F" value="#00FF7F">#00FF7F</option>
								<option style="background-color:#4682B4;color: #4682B4" value="#4682B4">#4682B4</option>
								<option style="background-color:#D2B48C;color: #D2B48C" value="#D2B48C">#D2B48C</option>
								<option style="background-color:#008080;color: #008080" value="#008080">#008080</option>
								<option style="background-color:#D8BFD8;color: #D8BFD8" value="#D8BFD8">#D8BFD8</option>
								<option style="background-color:#FF6347;color: #FF6347" value="#FF6347">#FF6347</option>
								<option style="background-color:#40E0D0;color: #40E0D0" value="#40E0D0">#40E0D0</option>
								<option style="background-color:#EE82EE;color: #EE82EE" value="#EE82EE">#EE82EE</option>
								<option style="background-color:#F5DEB3;color: #F5DEB3" value="#F5DEB3">#F5DEB3</option>
								<option style="background-color:#FFFFFF;color: #FFFFFF" value="#FFFFFF">#FFFFFF</option>
								<option style="background-color:#F5F5F5;color: #F5F5F5" value="#F5F5F5">#F5F5F5</option>
								<option style="background-color:#FFFF00;color: #FFFF00" value="#FFFF00">#FFFF00</option>
								<option style="background-color:#9ACD32;color: #9ACD32" value="#9ACD32">#9ACD32</option>
							</select>
							<select language="javascript"  id="select" title="字号大小" onChange="FormatText('fontsize',this[this.selectedIndex].value);" name="select" onMouseOver="window.status='选择选定文字的字号大小。';return true;" onMouseOut="window.status='';return true;">
								<option class="heading" selected>字号 
								<option value="7">一号 
								<option value="6">二号 
								<option value="5">三号 
								<option value="4">四号 
								<option value="3">五号 
								<option value="2">六号 
								<option value="1">七号</option>
							</select><br>
							<img class=None src="images/selectall.gif" align="absmiddle" border="0" alt="全部选择" onClick="FormatText('selectall')" style="cursor: hand;"> 
							<img class=None src="images/cut.gif"  align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="剪切"> 
							<img class=None src="images/copy.gif"  align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="复制"> 
							<img class=None src="images/paste.gif"  align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="粘贴"> 
							<img class=None src="images/DELETE.gif" align="absmiddle" border="0" alt="删除" onClick="FormatText('DELETE')" style="cursor: hand;"> 
							<img class=None src="images/undo.gif" align="absmiddle" border="0" alt="撤消" onClick="FormatText('undo')" style="cursor: hand;"> 
							<img class=None src="images/redo.gif" align="absmiddle" border="0" alt="恢复" onClick="FormatText('redo')" style="cursor: hand;"> 
							<img class=None src="images/bold.gif" align="absmiddle" alt="粗体" onClick="FormatText('bold', '')" style="cursor: hand;"> 
							<img class=None src="images/italic.gif" align="absmiddle" alt="斜体" onClick="FormatText('italic', '')" style="cursor: hand;"> 
							<img class=None src="images/underline.gif" align="absmiddle" alt="下划线" onClick="FormatText('underline', '')" style="cursor: hand;"> 
							<img class=None src="images/ALEFT.gif" align="absmiddle" onClick="FormatText('Justifyleft', '')" style="cursor: hand;" alt="左对齐"> 
							<img class=None src="images/center.gif" align="absmiddle" border="0" alt="居中" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"> 
							<img class=None src="images/aright.gif" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="右对齐"> 
							<img class=None src="images/list.gif" align="absmiddle" border="0" alt="项目符号" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"> 
							<img class=None src="images/number.gif" align="absmiddle" alt="编号" border="0" onClick="FormatText('insertorderedlist', '')" style="cursor: hand;"> 
							<img class=None src="images/outdent.gif" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="回退"> 
							<img class=None src="images/indent.gif" align="absmiddle" border="0" alt="缩进" onClick="FormatText('indent', '')" style="cursor: hand;"> 
							<img class=None src="images/line.gif" align="absmiddle" alt="普通水平线" border="0" onClick="FormatText('InsertHorizontalRule', '')"  style="cursor: hand;"> 
							<img class=None src="images/link.gif" align="absmiddle" border="0" alt="超级链接" onClick="FormatText('createLink')" style="cursor: hand;"> 
							<img class=None src="images/unlink.gif" align="absmiddle" border="0" alt="取消超级链接" onClick="FormatText('unLink')" style="cursor: hand;"> 
							<img class=None src="images/sup.gif" align="absmiddle" border="0" alt="上标" onClick="FormatText('superscript')" style="cursor: hand;"> 
							<img class=None src="images/sub.gif" align="absmiddle" border="0" alt="下标" onClick="FormatText('subscript')" style="cursor: hand;"> 
							<img class=None src="images/clear.gif" align="absmiddle" border="0" alt="删除文字格式" onClick="FormatText('RemoveFormat')" style="cursor: hand;"> 
							<br><br>
							<script language="javascript">
								document.write ('<iframe src="guestbox.asp?action=modify&reviewid=<%=reviewid%>" id="message" width="500" height="200" align=left></iframe>')
								frames.message.document.designMode = "On";
							</script>
						</td>
				</tr>
				</table>
				</div>
			</td>
		</tr>
		</table>
	</td>
</tr>
</table>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="efefef">
<tr> 
	<td height="25" align="center" background="images/menu-guest-l.gif">以上内容仅写给管理员（悄悄话）：&nbsp;&nbsp;
		<%if qqhflag=-1 then%>
		<input type="radio" name="qqh" checked value=-1>是&nbsp;&nbsp;
		<input type="radio" name="qqh" value=0>否
		<%else%>
		<input type="radio" name="qqh" value=-1>是&nbsp;&nbsp;
		<input type="radio" name="qqh" checked value=0>否
		<%end if%>
		<br><br>
	</td>
</tr>
<tr> 
	<td height="25" align="center" background="images/menu-guest-l.gif"> <input name="button" type="button" onclick="javascript:history.go(-1)" value=" 返 回 "> 
		&nbsp;&nbsp; <input type="submit" value=" 修 改 " name="Submit" OnClick="document.form1.content.value = frames.message.document.body.innerHTML;"> 
		&nbsp; <input type="reset" value=" 清 除 " name="Reset"> <input type="hidden" name="content" value=""> 
	</td>
</tr>
<tr bgcolor="759BAC"> 
	<td height="19" align="center" background="images/menu-guest-t.gif">　</td>
</tr>
</form>
</table>
</body>
</html>
<!--#include file="bottom.asp"-->
				<%else
					Show_Err("您不能编辑别人的留言！<br><br><a href='javascript:history.back()'>返回</a>")
					response.end
				end if
			end if
		else
				Show_Err("您无权编辑此留言！<br><br><a href='javascript:history.back()'>返回</a>")
		 		response.end
		end if
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
