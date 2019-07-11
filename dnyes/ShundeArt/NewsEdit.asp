<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<%

if not(request.cookies(Forcast_SN)("KEY")="super" or request.cookies(Forcast_SN)("KEY")="typemaster" or request.cookies(Forcast_SN)("KEY")="bigmaster" or request.cookies(Forcast_SN)("key")="selfreg") then
	Show_Err("对不起，您的后台管理权限不够操作！")
	response.end
else
	dim jingyong
	set rs=server.createobject("adodb.recordset")
	sql="select * from "& db_User_Table &" where username='"&request.cookies(Forcast_SN)("username")&"'"
	rs.open sql,ConnUser,1,3
	jingyong=rs("jingyong")
	rs.close
	set rs=nothing
	
	set rs8=server.createobject("adodb.recordset")
	sql8="select * from "& db_System_Table &" where name"
	rs8.open sql8,conn,1,3
	name=rs8("name")
	
	NewsID=request("NewsID")
	if NewsID="" then
		Show_Err("对不起，要修改的文章ID不能为空！")
		response.end
	end if
	
	set rs=server.CreateObject ("ADODB.RecordSet")
	rs.Source="select * from "& db_News_Table &" where NewsID=" & NewsID
	rs.Open rs.source,conn,3,1
	if rs.EOF then
		Show_Err("对不起，文章ID超出范围！！")
		response.end
	end if
	newsrelated=rs("newsrelated")
	typeid=rs("typeid")
	bigclassid=rs("bigclassid")
	smallclassid=rs("smallclassid")
	if smallclassid<>"" then
	set rs3=server.createobject("adodb.recordset")
	sql3="select * from "& db_SmallClass_Table &" where smallclassid="&smallclassid
	rs3.open sql3,conn,1,3
	master6=rs3("smallclassma")
	if master6<>"" then
	smallmaster6=split(master6,"|")
	 for i=0 to ubound(smallmaster6)
		if request.cookies(Forcast_SN)("username")=trim(smallmaster6(i)) then
		smallclassmaster6=true
		exit for
		else
		smallclassmaster6=false
		end if
	 next
	end if
	 rs3.close
	set rs3=nothing
	end if
	set rs1=server.createobject("adodb.recordset")
	sql1="select * from "& db_BigClass_Table &" where bigclassid="&rs("bigclassid")
	rs1.open sql1,conn,1,3
	master5=rs1("bigclassmaster")
	if master5<>"" then
	bigmaster5=split(master5,"|")
	 for i=0 to ubound(bigmaster5)
		if request.cookies(Forcast_SN)("username")=trim(bigmaster5(i)) then
		bigclassmaster5=true
		exit for
		else
		bigclassmaster5=false
		end if
	 next
	end if
	set rs2=server.createobject("adodb.recordset")
	sql2="select * from "& db_Type_Table &" where typeid="&typeid
	rs2.open sql2,conn,1,3
	master4=rs2("typemaster")
	if master4<>"" then
	tmaster4=split(master4,"|")
	 for i=0 to ubound(tmaster4)
		if request.cookies(Forcast_SN)("username")=trim(tmaster4(i)) then
		typemaster4=true
		exit for
		else
		typemaster4=false
		end if
	 next
	end if
	rs2.close
	set rs2=nothing
	rs1.close
	set rs1=nothing
	set rs2=server.CreateObject("ADODB.RecordSet")
	rs2.Source="select * From "& db_BigClass_Table &" Where BigClassid=" & rs("BigClassid")
	rs2.Open rs2.Source,conn,1,1
	if rs("smallclassid")<>"" then
	set rs3=server.CreateObject("ADODB.RecordSet")
	rs3.Source="select * From "& db_SmallClass_Table &" Where SmallClassid=" & rs("SmallClassid")
	rs3.Open rs3.Source,conn,1,1
	end if
	if rs("checkked")="0" or request.cookies(Forcast_SN)("key")="super" or (request.cookies(Forcast_SN)("key")="bigmaster" and bigclassmaster5=true) or (request.cookies(Forcast_SN)("key")="typemaster" and typemaster4=true) or (request.cookies(Forcast_SN)("key")="check" and checkmod="1") or (modnewsable="1" and request.cookies(Forcast_SN)("key")="smallmaster" and smallclassmaster6=true) or (request.cookies(Forcast_SN)("key")="smallmaster" and rs("checkked")="1" and modnewsable="1" and smallclassmaster6=true) or (request.cookies(Forcast_SN)("key")="selfreg" and jingyong=0 and fabiaomod="1") then
	%>
<html><head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title><%=copyright%><%=version%>&nbsp;<%=ver%> - 修改文章</title>
<script language="JavaScript">
<!--
function na_select_form (fname, type_name) 
{
  document.forms[fname].elements[type_name].select()
  document.forms[fname].elements[type_name].focus()
}
function onTable(num) {

	
	
	if(	document.all.frame1.style.visibility == "visible" ) {
		cancelTable();
		return;
	}
	
	var str = "<div id=\"tblsel\" style=\"background-color:#000066;position:absolute;";
	str = str + "width:0;height:0;z-index:-1;\"></div>";
	str = str + makeTable(4, 5);
	str = str + "<div style=\"text-align:center;background-color:menu;font-size: 12px\" id=\"tblstat\">插入表格</div>";
	
	var ifrm = document.frames("frame1");
	var obj=eval("document.all.ae_tbtn"+1);
	var x=0;
	var y=0;

	ifrm.document.body.innerHTML=str;
	

	while(obj.tagName!="BODY") {

		x+=obj.offsetLeft;
		y+=obj.offsetTop;
		obj=obj.offsetParent;

	}	
	
	document.all.frame1.style.pixelTop = y + 24;
	document.all.frame1.style.pixelLeft = x;
	document.all.frame1.style.pixelWidth = 0;
	document.all.frame1.style.pixelHeight = 0;
	document.all.frame1.style.visibility = "visible";
	ae_hot=num;
	document.frames("frame1").document.body.onmouseover = paintTable;	
	document.frames("frame1").document.body.onclick = insertTable;
	if(typeof(document.onmousedown)=="function")
		ae_olddocmd = document.onmousedown;
	else ae_olddoccmd=null;
	
	document.onmousedown = cancelTable;
	DHTMLSafe.onmousedown = cancelTable;
	event.cancelBubble = true;

	ifrm.document.body.onselectstart=new Function("return false;");
	

	document.all.frame1.style.pixelWidth = ifrm.document.all.oTable.offsetWidth + 3
	document.all.frame1.style.pixelHeight = ifrm.document.all.oTable.offsetHeight + 3 +
		ifrm.document.all.tblstat.offsetHeight;

}
function makeTable(rows, cols) {
	var a, b, str, n;
	str = "<table style=\"table-layout:fixed;border-style:solid; cursor:default;\" "; 
	str = str + "id=\"oTable\" cellpadding=\"0\" ";
	str = str + "cellspacing=\"0\" cols=" + cols;
	str = str + " border=6>\n";

	for (a=0;a<rows;a++) {
		str = str + "<tr>\n";
		for(b=0;b<cols;b++) {			
			str = str + "<td width=\"20\">" 
			str = str + "&nbsp;</td>\n";	
		}	
		str = str + "</tr>\n";
	}
	str = str + "</table>"
	return str;
}

function cancelTable() {
	document.onmousedown=null;
	document.all.frame1.style.visibility = "hidden";
	document.all.frame1.style.pixelWidth = 0;
	document.all.frame1.style.pixelHeight = 0;

	if(a==false) return;

	if(typeof(ae_olddocmd)=="function") {
		ae_olddocmd(false);
		document.onmousedown = ae_olddocmd;
	}
	ae_olddocmd = null;

	document.all.frame1.style.pixelWidth = 10;
	document.all.frame1.style.pixelHeight = 10;
	
}

// -->
</script><script language="JavaScript" type="text/JavaScript">
//菜单列表
var menu_table="<table width='100%' cellspacing='0' cellpadding='2'>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_cr.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertTable()'>插入表格</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_sx.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tableProp()'>表格属性</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_sx2.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='cellProp()'>单元格属性</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_tr.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(1)'>插入一行</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_trdel.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(2)'>删除一行</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_td.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(3)'>插入一列</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_tddel.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(4)'>删除一列</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_hby.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(5)'>向右合并</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_hbx.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(6)'>向下合并</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/table_cf.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='tablecommand(7)'>拆分单元格</a></td></tr>";
menu_table+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><a>---</a></td></tr>";
menu_table+="</table>";
var menu_chars="<table width='100%' cellspacing='0' cellpadding='2'>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/chars1.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertChars(0)'>换行符</a></td></tr>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/chars2.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertChars(1)'>版权符号</a></td></tr>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/chars3.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertChars(2)'>注册商标</a></td></tr>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/chars4.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertChars(3)'>商标符号</a></td></tr>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/chars5.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertChars(4)'>圆点</a></td></tr>";
menu_chars+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><a>---</a></td></tr>";
menu_chars+="</table>";

var menu_chars1="<table width='100%' cellspacing='0' cellpadding='2'>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/1.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(0)'>微笑</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/2.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(1)'>痛哭</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/3.gif' width='17' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(2)'>高兴</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/4.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(3)'>傻笑</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/5.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(4)'>酷 ！</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/6.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(5)'>伤心</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/7.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(6)'>汗 ！</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/8.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(7)'>动心</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/9.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(8)'>嚷嚷</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='images/newsface/10.gif' width='13' height='12' align='absmiddle'></td><td><a href='#' onclick='Insertface(9)'>害羞</a></td></tr>";
menu_chars1+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><a>---</a></td></tr>";
menu_chars1+="</table>";

var menu_eq="<table width='100%' cellspacing='0' cellpadding='2'>";
menu_eq+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/eq1.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InsertEQ()'>插入公式</a></td></tr>";
menu_eq+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><img src='Images/eq2.gif' width='16' height='16' align='absmiddle'></td><td><a href='#' onclick='InstallEQ()'>安装公式编辑器插件</a></td></tr>";
menu_eq+="<tr onmouseout='scolor(this)' onmouseover='rcolor(this)'><td><a>-</a></td></tr>";
menu_eq+="</table>";

//下拉菜单相关代码
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var MENU_SHADOW_COLOR='#E1F4EE';//定义下拉菜单阴影色
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array

function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=menu onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=MenuBody>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+2   ;
	t = vSrc.offsetTop + topMar + h + space-32;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
    makeRectangularDropShadow(submenu, MENU_SHADOW_COLOR, 4)
}

function makeRectangularDropShadow(el, color, size)
{
	var i;
	for (i=size; i>0; i--)
	{
		var rect = document.createElement('div');
		var rs = rect.style
		rs.position = 'absolute';
		rs.left = (el.style.posLeft + i) + 'px';
		rs.top = (el.style.posTop + i) + 'px';
		rs.width = el.offsetWidth + 'px';
		rs.height = el.offsetHeight + 'px';
		rs.zIndex = el.style.zIndex - i;
		rs.backgroundColor = color;
		var opacity = 1 - i / (i + 1);
		rs.filter = 'alpha(opacity=' + (100 * opacity) + ')';
		el.insertAdjacentElement('afterEnd', rect);
		global.fo_shadows[global.fo_shadows.length] = rect;
	}
}
function scolor(obj)
{
  obj.style.backgroundColor="";
}
function rcolor(obj)
{
  obj.style.backgroundColor="#E1F4EE";
}
</script>

<script language="javascript">
<!--
function checkdata()
{
if (document.form1.title.value=="")
	{
	  alert("对不起，请输入文章标题！")
	  document.form1.title.focus()
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

//Function to add image
function AddImage(){	
	imagePath = prompt('请输入图片地址', 'http://');				
	
	if ((imagePath != null) && (imagePath != "")){					
		frames.message.document.execCommand('InsertImage', false, imagePath);
  		frames.message.focus();
	}
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
function tablecommand(command)
{
	var cellflag=false;
	var rowflag=false;
	var tableflag=false;
	var cellindex,rowindex,tableref;
	message.focus();
	var xsel=message.document.selection;
	var xobj=message.document.selection.createRange();
	if(xsel.type=="None"||xsel.type=="Text"){
		xsel=xobj.parentElement();
		while(xsel.tagName!="BODY"&&cellflag==false){
			if(xsel.tagName=="TD"){cellindex=xsel.cellIndex;cellflag=true;}
			if(cellflag==false){xsel=xsel.parentElement;}
		}
	}else if(xsel.type=="Control"){
		xsel=xobj.item(0);
		if(xsel.tagName=="TD"){
			cellindex=xsel.cellIndex;
			cellflag=true;
		}else{
			while(xsel.tagName!="BODY"&&cellflag==false){
				if(xsel.tagName=="TD"){cellindex=xsel.cellIndex;cellflag=true;}
				if(cellflag==false){xsel=xsel.parentElement;}
			}
		}
	}
	if(cellflag==true){
		xsel=message.document.selection;
		xobj=message.document.selection.createRange();
		if(xsel.type=="None"||xsel.type=="Text"){
			xsel=xobj.parentElement();
			while(xsel.tagName!="BODY"&&rowflag==false){
				if(xsel.tagName=="TR"){
					rowindex=xsel.rowIndex;
					rowflag=true;
				}
				if(rowflag==false){xsel=xsel.parentElement;}
			}
		}else if(xsel.type=="Control"){
			xsel=xobj.item(0);
			if(xsel.tagName=="TR"){
				rowindex=xsel.rowIndex;
				rowflag=true;
			}else{
				while(xsel.tagName!="BODY"&&rowflag==false){
					if(xsel.tagName=="TR"){
						rowindex=xsel.rowIndex;
						rowflag=true;
					}
					if(rowflag==false){
						xsel=xsel.parentElement;
					}
				}
			}
		}
		xsel=message.document.selection;
		xobj=message.document.selection.createRange();
		if(xsel.type=="None"||xsel.type=="Text"){
			xsel=xobj.parentElement();
			while(xsel.tagName!="BODY"&&tableflag==false){
				if(xsel.tagName=="TABLE"){tableflag=true;}
				if(tableflag==false){xsel=xsel.parentElement;}
			}
		}else if(xsel.type=="Control"){
			xsel=xobj.item(0);
			if(xsel.tagName=="TABLE"){
				tableflag=true;
			}else{
				while(xsel.tagName!="BODY"&&tableflag==false){
					if(xsel.tagName=="TABLE"){tableflag=true;}
					if(tableflag==false){xsel=xsel.parentElement;}
				}
			}
		}
		if(command==3){
			var temprowcount=xsel.rows.length;
			var tempcell;
			var tempspancount=0;
			var tempspanholder;
			var tempcellwidth=xsel.rows[rowindex].cells[cellindex].width;
			var xpositequiv=-1;
			var xposcount=0;
			while(xposcount<=cellindex){
				xpositequiv+=parseInt(xsel.rows[rowindex].cells[xposcount].colSpan);
				xposcount++;
			}
			var ypositequiv=-1;
			var yposcount=0;
			var ymax=xsel.rows[rowindex].cells.length;
			while(yposcount<=ymax-1){
				ypositequiv+=parseInt(xsel.rows[rowindex].cells[yposcount].colSpan);
				yposcount++;
			}
			var idealinsert=xpositequiv+1;
			var zi2=0;
			var zirowtouse=0;
			var zirowtot=xsel.rows.length;
			var rowarray=new Array(zirowtot);
			var rowarray2=new Array(zirowtot);
			for(init1=0;init1<=zirowtot-1;init1++){
				rowarray[init1]=0;
				rowarray2[init1]=0;
			}
			for(zi1=0;zi1<=zirowtot-1;zi1++){
				zi2=0;
				while(zi2<idealinsert&&(rowarray[zi1]==null||rowarray[zi1]<idealinsert)){
					rowarray[zi1]+=parseInt(xsel.rows[zi1].cells[zi2].colSpan);
					rowarray2[zi1]++;
					zi2++;
				}
			}
			var allequal=true;
			var zi3a,zi3b;
			var zthemax=0;
			for(zi3=0;zi3<=zirowtot-1;zi3++){
				zi3a=rowarray[0];
				zi3b=rowarray[zi3];
				if(zi3b>zthemax){zthemax=zi3b;}
				if(zi3a!=zi3b){allequal=false;}
			}
			if(allequal==false){
				var zi4=0;
				var allequal2=true;
				while(zthemax<=ypositequiv&&allequal==false){
					for(zi5=0;zi5<=zirowtot-1;zi5++){
						rowarray[zi5]+=parseInt(xsel.rows[zi5].cells[rowarray2[zi5]].colSpan);
					}
					for(zi3=0;zi3<=zirowtot-1;zi3++){
						zi3a=rowarray[0];
						zi3b=rowarray[zi3];
						if(zi3b>zthemax){zthemax=zi3b;}
						if(zi3a!=zi3b){allequal2=false;}
					}
					if(allequal2==true){allequal=true;}
					for(zi8=0;zi8<=zirowtot-1;zi8++){rowarray2[zi8]++;}
					}
				}
				var zi9;
				for(zi7=0;zi7<=zirowtot-1;zi7++){
					zi9=xsel.rows[zi7].insertCell(rowarray2[zi7]);
					zi9.width=tempcellwidth;
				}
		}else if(command==4){
			var temprowcount=xsel.rows.length;
			for(iccount=0;iccount<=temprowcount-1;iccount++){
				xsel.rows[iccount].deleteCell(cellindex);
			}
			}else if(command==1){
				var tempcell;
				var tempcellb;
				var tempcellcount=xsel.rows[rowindex].cells.length;
				var cellcolarray=new Array(tempcellcount);
				var cellrowarray=new Array(tempcellcount);
				for(cacount=0;cacount<=tempcellcount-1;cacount++){
					cellcolarray[cacount]=xsel.rows[rowindex].cells(cacount).colSpan;
					cellrowarray[cacount]=xsel.rows[rowindex].cells(cacount).rowSpan;
				}
				tempcell=xsel.insertRow(rowindex);
				for(cbcount=0;cbcount<=tempcellcount-1;cbcount++){
					tempcellb=tempcell.insertCell();
					if(cellcolarray[cbcount]!=1){tempcellb.colSpan=cellcolarray[cbcount];}
				}
		}else if(command==2){
				var temprowcount=xsel.rows.length;tempcell=xsel.deleteRow(rowindex);
		}else if(command==5){
				if(xsel.rows[rowindex].cells[cellindex+1]){
					var x=parseInt(xsel.rows[rowindex].cells[cellindex].colSpan)+parseInt(xsel.rows[rowindex].cells[cellindex+1].colSpan);
					var y=xsel.rows[rowindex].cells[cellindex].innerHTML+" "+xsel.rows[rowindex].cells[cellindex+1].innerHTML;
					xsel.rows[rowindex].deleteCell(cellindex+1);
					xsel.rows[rowindex].cells[cellindex].colSpan=x;
					xsel.rows[rowindex].cells[cellindex].innerHTML=y;
				}
		}else if(command==6){
				var yatemprow=xsel.rows.length;
				var yamax=0;
				for(ya1=0;ya1<=yatemprow-1;ya1++){
					var ypositequiv=-1;
					var yposcount=0;
					var ymax=xsel.rows[ya1].cells.length;
					while(yposcount<=ymax-1){
						ypositequiv+=parseInt(xsel.rows[ya1].cells[yposcount].colSpan);
						yposcount++;
					}
					if(ypositequiv>yamax){yamax=ypositequiv;}
				}
				var rowarray=new Array();
				var rowarray2=new Array();
				var myrowcount=xsel.rows.length;
				for(ra1=0;ra1<=myrowcount-1;ra1++){
					rowarray[ra1]=new Array();
					rowarray2[ra1]=0;
					for(cr1=0;cr1<=yamax;cr1++){rowarray[ra1][cr1]=777;}
				}
				var tempra;
				var ra2=0;
				for(ra3=0;ra3<=yamax;ra3++){
					ra2=0;
					while(ra2<=myrowcount-1){
						if(xsel.rows[ra2].cells[ra3]){
							tempra=parseInt(xsel.rows[ra2].cells[ra3].rowSpan);
							if(tempra>1){
								rowarray[ra2][ra3]=ra3+rowarray2[ra2];
								for(zoo=1;zoo<=tempra-1;zoo++){rowarray2[ra2+zoo]--;}
							}
						}
						if(rowarray[ra2][ra3-1]!=ra3+rowarray2[ra2]){
							rowarray[ra2][ra3]=ra3+rowarray2[ra2];
						}else{
							rowarray[ra2][ra3]=555;
						}
						ra2++;
					}
				}
				var samx="";
				var samcount=0;
				for(rx1=0;rx1<=myrowcount-1;rx1++){
					samcount=rowarray[rx1].length;
					for(rx2=0;rx2<=samcount-1;rx2++){
						samx+="-"+rowarray[rx1][rx2];
					}
					samx+="\n";
				}
				var j=parseInt(xsel.rows[rowindex].cells[cellindex].rowSpan);
				var jcount=rowarray[rowindex].length;
				var jval=0;
				for(jc1=0;jc1<=jcount-1;jc1++){
					if(rowarray[rowindex][jc1]==cellindex){jval=jc1;}
				}
				if(xsel.rows[rowindex+j]){
					var cellindex2=rowarray[rowindex+j][jval];
					var x=parseInt(xsel.rows[rowindex].cells[cellindex].rowSpan)+parseInt(xsel.rows[rowindex+j].cells[cellindex2].rowSpan);
					var y=xsel.rows[rowindex].cells[cellindex].innerHTML+" "+xsel.rows[rowindex+j].cells[cellindex2].innerHTML;
					xsel.rows[rowindex+j].deleteCell(cellindex2);
					xsel.rows[rowindex].cells[cellindex].rowSpan=x;
					xsel.rows[rowindex].cells[cellindex].innerHTML=y;
				}
		}else if(command==7){
				var getrowspan=parseInt(xsel.rows[rowindex].cells[cellindex].rowSpan);
				var getcolspan=parseInt(xsel.rows[rowindex].cells[cellindex].colSpan);
				if(getrowspan>1){
					var xr1=getrowspan-1;
					var xrposit=rowindex;
					var xrcposit=cellindex;
					var xrholder;xsel.rows[rowindex].cells[cellindex].rowSpan=1;
					for(xr2=1;xr2<=xr1;xr2++){
						xrholder=xsel.rows[xrposit+xr2].insertCell(xrcposit);
						xrholder.colSpan=xsel.rows[rowindex].cells[cellindex].colSpan;
					}
				}
				if(getcolspan>1){
					var yr1=getcolspan-1;
					var yrposit=rowindex;
					var yrcposit=cellindex;
					var yrholder;xsel.rows[rowindex].cells[cellindex].colSpan=1;
					for(yr2=1;yr2<=yr1;yr2++){
						yrholder=xsel.rows[yrposit].insertCell(yrcposit);
						yrholder.rowSpan=xsel.rows[rowindex].cells[cellindex].rowSpan;
					}
				}
			}
		}
}


function tableProp(){
	var tableflag=false;
	message.focus();
	var xsel=message.document.selection;
	var xobj=message.document.selection.createRange();
	if(xsel.type=="None"||xsel.type=="Text"){
		xsel=xobj.parentElement();
		while(xsel.tagName!="BODY"&&tableflag==false){
			if(xsel.tagName=="TABLE"){tableflag=true;}
			if(tableflag==false){xsel=xsel.parentElement;}
		}
	}else if(xsel.type=="Control"){
		xsel=xobj.item(0);
		if(xsel.tagName=="TABLE"){
			tableflag=true;
		}else{
			while(xsel.tagName!="BODY"&&tableflag==false){
				if(xsel.tagName=="TABLE"){tableflag=true;}
				if(tableflag==false){xsel=xsel.parentElement;}
			}
		}
	}
	if(tableflag==true){
		if(xsel.className!=""&&xsel.className!=null){tableclass=xsel.className;}else{tableclass="";}
		if(xsel.width!=""&&xsel.width!=null){tablewidthspecified="yes";tablewidth=xsel.width;}else{tablewidthspecified="no";tablewith="";}
		if(xsel.align!=""&&xsel.align!=null){tablealign=xsel.align;}else{tablealign="";}
		if(xsel.border!=""&&xsel.border!=null){tablebordersize=xsel.border;}else{tablebordersize="";}
		if(xsel.cellPadding!=""&&xsel.cellPadding!=null){tablecellpadding=xsel.cellPadding;}else{tablecellpadding="";}
		if(xsel.cellSpacing!=""&&xsel.cellSpacing!=null){tablecellspacing=xsel.cellSpacing;}else{tablecellspacing="";}
		if(xsel.borderColor!=""&&xsel.borderColor!=null){tablebordercolor=xsel.borderColor;}else{tablebordercolor="";}
		if(xsel.bgColor!=""&&xsel.bgColor!=null){tablebackgroundcolor=xsel.bgColor;}else{tablebackgroundcolor="";}
		tableiscancel="";
		window.showModalDialog("editor_tableprops.asp",window," dialogWidth: 350px; dialogHeight: 310px; help: no;scroll: no; status: no");
		if(tableiscancel=="no"){
			if(tablewidthspecified=="yes"){
				var tw1="";
				if(tablewidthtype=="percentage"){
					tw1=tablewidth+"%";
				}else{
					tw1=tablewidth;
				}
				xsel.width=tw1;
			}else{
				xsel.removeAttribute("width",0);
			}
			if(tablealign!=""&&tablealign!="Default"){xsel.align=tablealign;}else{xsel.removeAttribute("align",0);}
			if(tableclass!=""&&tableclass!="Default"){xsel.className=tableclass;}else{xsel.removeAttribute("className",0);}
			if(tablebordersize!=""&&tablebordersize!=null){xsel.border=tablebordersize;}else{xsel.removeAttribute("border",0);}
			if(tablecellpadding!=""&&tablecellpadding!=null){xsel.cellPadding=tablecellpadding;}else{xsel.removeAttribute("cellPadding",0);}
			if(tablecellspacing!=""&&tablecellspacing!=null){xsel.cellSpacing=tablecellspacing;}else{xsel.removeAttribute("cellSpacing",0);}
			if(tablebordercolor!=""&&tablebordercolor!="Default"){xsel.borderColor=tablebordercolor;}else{xsel.removeAttribute("borderColor",0);}
			if(tablebackgroundcolor!=""&&tablebackgroundcolor!="Default"){xsel.bgColor=tablebackgroundcolor;}else{xsel.removeAttribute("bgColor",0);}
		}
	}
}

function cellProp(){
	var cellflag=false;
	message.focus();
	var xsel=message.document.selection;
	var xobj=message.document.selection.createRange();
	if(xsel.type=="None"||xsel.type=="Text"){
		xsel=xobj.parentElement();
		while(xsel.tagName!="BODY"&&cellflag==false){
			if(xsel.tagName=="TD"){cellflag=true;}
			if(cellflag==false){xsel=xsel.parentElement;}
		}
	}else if(xsel.type=="Control"){
		xsel=xobj.item(0);
		if(xsel.tagName=="TD"){
			cellflag=true;
		}else{
			while(xsel.tagName!="BODY"&&cellflag==false){
				if(xsel.tagName=="TD"){cellflag=true;}
				if(cellflag==false){xsel=xsel.parentElement;}
			}
		}
	}
	if(cellflag==true){
		if(xsel.width!=""&&xsel.width!=null){tablewidthspecified="yes";tablewidth=xsel.width;}else{tablewidthspecified="no";tablewith="";}
		if(xsel.align!=""&&xsel.align!=null){tablealign=xsel.align;}else{tablealign="";}
		if(xsel.className!=""&&xsel.className!=null){tablecellclass=xsel.className;}else{tablecellclass="";}
		if(xsel.vAlign!=""&&xsel.vAlign!=null){tablevalign=xsel.vAlign;}else{tablevalign="";}
		if(xsel.borderColor!=""&&xsel.borderColor!=null){tablebordercolor=xsel.borderColor;}else{tablebordercolor="";}
		if(xsel.bgColor!=""&&xsel.bgColor!=null){tablebackgroundcolor=xsel.bgColor;}else{tablebackgroundcolor="";}
		tableiscancel="";
		window.showModalDialog("editor_cellprops.asp",window,"dialogWidth: 400px; dialogHeight: 240px;help: no;scroll: no; status: no");
		if(tableiscancel=="no"){
			if(tablewidthspecified=="yes"){
				var tw1="";
				if(tablewidthtype=="percentage"){tw1=tablewidth+"%";}else{tw1=tablewidth;}
				xsel.width=tw1;
			}else{
				xsel.removeAttribute("width",0);
			}
			if(tablealign!=""&&tablealign!="Default"){xsel.align=tablealign;}else{xsel.removeAttribute("align",0);}
			if(tablevalign!=""&&tablevalign!="Default"){xsel.vAlign=tablevalign;}else{xsel.removeAttribute("vAlign",0);}
			if(tablecellclass!=""&&tablecellclass!="Default"){xsel.className=tablecellclass;}else{xsel.removeAttribute("className",0);}
			if(tablebordercolor!=""&&tablebordercolor!="Default"){xsel.borderColor=tablebordercolor;}else{xsel.removeAttribute("borderColor",0);}
			if(tablebackgroundcolor!=""&&tablebackgroundcolor!="Default"){xsel.bgColor=tablebackgroundcolor;}else{xsel.removeAttribute("bgColor",0);}
		}
	}
}
function table_ir()
{
	tablecommand("ir");
}
function table_dr()
{
	tablecommand("dr");
}
function table_ic()
{
	tablecommand("ic");
}
function table_dc()
{
	tablecommand("dc");
}
function table_mc()
{
	tablecommand("mc");
}
function table_md()
{
	tablecommand("md");
}
function table_sc()
{
	tablecommand("sc");
}
function InsertTable()
{
  if (!	validateMode())	return;
  message.focus();
  var range = message.document.selection.createRange();
  var arr = showModalDialog("editor_inserttable.asp", "", "dialogWidth:450px;dialogHeight:200px;help: no; scroll: no; status: no");

  if (arr != null){
	range.pasteHTML(arr);
  }
  message.focus();
}

function about()
{
window.showModalDialog("about.htm","","dialogWidth:420px;dialogHeight:240px;scroll:no;status:no;help:no");
}
function validateMode()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码，进入“编辑”状态，然后再使用系统编辑功能!");
  message.focus();
  return false;
}
function validateModea()
{
  if (!bTextMode) return true;
  alert("请先取消查看HTML源代码!");
  message.focus();
  return false;
}

function InsertChars(CharIndex)
{
  if (!	validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var Chars=new Array("<br>","&copy;","&reg;","&#8482;","&#8226;");
  range.pasteHTML(Chars[CharIndex]);
  message.focus();
}

function Insertface(faceIndex)
{
  if (!	validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var Chars=new Array("<IMG src='images/newsface/1.gif' border=0>","<IMG src='images/newsface/2.gif' border=0>","<IMG src='images/newsface/3.gif' border=0>","<IMG src='images/newsface/4.gif' border=0>","<IMG src='images/newsface/5.gif' border=0>","<IMG src='images/newsface/6.gif' border=0>","<IMG src='images/newsface/7.gif' border=0>","<IMG src='images/newsface/8.gif' border=0>","<IMG src='images/newsface/9.gif' border=0>","<IMG src='images/newsface/10.gif' border=0>");
  range.pasteHTML(Chars[faceIndex]);
  message.focus();
}

function InsertEQ()
{
if (! validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var arr = showModalDialog("editor_inserteq.asp", "", "dialogWidth:40em; dialogHeight:20em; status:0;help:0");
  
  if (arr != null){
    var ss;
    ss=arr.split("*")
    a=ss[0];
    b=ss[1];
    var str1;
    str1="<applet codebase='./' code='webeq3.ViewerControl' WIDTH=320 HEIGHT=100>"
    str1=str1+"<PARAM NAME='parser' VALUE='mathml'><param name='color' value='"+b+"'><PARAM NAME='size' VALUE='18'>"
    str1=str1+"<PARAM NAME=eq id=eq VALUE='"+a+"'></applet>"
    range.pasteHTML(str1);
  }
  message.focus();
}
function InstallEQ()
{
  window.open ("editor_inserteq.asp?Action=Install", "", "height=200, width=300,left="+(screen.AvailWidth-300)/2+",top="+(screen.AvailHeight-200)/2+", toolbar=no, menubar=no, scrollbars=no, resizable=no,location=no, status=no")
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
</script><script language="JavaScript">
function Check(t,n)
{
if(n==1) t.className ="Up"
else
if(n==2) t.className ="Down"
else t.className ="None"
}
function findstr()
{
  if (!	validateMode())	return;
  var arr = showModalDialog("editor_find.asp", window, "dialogWidth:420px; dialogHeight:150px; help: no; scroll: no; status: no");
}
function calculator()
{
  message.focus();
  var range =message.document.selection.createRange();
  var arr = showModalDialog("editor_calculator.asp", "", "dialogWidth:207px; dialogHeight:216px; status:0;help:0");
  
  if (arr != null){
    var ss;
    ss=arr.split("*")
    a=ss[0];
    b=ss[1];
    var str1;
    str1=""+a+""
    range.pasteHTML(str1);
  }
  message.focus();
}
function nowdate()
{
  if (!	validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var d = new Date();
  var str1=d.getYear()+"年"+(d.getMonth() + 1)+"月"+d.getDate() +"日";
  range.pasteHTML(str1);
  message.focus();
}
function nowtime()
{
  if (!	validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var d = new Date();
  var str1=d.getHours() +":"+d.getMinutes()+":"+d.getSeconds();
  range.pasteHTML(str1);
  message.focus();
}
function excel()
{
  if (!	validateMode())	return;
  message.focus();
  var range =message.document.selection.createRange();
  var str1="<object classid='clsid:0002E510-0000-0000-C000-000000000046' id='Spreadsheet1' codebase='file:\\Bob\software\office2000\msowc.cab' width='100%' height='250'><param name='EnableAutoCalculate' value='-1'><param name='DisplayTitleBar' value='0'><param name='DisplayToolbar' value='-1'><param name='ViewableRange' value='1:65536'></object>";
  range.pasteHTML(str1);
  message.focus();
}
function fortable()
{if (!validateMode()) return;
  message.focus();

  var range =message.document.selection.createRange();

  var arr = showModalDialog("table.html", "", "dialogWidth:22em; dialogHeight:20em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  f=ss[5];
  g=ss[6];
  h=ss[7];
  var string;
  string="<table border="+d+" cellpadding="+e+" cellspacing="+f+" style='border-collapse: collapse' bordercolor="+g+" bgcolor="+h+" width=100 align="+c+">";
  for(i=1;i<=a;i++){
  string=string+"<tr>";
  for(j=1;j<=b;j++){
  string=string+"<td width="+100/a+"%></td>";
  }
  string=string+"</tr>";
  }
  string=string+"</table>";
    range.pasteHTML(string);
  }
  else message.focus();
}

function swf()
{if (!validateMode()) return;
 message.focus();

  var range =message.document.selection.createRange();
  var arr = showModalDialog("flash.asp", "", "dialogWidth:27.5em; dialogHeight:14.5em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  f=ss[5];
  g=ss[6];
  h=ss[7];
  j=ss[8];
  var string;
string="<table align="+g+"><tr><td><object width='"+e+"' height='"+f+"' classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000'  codebase='http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0'><param name=movie value='"+a+"'><param name=quality value=high><param name='Play' value='"+b+"'><param name='Loop' value='"+c+"'><param name='Menu' value='"+d+"'><embed src='"+a+"' pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash'></embed></object><br>"+h+"</td></tr></table>"
   form1.PicUrl.value=a;
   form1.PicList.options[form1.PicList.length] = new Option(a,a);
   range.pasteHTML(string);
  }
  else message.focus();
}

function wmv()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("media.asp", "", "dialogWidth:28em; dialogHeight:14em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  f=ss[5];
  g=ss[6];
  h=ss[7];
  var string;
string="<p align="+g+"><embed src="+a+"  width="+e+" height="+f+" autostart="+b+" loop="+c+"></embed><br>"+"<a href="+h+">"+"下载地址"+h+"</a>"+"</p>"
   range.pasteHTML(string);
  }
  else message.focus();
}


function rm()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("rm.asp", "", "dialogWidth:28em; dialogHeight:16em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  f=ss[5];
  g=ss[6];
  var string;
string="<p align="+d+"><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width="+f+" height="+e+"><param name='CONTROLS' value='ImageWindow'><param name=src value="+a+"><param name='CONSOLE' value='Clip1'><param name='AUTOSTART' value="+b+"><param name='LOOP' value="+c+"></object><br><object classid='clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA' width="+f+" height=60><param name='CONTROLS' value='ControlPanel,StatusBar'><param name='CONSOLE' value='Clip1'></object><br>"+"<a href="+g+">"+"下载地址"+g+"</a>"+"</p>"
  range.pasteHTML(string);
 }
  else message.focus();
}


function pic()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("pic.asp", "", "dialogWidth:26.5em; dialogHeight:13em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  f=ss[5];
  var str1;
str1="<a title=点击图片看全图 target='_blank' href='"+a+"'>"
str1=str1+"<p align=center><img src='"+a+"' alt='"+b+"' style='"+c+"'"
str1=str1+" onload='javascript:if(this.width>screen.width-333)this.width=screen.width-333'"
str1=str1+" border=0 align='"+e+"'></a><br>"+d+"</P>"
   form1.PicUrl.value=a;
   form1.PicList.options[form1.PicList.length] = new Option(a,a);
   range.pasteHTML(str1);
   }
  else message.focus();
}
function page()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="[---分页---]"
   range.pasteHTML(str1);
}

function watermark()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<font color='#FFFFFF'>[<%=rs8("name")%>]</font>"
   range.pasteHTML(str1);
}

function adduction()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();
str1="<P><TABLE style='BORDER-RIGHT: #6595d6 1px dotted; TABLE-LAYOUT: fixed; BORDER-TOP: #6595d6 1px dotted; BORDER-LEFT: #6595d6 1px dotted; BORDER-BOTTOM: #6595d6 1px dotted' cellSpacing=0 cellPadding=6 width='95%' align=center border=0><TBODY><TR><TD style='WORD-WRAP: break-word' bgColor=#e8f4ff><FONT style='FONT-WEIGHT: bold; COLOR: #990000'>以下是引用片段：</FONT><BR></TD></TR></TBODY></TABLE></P>"
   range.pasteHTML(str1);
}

function FIELDSET()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("fieldset.htm", "", "dialogWidth:16.5em; dialogHeight:13.5em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  var str1;
str1="<FIELDSET "
str1=str1+"align="+a+""
str1=str1+" style='"
if(c.value!='')str1=str1+"color:"+c+";"
if(d.value!='')str1=str1+"background-color:"+d+";"
str1=str1+"'><Legend"
str1=str1+" align="+b+""
str1=str1+">标题</Legend>内容</FIELDSET>"
  range.pasteHTML(str1);
  }
  else message.focus();
}

function iframe()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("iframe.htm", "", "dialogWidth:26.5em; dialogHeight:13em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  var str1;
str1="<iframe src='"+a+"'"
str1+=" scrolling="+b+""
str1+=" frameborder="+c+""
if(d!='')str1+=" marginheight="+d
if(e!='')str1+=" marginwidth="+e
str1=str1+"></iframe>"
  range.pasteHTML(str1);
 }
  else message.focus();
}

function hr()
{if (!validateMode()) return;
message.focus();

var range =message.document.selection.createRange();

  var arr = showModalDialog("hr.htm", "", "dialogWidth:20em; dialogHeight:13.5em; status:0;help:0");
  
  if (arr != null){
  var ss;
  ss=arr.split("*")
  a=ss[0];
  b=ss[1];
  c=ss[2];
  d=ss[3];
  e=ss[4];
  var str1;
str1="<hr"
str1=str1+" color='"+a+"'"
str1=str1+" size="+b+"'"
str1=str1+" "+c+""
str1=str1+" align="+d+""
str1=str1+" width="+e
str1=str1+">"
  range.pasteHTML(str1);
  }
  else message.focus();
}

function insertbr(br){

	document.form1.title.value+=br;
	document.form1.title.focus();
}

function insertbr1(br){

	document.form1.hongtou.value+=br;
	document.form1.hongtou.focus();
}

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
		case 6:	form1.title.style.removeAttribute ("fontStyle");
				form1.title.style.removeAttribute ("fontWeight");
				form1.title.style.textDecoration = "strikethrough";
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
		case 1: form1.title.style.family = "宋体";
				break;
		case 2:	form1.title.style.family = "楷体_GB2312";
				break;
		case 3:	form1.title.style.family = "新宋体";
				break;
		case 4:	form1.title.style.family = "黑体";
				break;
		case 5:	form1.title.style.family = "隶书";
				break;
		case 6:	form1.title.style.family = "幼圆";
				break;
	}
	form1.title.focus ();
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}


</script><script language=javascript>
function fnPreHandle()
{if (!validateModea()) return;

var iCount;
var strData; 
var iMaxChars = 50000;
var iBottleNeck = 2000000;
var strHTML;

strData = frames.message.document.body.innerHTML;

if (strData.length > iBottleNeck)
{
if (confirm("您要发布的文章太长，建议您拆分为几部分分别发布。\n如果您坚持提交，注意需要较长时间才能提交成功。\n\n是否坚持提交？") == false)
return false;
}

iCount = parseInt(strData.length / iMaxChars) + 1;

//hdnCount
strHTML = "<input type=hidden name=hdnCount value=" + iCount + ">";

for (var i = 1; i <= iCount; i++)
{
strHTML = strHTML + "\n" + "<input type=hidden name=hdnBigField" + i + ">";
}

document.all.divHidden.innerHTML = strHTML;

for (var i = 1; i <= iCount; i++)
{
form1.elements["hdnBigField" + i].value = strData.substring((i - 1) * iMaxChars, i * iMaxChars);
}

}
</script>
<link rel="stylesheet" type="text/css" href="site.css">
<link rel="STYLESHEET" type="text/css" href="editor.css">
</head>
<body topmargin="0">
<!--#include file=Admin_Top.asp-->
<div id=menuDiv style='Z-INDEX: 1000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<%set rs5=server.CreateObject("ADODB.RecordSet")
rs5.Source="select * from "& db_System_Table &""
rs5.Open rs5.source,conn,1,1
dim typename
set rs6=server.CreateObject("ADODB.RecordSet")
rs6.Source="select * from "& db_Type_Table &" where typeid="&rs("typeid")
rs6.Open rs6.source,conn,1,1
typename=rs6("typename")
rs6.close
set rs6=nothing
%>
<div align="center">
<center>
<table border="1" width="100%" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=border%>">
  <tr> 
    <td colspan="4" height="25" width="100%"  class="TDtop"> 
      <div align="center">┊ <B>在<b>[<a href="BigClassManage.asp?TypeID=<%=rs("typeid")%>"><%=typename%></a>]-[<a href="SmallClassManage.asp?BigClassID=<%=rs("bigclassid")%>"><%=rs2("bigclassname")%></a>]<%if rs("smallclassid")<>"" then%>-[<a href="ListNews.asp?SmallClassID=<%=rs("smallclassid")%>"><%=rs3("smallclassname")%></a>]<%end if%></b>中修改文章，ID号为<%=newsid%></B> ┊</div>
    </td>
  </tr>
  <%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="selfreg" then%>
  <form action=NewsClassEdit1.asp?NewsID=<%=NewsID%> method=post>
    <tr> 
      <td width="10%" bgcolor="#FFFFFF"></td>
      <td width="90%" bgcolor="#FFFFFF" colspan="3" align=right> &nbsp; 
        <input name=btn_Modi_class type=submit value="修改文章类别" onMouseOver="window.status='按这个按钮修改文章类别';return true;" onMouseOut="window.status='';return true;" title="按这个按钮修改文章类别"  style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
      </td>
    </tr>
  </form><%end if%>
  <form method="POST" OnSubmit="return checkdata()" onReset="return ResetForm();" name="form1" action="NewsEditOK.asp?NewsID=<%=NewsID%>">
    <tr> 
      <td align="right" bgcolor="#FFFFFF">所属专题：</td>
      <td bgcolor="#FFFFFF" colspan="3"> &nbsp; 
        <select size="1" name="SpecialID" onMouseOver="window.status='点这里选择文章专题';return true;" onMouseOut="window.status='';return true;">
        <%  set rs4=server.CreateObject ("ADODB.RecordSet")
rs4.Source="select * from "& db_Special_Table &""
rs4.Open rs4.source,conn,1,1%>
          <option value=0 <%if rs("SpecialID")=0 then Response.Write "selected"%>>-----------------</option>
          <%if rs4.EOF then %>
          <option value=0>暂无任何专题</option>
          <%else
while not rs4.EOF
SpecialID=rs4("SpecialID")
%>
          <option value=<%=SpecialID%> <%if rs("SpecialID")=SpecialID then Response.Write " selected"%>><%=trim(rs4("SpecialName"))%></option>
          <%
rs4.MoveNext
wend
end if
rs4.close 
set rs4=nothing  
%>
        </select>
         &nbsp;第二专题 
        <select size="1" name="SpecialID2" onMouseOver="window.status='点这里选择文章专题';return true;" onMouseOut="window.status='';return true;">
           <%  set rs4=server.CreateObject ("ADODB.RecordSet")
rs4.Source="select * from "& db_Special_Table &""
rs4.Open rs4.source,conn,1,1%>
          <option value=0 <%if rs("SpecialID2")=0 then Response.Write "selected"%>>-----------------</option>
          <%if rs4.EOF then %>
          <option value=0>暂无任何专题</option>
          <%else
while not rs4.EOF
SpecialID2=rs4("SpecialID")
%>
          <option value=<%=SpecialID2%> <%if rs("SpecialID2")=SpecialID2 then Response.Write " selected"%>><%=trim(rs4("SpecialName"))%></option>
          <%
rs4.MoveNext
wend
end if
rs4.close 
set rs4=nothing  
%>
        </select>
      </td>
    </tr>
    <tr> 
      <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">文章标题：</td>
      <td bgcolor="#FFFFFF" colspan="3"> <span style='cursor:hand' title='缩短对话框' onClick='if (me.size>10)me.size=me.size-2'>-</span> 
        <input name="title" id=me type="TEXT" size=60 maxlength=100 style="background-color:ffffff;border: 1 double" value="<%=trim(rs("Title"))%>" onMouseOver="window.status='在这里修改文章标题';return true;" onMouseOut="window.status='';return true;" title="在这里修改文章标题">
        <span style='cursor:hand' title='加长对话框' onClick='if (me.size<102)me.size=me.size+2'>+</span>
<SELECT size=1 id=title_color name=title_color LANGUAGE=javascript onchange="return title_color_onchange()" style="width:55">
<OPTION value="">颜色</OPTION>
<OPTION <%if rs("titlecolor")="0" then%>selected<%end if%> value="0">缺省</OPTION>
<OPTION <%if rs("titlecolor")="#000000" then%>selected<%end if%> value="#000000" style="background-color:#000000"></OPTION>
<OPTION <%if rs("titlecolor")="#FFFFFF" then%>selected<%end if%> value="#FFFFFF" style="background-color:#FFFFFF"></OPTION>
<OPTION <%if rs("titlecolor")="#008000" then%>selected<%end if%> value="#008000" style="background-color:#008000"></OPTION>
<OPTION <%if rs("titlecolor")="#800000" then%>selected<%end if%> value="#800000" style="background-color:#800000"></OPTION>
<OPTION <%if rs("titlecolor")="#808000" then%>selected<%end if%> value="#808000" style="background-color:#808000"></OPTION>
<OPTION <%if rs("titlecolor")="#000080" then%>selected<%end if%> value="#000080" style="background-color:#000080"></OPTION>
<OPTION <%if rs("titlecolor")="#800080" then%>selected<%end if%> value="#800080" style="background-color:#800080"></OPTION>
<OPTION <%if rs("titlecolor")="#808080" then%>selected<%end if%> value="#808080" style="background-color:#808080"></OPTION>
<OPTION <%if rs("titlecolor")="#FFFF00" then%>selected<%end if%> value="#FFFF00" style="background-color:#FFFF00"></OPTION>
<OPTION <%if rs("titlecolor")="#00FF00" then%>selected<%end if%> value="#00FF00" style="background-color:#00FF00"></OPTION>
<OPTION <%if rs("titlecolor")="#00FFFF" then%>selected<%end if%> value="#00FFFF" style="background-color:#00FFFF"></OPTION>
<OPTION <%if rs("titlecolor")="#FF00FF" then%>selected<%end if%> value="#FF00FF" style="background-color:#FF00FF"></OPTION>
<OPTION <%if rs("titlecolor")="#FF0000" then%>selected<%end if%> value="#FF0000" style="background-color:#FF0000"></OPTION>
<OPTION <%if rs("titlecolor")="#0000FF" then%>selected<%end if%> value="#0000FF" style="background-color:#0000FF"></OPTION>
<OPTION <%if rs("titlecolor")="#008080" then%>selected<%end if%> value="#008080" style="background-color:#008080"></OPTION>
</SELECT>

<SELECT size=1 id=title_type name=title_type LANGUAGE=javascript onchange="return title_type_onchange()" style="width:55">
<OPTION value="0">类型</OPTION>
<OPTION <%if rs("titletype")="0" then%>selected<%end if%> value="0">普通</OPTION>
<OPTION <%if rs("titletype")="b" then%>selected<%end if%> value="b">加粗</OPTION>
<OPTION <%if rs("titletype")="i" then%>selected<%end if%> value="i">倾斜</OPTION>
<OPTION <%if rs("titletype")="u" then%>selected<%end if%> value="u">下划线</OPTION>

</SELECT>

</SELECT>
<SELECT size=1 id=title_size name=title_size LANGUAGE=javascript onchange="return title_size_onchange()" style="width:55">
<OPTION value="0" selected>不评</OPTION>
<OPTION <%if rs("titlesize")="0" then%>selected<%end if%> value="0">不评</OPTION>
<OPTION <%if rs("titlesize")="1" then%>selected<%end if%> value="1">要评</OPTION>
</SELECT>

      </td>
    </tr>
    <tr> 
      <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">链接文章：</td>
      <td bgcolor="#FFFFFF" colspan="3"> <span style='cursor:hand' title='缩短对话框' onClick='if (title_face.size>10)title_face.size=title_face.size-2'>-</span> 
        <input name="title_face" id=title_face type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" value="<%=trim(rs("titleface"))%>" onMouseOver="window.status='文章为直接链接文章时，请填写网址';return true;" onMouseOut="window.status='';return true;" title="文章为直接链接文章时，请填写网址">
        <span style='cursor:hand' title='加长对话框' onClick='if (title_face.size<102)title_face.size=title_face.size+2'>+</span> 
      (文章为直接链接文章时，请填写网址)</td>
    </tr>
	<tr> 
      <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">文章来源：</td>
      <td bgcolor="#FFFFFF" colspan="3"> <span style='cursor:hand' title='缩短对话框' onClick='if (message.size>10)message.size=message.size-2'>-</span> 
        <input name="Original" id=message type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" value="<%=trim(rs("Original"))%>" onMouseOver="window.status='按这个按钮修改文章来源';return true;" onMouseOut="window.status='';return true;" title="按这个按钮修改文章来源">
        <span style='cursor:hand' title='加长对话框' onClick='if (message.size<102)message.size=message.size+2'>+</span> 
      (网页模版时，请填写网址)</td>
    </tr>
    <tr> 
      <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">文章作者：</td>
      <td bgcolor="#FFFFFF" colspan="3"><span style='cursor:hand' title='缩短对话框' onClick='if (mess.size>10)mess.size=mess.size-2'>-</span> 
        <input name="Author" id=mess type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" value="<%=trim(rs("Author"))%>" onMouseOver="window.status='按这个按钮修改文章作者';return true;" onMouseOut="window.status='';return true;" title="按这个按钮修改文章作者">
        <span style='cursor:hand' title='加长对话框' onClick='if (mess.size<102)mess.size=mess.size+2'>+</span> 
      </td>
    </tr>

<tr> 
                <td align="right" valign="top" class="unnamed2" bgcolor="#FFFFFF">文章内容：<br><font color="#FF0000">*</font><br><br><br><br><br><br><br><br><br><br>换行请按Shift+Enter <br><br>另起一段请按Enter</td>
                <td colspan="3">&nbsp;
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
 <select language="javascript"  id="FontSize" title="字号大小" onChange="FormatText('fontsize',this[this.selectedIndex].value);" name="select" onMouseOver="window.status='选择选定文字的字号大小。';return true;" onMouseOut="window.status='';return true;">
                <option class="heading" selected>字号 
                <option value="7">一号 
                <option value="6">二号 
                <option value="5">三号 
                <option value="4">四号 
                <option value="3">五号 
                <option value="2">六号 
                <option value="1">七号</option>
              </select><%if request.cookies(Forcast_SN)("key")="super" then%><input onClick="setMode(this.checked);" name="viewhtml" type="checkbox" value="ON">查看HTML源代码<%end if%><br>
		&nbsp;
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/selectall.gif" align="absmiddle" border="0" alt="全部选择" onClick="FormatText('selectall')" style="cursor: hand;">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/find.gif" align="middle" border="0" alt="查找" onClick="findstr()" style="cursor: hand;">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/cut.gif"  align="absmiddle" onClick="FormatText('cut')" style="cursor: hand;" alt="剪切"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)"src="images/copy.gif"  align="absmiddle" onClick="FormatText('copy')" style="cursor: hand;" alt="复制"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/paste.gif"  align="absmiddle" onClick="FormatText('paste')" style="cursor: hand;" alt="粘贴"> 		
	          <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/DELETE.gif" align="absmiddle" border="0" alt="删除" onClick="FormatText('DELETE')" style="cursor: hand;">
			  <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/clear.gif" align="absmiddle" border="0" alt="删除文字格式" onClick="FormatText('RemoveFormat')" style="cursor: hand;">
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/undo.gif" align="absmiddle" border="0" alt="撤消" onClick="FormatText('undo')" style="cursor: hand;">
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/redo.gif" align="absmiddle" border="0" alt="恢复" onClick="FormatText('redo')" style="cursor: hand;">
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/bold.gif" align="absmiddle" alt="粗体" onClick="FormatText('bold', '')" style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/italic.gif" align="absmiddle" alt="斜体" onClick="FormatText('italic', '')" style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/underline.gif" align="absmiddle" alt="下划线" onClick="FormatText('underline', '')" style="cursor: hand;">
			  <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/strikethrough.gif" align="absmiddle" alt="删除线" onClick="FormatText('strikethrough', '')" style="cursor: hand;">
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/ALEFT.gif" align="absmiddle" onClick="FormatText('Justifyleft', '')" style="cursor: hand;" alt="左对齐"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/center.gif" align="absmiddle" border="0" alt="居中" onClick="FormatText('JustifyCenter', '')" style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/aright.gif" align="absmiddle" onClick="FormatText('JustifyRight', '')" style="cursor: hand;" alt="右对齐"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/list.gif" align="absmiddle" border="0" alt="项目符号" onClick="FormatText('InsertUnorderedList', '')" style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/number.gif" align="absmiddle" alt="编号" border="0" onClick="FormatText('insertorderedlist', '')" style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/outdent.gif" align="absmiddle" onClick="FormatText('Outdent', '')" style="cursor: hand;" alt="回退"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/indent.gif" align="absmiddle" border="0" alt="缩进" onClick="FormatText('indent', '')" style="cursor: hand;"> 
		      <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/line.gif" align="absmiddle" alt="普通水平线" border="0" onClick="FormatText('InsertHorizontalRule', '')"  style="cursor: hand;">
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/sline.gif" align="absmiddle" alt="特殊水平线" border="0" onClick="hr()"  style="cursor: hand;"> 
              <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/link.gif" align="absmiddle" border="0" alt="超级链接" onClick="FormatText('createLink')" style="cursor: hand;">
		      <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/unlink.gif" align="absmiddle" border="0" alt="取消超级链接" onClick="FormatText('unLink')" style="cursor: hand;">
		      <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/sup.gif" align="absmiddle" border="0" alt="上标" onClick="FormatText('superscript')" style="cursor: hand;">
		      <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/sub.gif" align="absmiddle" border="0" alt="下标" onClick="FormatText('subscript')" style="cursor: hand;">
		<br>&nbsp;&nbsp;<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/calculator.gif"  align="middle" onClick="calculator()" style="cursor: hand;" alt="计算">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/date.gif" align="middle" border="0" style="cursor:hand;" alt="插入日期" LANGUAGE="javascript" onclick="nowdate()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/time.gif" align="middle" border="0" style="cursor:hand;" alt="插入时间" LANGUAGE="javascript" onclick="nowtime()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/FIELDSET.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入栏目框" LANGUAGE="javascript" onclick="FIELDSET()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/iframe.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入网页" LANGUAGE="javascript" onclick="iframe()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/excel.gif" align="middle" border="0" style="cursor:hand;" alt="插入Excel表格" LANGUAGE="javascript" onclick="excel()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/table.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入表格" LANGUAGE="javascript" onclick="InsertTable()"><img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/arrow.gif" align="middle" border="0" style="cursor:hand;" alt="表格操作" LANGUAGE="javascript" onclick="ShowMenu(menu_table,100)">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/eq.gif" align="middle" border="0" style="cursor:hand;" alt="插入公式" LANGUAGE="javascript" onclick="InsertEQ()"><img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/arrow.gif" align="middle" border="0" style="cursor:hand;" alt="公式操作" LANGUAGE="javascript" onclick="ShowMenu(menu_eq,100)">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/chars.gif" align="middle" border="0" style="cursor:hand;" alt="插入换行符号" LANGUAGE="javascript" onclick="InsertChars(0)"><img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/arrow.gif" align="middle" border="0" style="cursor:hand;" alt="更多特殊符号" LANGUAGE="javascript" onclick="ShowMenu(menu_chars,100)">
<!--增加表情-->
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/face.gif" align="middle" border="0" style="cursor:hand;" alt="插入表情" LANGUAGE="javascript" onclick="InsertChars(0)"><img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="IMAGES/arrow.gif" align="middle" border="0" style="cursor:hand;" alt="更多表情" LANGUAGE="javascript" onclick="ShowMenu(menu_chars1,100)">
<!--增加表情结束-->
<!--增加引用样式-->
        <img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/adduction.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入引用样式" lANGUAGE="javascript" onclick="adduction()">
<!--增加引用样式结束-->
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/swf.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入FLASH" LANGUAGE="javascript" onclick="swf()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/mp.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入视频文件，支持格式为：avi、wmv、asf" LANGUAGE="javascript" onclick="wmv()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/real.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入RealPlay文件，支持格式为：rm、ra、ram" LANGUAGE="javascript" onclick="rm()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/image.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入网上图片，支持格式为：gif、jpg、png、bmp" lANGUAGE="javascript" onclick="pic()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/page.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入分页符" lANGUAGE="javascript" onclick="page()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/watermark.gif" align="absmiddle" border="0" style="cursor:hand;" alt="插入水印版权保护" lANGUAGE="javascript" onclick="watermark()">
		<img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/help.gif" align="absmiddle" border="0" style="cursor:hand;" alt="使用帮助" lANGUAGE="javascript" onclick="help()">
		<a href=javascript:about()><img class=None onmousedown="Check(this,2)" onmouseover="Check(this,1)" onmouseout="Check(this,0)" onmouseup="Check(this,1)" src="images/about.gif" align="absmiddle" border="0" style="cursor:hand;" alt="关于本系统"></a>
		<br>&nbsp;&nbsp;<iframe style="top:2px" ID="UploadFiles" src="upme.htm" frameborder=0 scrolling=no width="280" height="25"></iframe></div>
                           <script language="javascript">
                           var bTextMode=false;
		    	document.write ('<iframe src="textbox.asp?action=modify&newsid=<%=newsid%>" id="message" width="97%" height="200" align=center></iframe>')
                frames.message.document.designMode = "On";
                </script>
                <DIV id=divHidden></DIV></td>
              </tr>  
    <tr> 
      <td align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">相关文章：</td>
      <td bgcolor="#FFFFFF" colspan="3"><span style='cursor:hand' title='缩短对话框' onClick='if (ss.size>10)ss.size=ss.size-2'>-</span> 
        <input name="about" id=ss type="TEXT" size=30 maxlength=100 style="background-color:ffffff;border: 1 double" value="<%=trim(rs("about"))%>" onMouseOver="window.status='在这里修改关键字';return true;" onMouseOut="window.status='';return true;" title="在这里修改关键字">
        <span style='cursor:hand' title='加长对话框' onClick='if (ss.size<102)ss.size=ss.size+2'>+</span> 
        填入关键字或完整标题 </td>
	</tr>
	<tr> 
		<td width="79" align="right" class="unnamed2" valign="middle" bgcolor="#FFFFFF">首页图片：</td>
		<td width="680" bgcolor="#FFFFFF">&nbsp;&nbsp;
			<input name="PicUrl" type="text" id="PicUrl" size="30" maxlength="200" style="background-color:ffffff;border: 1 double" value="<%=trim(rs("picname"))%>">
				用于在首页的图片文章处显示或者直接从上传图片中选择：<select name="PicList" id="PicList" onChange="PicUrl.value=this.value;">
			<option selected value=" <%=rs("picname")%> "><%=trim(rs("picname"))%></option> 
			<option>不指定首页图片</option>
			<%
		set rss=server.createobject("adodb.recordset")
		sqls="select * from "& db_Attach_Table &" where NewsID=" & NewsID 
		rss.open sqls,conn,1,3
		do while not rss.eof
			if right(rss("filename"),3)="jpg" or right(rss("filename"),3)="bmp" or right(rss("filename"),3)="png" or right(rss("filename"),3)="gif" or right(rss("filename"),3)="swf" then
				response.write "<option value='" & rss("filename") & "'>" & rss("filename") & "</option>"
			end if
			RSs.MoveNext
		loop
		rss.close
		set rss=nothing%>
		</select>
		</td>
	</tr>
	<tr> 
		<td width="79" align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >保存图片：</td>
		<td width="680" valign="middle" bgcolor="#FFFFFF" >&nbsp;
			<input type="radio" value="1"  name="getphoto" onMouseOver="window.status='点击这里，将保存远程图片';return true;" onMouseOut="window.status='';return true;" title='点击这里，将保存远程图片'>是 
			<input type="radio" value="0" checked name="getphoto">否&nbsp;&nbsp;保存远程图片&nbsp;&nbsp;(从其它网站上复制内容，并且内容中包含有图片)</td>
              </tr>
     
    <%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" then%>
    <tr> 
      <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF">推荐文章：</td>
      <td valign="middle" bgcolor="#FFFFFF" colspan="3">&nbsp; 
        <input type="radio" value="1" <%if rs("goodnews")=1 then Response.Write "checked"%>  name="goodnews">
        是 
        <input type="radio" value="0" <%if rs("goodnews")=0 then Response.Write "checked"%> name="goodnews">
        否&nbsp;&nbsp;生成推荐文章</td>
    </tr>
<tr> 
      <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF">固顶文章：</td>
      <td valign="middle" bgcolor="#FFFFFF" colspan="3">&nbsp; 
        <input type="radio" value="1" <%if rs("istop")=1 then Response.Write "checked"%>  name="istop">
        是
        <input type="radio" value="0" <%if rs("istop")=0 then Response.Write "checked"%> name="istop">
        否&nbsp;&nbsp;生成固顶文章</td>
    </tr>
<tr> 
                <td align="right" valign="middle" class="unnamed2" bgcolor="#FFFFFF" >自动缩进：</td>
                <td valign="middle" bgcolor="#FFFFFF" >&nbsp; 
                  <input type="radio" value="1" <%if rs("backtype")=1 then Response.Write "checked"%>  name="backtype">
                  是 
                  <input type="radio" value="0" <%if rs("backtype")=0 then Response.Write "checked"%> name="backtype">
                  否&nbsp;&nbsp;文章首行缩进排版</td>
</tr>	

<%if uselevel=1 then%>
<tr> 
                <td width="79" align="right" bgcolor="#FFFFFF">文章级别：</td>
                <td width="680" bgcolor="#FFFFFF" colspan="3"> &nbsp; <select size="1" name="newslevel" onMouseOver="window.status='在这里选择您要添加的文章分级浏览级别';return true;" onMouseOut="window.status='';return true;">
<%for i=0 to 3%>
                    <option <%if rs("newslevel")=i then%>selected<%end if%> value="<%=i%>"><%=i%></option>
<%next%>
                    </select>文章级别指可浏览该文章会员的级别。0为游客，1为普通会员，2为高级会员，3为特级会员。
                  </td>
              </tr>
   <%end if%><%end if%>
        <tr> 
      <td colspan="4" align="center" height="25" width="100%"  class="TDtop"> 
<input name="newsrelated" type="hidden" value="<%=newsrelated%>">
<input name="checkked" type="hidden" value="<%if request.cookies(Forcast_SN)("key")="super" or request.cookies(Forcast_SN)("key")="check" or request.cookies(Forcast_SN)("key")="typemaster" or request.cookies(Forcast_SN)("key")="bigmaster" then%>1<%else%><%if request.cookies(Forcast_SN)("key")="selfreg" and fabiaocheck="1" then%>1<%else%><%if request.cookies(Forcast_SN)("key")="smallmaster" and rs5("usecheck")="1" then%>1<%else%>0<%end if%><%end if%><%end if%>"><%rs5.close%>
<input name="encode" type="hidden" value="html">
<input type="button" value=" 返 回 " onclick="javascript:history.go(-1)" class="unnamed5" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
<input type="submit" value=" 修 改 " name="Submit" class="unnamed5" OnClick="fnPreHandle()" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">&nbsp;
<input type="reset" value=" 清 除 " name="Reset" class="unnamed5" style="font-size: 9pt;  color: #000000; background-color: #EAEAF4; solid #EAEAF4" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#EAEAF4'">
	<%rs8.close
	set rs8=nothing%>
</td>
</tr>
</form>
</table>
</center>
</div>
</body>
</html>
<!--#include file=Admin_Bottom.asp-->
	<%else
		Show_Err("您无权修改！<br><br><a href='javascript:history.back()'>回去重来</a>")
		response.end	
	end if
	if rs("smallclassid")<>"" then
		rs3.close 
		set rs3=nothing
	end if
	rs.close 
	set rs=nothing 
	rs2.close 
	set rs2=nothing
	conn.close  
	set conn=nothing
end if%>
