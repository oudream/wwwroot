<SCRIPT language=javascript>
//==========================
//        菜单开始
//==========================
mtDropDown.spacerGif = "images/Menu_x.gif"; 
mtDropDown.dingbatOn = "images/Menu_expand-1.gif"; 
mtDropDown.dingbatOff = "images/Menu_expand-0.gif"; 
mtDropDown.dingbatSize = 14; 
mtDropDown.menuPadding = 1; 
mtDropDown.itemPadding = 4; 
mtDropDown.shadowSize = 2; 
mtDropDown.shadowOffset = 3; 
mtDropDown.shadowColor = "#888"; 
mtDropDown.shadowPng = "images/005.png"; 
mtDropDown.backgroundColor = "#EAEAF4"; 
mtDropDown.backgroundPng = "images/004.png"; 
mtDropDown.hideDelay = 200; 
mtDropDown.slideTime = 200; 

mtDropDown.reference = {topLeft:1,topRight:2,bottomLeft:3,bottomRight:4};
mtDropDown.direction = {down:1,right:2};
mtDropDown.registry = [];
mtDropDown._maxZ = 100;

mtDropDown.isSupported = function() {
if (typeof mtDropDown.isSupported.r == "boolean") 
return mtDropDown.isSupported.r;
var ua = navigator.userAgent.toLowerCase();
var an = navigator.appName;
var r = false;
if (ua.indexOf("gecko") > -1) r = true; 
else if (an == "Microsoft Internet Explorer") {
if (document.getElementById) r = true; 
}
mtDropDown.isSupported.r = r;
return r;
}

mtDropDown.initialize = function() {
for (var i = 0, menu = null; menu = this.registry[i]; i++) {
menu.initialize();
}
}

mtDropDown.renderAll = function() {
var aMenuHtml = [];
for (var i = 0, menu = null; menu = this.registry[i]; i++) {
aMenuHtml[i] = menu.toString();
}
document.write(aMenuHtml.join(""));
}

function mtDropDown(oActuator, iDirection, iLeft, iTop, iReferencePoint, parentMenuSet) {

this.addItem = addItem;
this.addMenu = addMenu;
this.toString = toString;
this.initialize = initialize;
this.isOpen = false;
this.show = show;
this.hide = hide;
this.items = [];

this.onactivate = new Function(); 
this.ondeactivate = new Function(); 
this.onmouseover = new Function(); 
this.onqueue = new Function(); 

this.index = mtDropDown.registry.length;
mtDropDown.registry[this.index] = this;
var id = "mtDropDown" + this.index;
var contentHeight = null;
var contentWidth = null;
var childMenuSet = null;
var animating = false;
var childMenus = [];
var slideAccel = -1;
var elmCache = null;
var ready = false;
var _this = this;
var a = null;
var pos = iDirection == mtDropDown.direction.down ? "top" : "left";
var dim = null;

function addItem(sText, sUrl) {
var item = new mtDropDownItem(sText, sUrl, this);
item._index = this.items.length;
this.items[item._index] = item;
}
function addMenu(oMenuItem) {
if (!oMenuItem.parentMenu == this) throw new Error("Cannot add a menu here");
if (childMenuSet == null) childMenuSet = new mtDropDownSet(mtDropDown.direction.right, -5, 2, mtDropDown.reference.topRight);
var m = childMenuSet.addMenu(oMenuItem);
childMenus[oMenuItem._index] = m;
m.onmouseover = child_mouseover;
m.ondeactivate = child_deactivate;
m.onqueue = child_queue;
return m;
}
function initialize() {
initCache();
initEvents();
initSize();
ready = true;
}
function show() {

if (ready) {
_this.isOpen = true;
animating = true;
setContainerPos();
elmCache["clip"].style.visibility = "visible";
elmCache["clip"].style.zIndex = mtDropDown._maxZ++;

slideStart();
_this.onactivate();
}
}
function hide() {
if (ready) {
_this.isOpen = false;
animating = true;
for (var i = 0, item = null; item = elmCache.item[i]; i++) 
dehighlight(item);
if (childMenuSet) childMenuSet.hide();
slideStart();
_this.ondeactivate();
}
}
function setContainerPos() {
var sub = oActuator.constructor == mtDropDownItem; 
var act = sub ? oActuator.parentMenu.elmCache["item"][oActuator._index] : oActuator; 
var el = act;
var x = 0;
var y = 0;
var minX = 0;
var maxX = (window.innerWidth ? window.innerWidth : document.body.clientWidth) - parseInt(elmCache["clip"].style.width);
var minY = 0;
var maxY = (window.innerHeight ? window.innerHeight : document.body.clientHeight) - parseInt(elmCache["clip"].style.height);

while (sub ? el.parentNode.className.indexOf("mtDropdownMenu") == -1 : el.offsetParent) {
x += el.offsetLeft;
y += el.offsetTop;
if (el.scrollLeft) x -= el.scrollLeft;
if (el.scrollTop) y -= el.scrollTop;
el = el.offsetParent;
}
if (oActuator.constructor == mtDropDownItem) {
x += parseInt(el.parentNode.style.left);
y += parseInt(el.parentNode.style.top);
}
switch (iReferencePoint) {
case mtDropDown.reference.topLeft:
break;
case mtDropDown.reference.topRight:
x += act.offsetWidth;
break;
case mtDropDown.reference.bottomLeft:
y += act.offsetHeight;
break;
case mtDropDown.reference.bottomRight:
x += act.offsetWidth;
y += act.offsetHeight;
break;
}
x += iLeft;
y += iTop;
x = Math.max(Math.min(x, maxX), minX);
y = Math.max(Math.min(y, maxY), minY);
elmCache["clip"].style.left = x + "px";
elmCache["clip"].style.top = y + "px";
}
function slideStart() {
var x0 = parseInt(elmCache["content"].style[pos]);
var x1 = _this.isOpen ? 0 : -dim;
if (a != null) a.stop();
a = new Accelimation(x0, x1, mtDropDown.slideTime, slideAccel);
a.onframe = slideFrame;
a.onend = slideEnd;
a.start();
}
function slideFrame(x) {
elmCache["content"].style[pos] = x + "px";
}
function slideEnd() {
if (!_this.isOpen) elmCache["clip"].style.visibility = "hidden";
animating = false;
}
function initSize() {

var ow = elmCache["items"].offsetWidth;
var oh = elmCache["items"].offsetHeight;
var ua = navigator.userAgent.toLowerCase();

elmCache["clip"].style.width = ow + mtDropDown.shadowSize + 2 + "px";
elmCache["clip"].style.height = oh + mtDropDown.shadowSize + 2 + "px";

elmCache["content"].style.width = ow + mtDropDown.shadowSize + "px";
elmCache["content"].style.height = oh + mtDropDown.shadowSize + "px";
contentHeight = oh + mtDropDown.shadowSize;
contentWidth = ow + mtDropDown.shadowSize;
dim = iDirection == mtDropDown.direction.down ? contentHeight : contentWidth;

elmCache["content"].style[pos] = -dim - mtDropDown.shadowSize + "px";
elmCache["clip"].style.visibility = "hidden";

if (ua.indexOf("mac") == -1 || ua.indexOf("gecko") > -1) {

elmCache["background"].style.width = ow + "px";
elmCache["background"].style.height = oh + "px";
elmCache["background"].style.backgroundColor = mtDropDown.backgroundColor;

elmCache["shadowRight"].style.left = ow + "px";
elmCache["shadowRight"].style.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize) + "px";
elmCache["shadowRight"].style.backgroundColor = mtDropDown.shadowColor;



elmCache["shadowBottom"].style.top = oh + "px";
elmCache["shadowBottom"].style.width = ow - mtDropDown.shadowOffset + "px";
elmCache["shadowBottom"].style.backgroundColor = mtDropDown.shadowColor;
}

else {

elmCache["background"].firstChild.src = mtDropDown.backgroundPng;
elmCache["background"].firstChild.width = ow;
elmCache["background"].firstChild.height = oh;

elmCache["shadowRight"].firstChild.src = mtDropDown.shadowPng;
elmCache["shadowRight"].style.left = ow + "px";
elmCache["shadowRight"].firstChild.width = mtDropDown.shadowSize;
elmCache["shadowRight"].firstChild.height = oh - (mtDropDown.shadowOffset - mtDropDown.shadowSize);



elmCache["shadowBottom"].firstChild.src = mtDropDown.shadowPng;
elmCache["shadowBottom"].style.top = oh + "px";
elmCache["shadowBottom"].firstChild.height = mtDropDown.shadowSize;
elmCache["shadowBottom"].firstChild.width = ow - mtDropDown.shadowOffset;
}
}
function initCache() {
var menu = document.getElementById(id);
var all = menu.all ? menu.all : menu.getElementsByTagName("*"); 
elmCache = {};
elmCache["clip"] = menu;
elmCache["item"] = [];
for (var i = 0, elm = null; elm = all[i]; i++) {
switch (elm.className) {
case "items":
case "content":
case "background":
case "shadowRight":
case "shadowBottom":
elmCache[elm.className] = elm;
break;
case "item":
elm._index = elmCache["item"].length;
elmCache["item"][elm._index] = elm;
break;
}
}

_this.elmCache = elmCache;
}
function initEvents() {

for (var i = 0, item = null; item = elmCache.item[i]; i++) {
item.onmouseover = item_mouseover;
item.onmouseout = item_mouseout;
item.onclick = item_click;
}

if (typeof oActuator.tagName != "undefined") {
oActuator.onmouseover = actuator_mouseover;
oActuator.onmouseout = actuator_mouseout;
}

elmCache["content"].onmouseover = content_mouseover;
elmCache["content"].onmouseout = content_mouseout;
}
function highlight(oRow) {
oRow.className = "item hover";
if (childMenus[oRow._index]) 
oRow.lastChild.firstChild.src = mtDropDown.dingbatOn;
}
function dehighlight(oRow) {
oRow.className = "item";
if (childMenus[oRow._index]) 
oRow.lastChild.firstChild.src = mtDropDown.dingbatOff;
}
function item_mouseover() {
if (!animating) {
highlight(this);
if (childMenus[this._index]) 
childMenuSet.showMenu(childMenus[this._index]);
else if (childMenuSet) childMenuSet.hide();
}
}
function item_mouseout() {
if (!animating) {
if (childMenus[this._index])
childMenuSet.hideMenu(childMenus[this._index]);
else 
dehighlight(this);
}
}
function item_click() {
if (!animating) {
if (_this.items[this._index].url) 
location.href = _this.items[this._index].url;
}
}
function actuator_mouseover() {
parentMenuSet.showMenu(_this);
}
function actuator_mouseout() {
parentMenuSet.hideMenu(_this);
}
function content_mouseover() {
if (!animating) {
parentMenuSet.showMenu(_this);
_this.onmouseover();
}
}
function content_mouseout() {
if (!animating) {
parentMenuSet.hideMenu(_this);
}
}
function child_mouseover() {
if (!animating) {
parentMenuSet.showMenu(_this);
}
}
function child_deactivate() {
for (var i = 0; i < childMenus.length; i++) {
if (childMenus[i] == this) {
dehighlight(elmCache["item"][i]);
break;
}
}
}
function child_queue() {
parentMenuSet.hideMenu(_this);
}
function toString() {
var aHtml = [];
var sClassName = "mtDropdownMenu" + (oActuator.constructor != mtDropDownItem ? " top" : "");
for (var i = 0, item = null; item = this.items[i]; i++) {
aHtml[i] = item.toString(childMenus[i]);
}
return '<div id="' + id + '" class="' + sClassName + '">' + 
'<div class="content"><table class="items" cellpadding="0" cellspacing="0" border="0">' + 
'<tr><td colspan="2"><img src="' + mtDropDown.spacerGif + '" width="1" height="' + mtDropDown.menuPadding + '"></td></tr>' + 
aHtml.join('') + 
'<tr><td colspan="2"><img src="' + mtDropDown.spacerGif + '" width="1" height="' + mtDropDown.menuPadding + '"></td></tr></table>' + 
'<div class="shadowBottom"><img src="' + mtDropDown.spacerGif + '" width="1" height="1"></div>' + 
'<div class="shadowRight"><img src="' + mtDropDown.spacerGif + '" width="1" height="1"></div>' + 
'<div class="background"><img src="' + mtDropDown.spacerGif + '" width="1" height="1"></div>' + 
'</div></div>';
}
}

mtDropDownSet.registry = [];
function mtDropDownSet(iDirection, iLeft, iTop, iReferencePoint) {

this.addMenu = addMenu;
this.showMenu = showMenu;
this.hideMenu = hideMenu;
this.hide = hide;

var menus = [];
var _this = this;
var current = null;
this.index = mtDropDownSet.registry.length;
mtDropDownSet.registry[this.index] = this;

function addMenu(oActuator) {
var m = new mtDropDown(oActuator, iDirection, iLeft, iTop, iReferencePoint, this);
menus[menus.length] = m;
return m;
}
function showMenu(oMenu) {
if (oMenu != current) {

if (current != null) hide(current); 

current = oMenu;

oMenu.show();
}
else {

cancelHide(oMenu);
}
}
function hideMenu(oMenu) {

if (current == oMenu && oMenu.isOpen) {

if (!oMenu.hideTimer) scheduleHide(oMenu);
}
}
function scheduleHide(oMenu) {

oMenu.onqueue();
oMenu.hideTimer = window.setTimeout("mtDropDownSet.registry[" + _this.index + "].hide(mtDropDown.registry[" + oMenu.index + "])", mtDropDown.hideDelay);
}
function cancelHide(oMenu) {

if (oMenu.hideTimer) {
window.clearTimeout(oMenu.hideTimer);
oMenu.hideTimer = null;
}
}
function hide(oMenu) { 
if (!oMenu && current) oMenu = current;
if (oMenu && current == oMenu && oMenu.isOpen) {

cancelHide(oMenu);
current = null;
oMenu.hideTimer = null;
oMenu.hide();
}
}
}

function mtDropDownItem(sText, sUrl, oParent) {
this.toString = toString;
this.text = sText;
this.url = sUrl;
this.parentMenu = oParent;
function toString(bDingbat) {
var sDingbat = bDingbat ? mtDropDown.dingbatOff : mtDropDown.spacerGif;
var iEdgePadding = mtDropDown.itemPadding + mtDropDown.menuPadding;
var sPaddingLeft = "padding:" + mtDropDown.itemPadding + "px; padding-left:" + iEdgePadding + "px;"
var sPaddingRight = "padding:" + mtDropDown.itemPadding + "px; padding-right:" + iEdgePadding + "px;"
return '<tr class="item"><td nowrap style="' + sPaddingLeft + '">' + 
sText + '</td><td width="14" style="' + sPaddingRight + '">' + 
'<img src="' + sDingbat + '" width="14" height="14"></td></tr>';
}
}

function Accelimation(from, to, time, zip) {
if (typeof zip == "undefined") zip = 0;
if (typeof unit == "undefined") unit = "px";
this.x0 = from;
this.x1 = to;
this.dt = time;
this.zip = -zip;
this.unit = unit;
this.timer = null;
this.onend = new Function();
this.onframe = new Function();
}

Accelimation.prototype.start = function() {
this.t0 = new Date().getTime();
this.t1 = this.t0 + this.dt;
var dx = this.x1 - this.x0;
this.c1 = this.x0 + ((1 + this.zip) * dx / 3);
this.c2 = this.x0 + ((2 + this.zip) * dx / 3);
Accelimation._add(this);
}

Accelimation.prototype.stop = function() {
Accelimation._remove(this);
}

Accelimation.prototype._paint = function(time) {
if (time < this.t1) {
var elapsed = time - this.t0;
this.onframe(Accelimation._getBezier(elapsed/this.dt,this.x0,this.x1,this.c1,this.c2));
}
else this._end();
}

Accelimation.prototype._end = function() {
Accelimation._remove(this);
this.onframe(this.x1);
this.onend();
}

Accelimation._add = function(o) {
var index = this.instances.length;
this.instances[index] = o;

if (this.instances.length == 1) {
this.timerID = window.setInterval("Accelimation._paintAll()", this.targetRes);
}
}

Accelimation._remove = function(o) {
for (var i = 0; i < this.instances.length; i++) {
if (o == this.instances[i]) {
this.instances = this.instances.slice(0,i).concat( this.instances.slice(i+1) );
break;
}
}

if (this.instances.length == 0) {
window.clearInterval(this.timerID);
this.timerID = null;
}
}

Accelimation._paintAll = function() {
var now = new Date().getTime();
for (var i = 0; i < this.instances.length; i++) {
this.instances[i]._paint(now);
}
}

Accelimation._B1 = function(t) { return t*t*t }
Accelimation._B2 = function(t) { return 3*t*t*(1-t) }
Accelimation._B3 = function(t) { return 3*t*(1-t)*(1-t) }
Accelimation._B4 = function(t) { return (1-t)*(1-t)*(1-t) }

Accelimation._getBezier = function(percent,startPos,endPos,control1,control2) {
return endPos * this._B1(percent) + control2 * this._B2(percent) + control1 * this._B3(percent) + startPos * this._B4(percent);
}

Accelimation.instances = [];
Accelimation.targetRes = 10;
Accelimation.timerID = null;

</SCRIPT>

<SCRIPT language=javascript>
//原 mt_dropdown_initialize.js
		var preloaded = [];

		// here is a lame preloading script i am putting in just for demonstration
		for (var i = 1; i <= 8; i++) {
			preloaded[i] = [loadImage(i + "-0.gif"), loadImage(i + "-1.gif")];
		}

		function init() {
			if (mtDropDown.isSupported()) {
				mtDropDown.initialize();
			}
		}

		// this is a crappy example preloader. Use whichever one you want.
		function loadImage(sFilename) {
			var img = new Image();
			img.src ="images/" + sFilename;
			return img;
		}
</SCRIPT>
</head>
<body bgcolor="#a9aaa5" onload="init();" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
						<table border=0 cellpadding=0 cellspacing=0 width="100%" align="left">
							<tbody> 
								<tr> 
								<td class="top" nowrap>
<%sqltype="select * from "& db_Type_Table & " order by typeorder"			'选择分类
Set rstype= Server.CreateObject("ADODB.Recordset")
rstype.open sqltype,conn,1,1

dim Type_i,strtype
Type_i=1
typeNum=rstype.recordcount
strtype= "<span style=' height:9px;color=#FFFFFF'>|</span>"& chr(10) &""
if rstype.bof and rstype.eof then 
	strtype = strtype & "栏目正在建设中| "
else
	rstype.movefirst
		do while not rstype.eof
			'菜单条(大类)
			strtype = strtype & "<span style=' height:9px;color=#FFFFFF'><a class=MenuType href='Type.asp?TypeId="& rstype("typeid") &"' id=menu"& Type_i &" >"& rstype("typeName") &"</a>&nbsp;|</span>"& chr(10) &""
			rstype.movenext
			Type_i=Type_i+1
		loop
end if
response.write strtype

if B_BG=0 then	'判断是否

	Type_i=1
	dim second_i,SubMenuContent
	rstype.movefirst
	typeNum=rstype.recordcount
	if rstype.bof and rstype.eof then 

	else
		response.write "<SCRIPT LANGUAGE=javascript>"& chr(10) &""
		response.write "if (mtDropDown.isSupported()) {"& chr(10) &""
		response.write "	var ms = new mtDropDownSet(mtDropDown.direction.down, 0, 0, mtDropDown.reference.bottomLeft);"& chr(10) &""
	
		do while not rstype.eof
			SubMenuContent = ""
			second_i=0

			SqlBigClass="select * from "& db_BigClass_Table & "  where typeid="& rstype("typeid") &" Order by bigclassorder"	'选择大组类
			Set rsBigClass= Server.CreateObject("ADODB.Recordset")
			rsBigClass.open SqlBigClass,conn,1,1
			if rsBigClass.bof and rsBigClass.eof then
				strMenu="没有任何子栏目"
				'定义一级子菜单
				response.write chr(10)
				response.write "	var menu"& Type_i &" = ms.addMenu(document.getElementById('menu"& Type_i &"'));"& chr(10) &""	
				'一级子菜单
				response.write "	menu"& Type_i &".addItem('无子栏目',''); "& chr(10) &""

				'response.write SubMenuContent
			else
				'定义一级子菜单
				response.write chr(10)
				response.write "	var menu"& Type_i &" = ms.addMenu(document.getElementById('menu"& Type_i &"'));"& chr(10) &""	
				do while not rsBigClass.eof

					'一级子菜单(大组类)
					response.write "	menu"& Type_i &".addItem('- "& rsBigClass("BigClassName") &"','BigClass.asp?TypeId="& rsBigClass("TypeId") &"&BigClassId="& rsBigClass("BigClassId") &"'); "& chr(10) &""
					
					'二级子菜单(小组类)
					SqlSmallClass="select * from "& db_SmallClass_Table & "  where BigClassId="& rsBigClass("BigClassId") &" Order by smallclassorder"	'选择小组类
					Set rsSmallClass= Server.CreateObject("ADODB.Recordset")
					rsSmallClass.open SqlSmallClass,conn,1,1
					if rsSmallClass.bof and rsSmallClass.eof then
						strSmallMenu="没有任何子栏目"
						second_i=second_i+1
					else
						'定义二级子菜单		
						SubMenuContent = SubMenuContent + ""& chr(10) &"	var subMenu"& second_i &" = menu"& Type_i &".addMenu(menu"& Type_i &".items["& second_i &"]);"& chr(10) &""

						do while not rsSmallClass.eof
							'二级子菜单(小组类)
							SubMenuContent = SubMenuContent + "	subMenu"& second_i &".addItem('- "& rsSmallClass("SmallClassName") &"','SmallClass.asp?TypeId="& rstype("TypeId") &"&BigClassID="& rsSmallClass("BigClassId") &"&smallclassid="& rsSmallClass("SmallClassId") &"');"& chr(10) &""
							rsSmallClass.movenext
						loop
						second_i=second_i+1
					end if
					rsBigClass.movenext
				loop
				response.write SubMenuContent
			end if
			rstype.movenext
			Type_i=Type_i+1
		loop
			response.write chr(10)
		response.write "	mtDropDown.renderAll();"& chr(10) &""
		response.write "	}"& chr(10) &""
		response.write "	//=========================="& chr(10) &""
		response.write "	//        菜单结束"& chr(10) &""
		response.write "	//=========================="& chr(10) &""
		response.write "</SCRIPT>"& chr(10) &""
	end if
end if

rsSmallClass.close
set rsSmallClass=nothing

rsBigClass.close
set rsBigClass=nothing

rstype.close
set rstype=nothing
%>
								</td>
								</tr>
							</tbody> 
						</table>
