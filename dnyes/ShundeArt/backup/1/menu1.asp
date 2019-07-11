<script>var mmenus    = new Array();
var misShow   = new Boolean(); 
misShow=false;
var misdown   = new Boolean();
misdown=false;
var musestatus=false;
var mpopTimer = 0;
mmenucolor='#535396';mfontcolor='';mmenuoutcolor='#7C7CB5';mmenuincolor='';mmenuoutbordercolor='';mmenuinbordercolor='#000000';mmidoutcolor='';mmidincolor='';mmenuovercolor='#ffffff';mitemedge='0';msubedge='0';mmenuunitwidth=100;mmenuitemwidth=100;mmenuheight=0;mmenuwidth='100%';mmenuadjust=0;mmenuadjustV=50;mfonts='font-family: ËÎÌו; font-size: 9pt; color: #ffffff; ';mcursor='hand';
var fadeSteps = 5;
var fademsec = 20;
var fadeArray = new Array();
function fade(el, fadeIn, steps, msec) {
	if (steps == null) steps = fadeSteps;
	if (msec == null) msec = fademsec;
	if (el.fadeIndex == null)
		el.fadeIndex = fadeArray.length;
	fadeArray[el.fadeIndex] = el;
	if (el.fadeStepNumber == null) {
		if (el.style.visibility == "hidden")
			el.fadeStepNumber = 0;
		else
			el.fadeStepNumber = steps;
		if (fadeIn)
			el.style.filter = "Alpha(Opacity=0)";
		else
			el.style.filter = "Alpha(Opacity=80)";
	}
	window.setTimeout("repeatFade(" + fadeIn + "," + el.fadeIndex + "," + steps + "," + msec + ")", msec);
}
function repeatFade(fadeIn, index, steps, msec) {	
	el = fadeArray[index];
	
	c = el.fadeStepNumber;
	if (el.fadeTimer != null)
		window.clearTimeout(el.fadeTimer);
	if ((c == 0) && (!fadeIn)) {
		el.style.visibility = "hidden";
		return;
	}
	else if ((c==steps) && (fadeIn)) {
		el.style.filter = "Alpha(Opacity=80);";
		el.style.visibility = "visible";
		return;
	}
	else {
		(fadeIn) ? 	c++ : c--;
		el.style.visibility = "visible";
		el.style.filter = "Alpha(Opacity=" + 80*c/steps + ")";
		el.fadeStepNumber = c;
		el.fadeTimer = window.setTimeout("repeatFade(" + fadeIn + "," + index + "," + steps + "," + msec + ")", msec);
	}
}

function stoperror(){
return true;
}
window.onerror=stoperror;
function mpopOut() {
mpopTimer = setTimeout('mallhide()', 500);
}
function getReal(el, type, value) {
	temp = el;
	while ((temp != null) && (temp.tagName != "BODY")) {
		if (eval("temp." + type) == value) {
			el = temp;
			return el;
		}
		temp = temp.parentElement;
	}
	return el;
}


function mMenuRegister(menu) 
{
  mmenus[mmenus.length] = menu
  return (mmenus.length - 1)
}
function mMenuItem(caption,command,target,isline,statustxt,img,sizex,sizey,pos){
	this.caption=caption;
	this.command=command;
	this.target=target;
	this.isline=isline;
	this.statustxt=statustxt;
	this.img=img;
	this.sizex=sizex;
	this.sizey=sizey;
	this.pos=pos;
}
function mMenu(caption,command,target,img,sizex,sizey,pos){
	this.items = new Array();
	this.caption=caption;
	this.command=command;
	this.target=target;
	this.img=img;
	this.sizex=sizex;
	this.sizey=sizey;
	this.pos=pos;
	this.id=mMenuRegister(this);
}
function mMenuAddItem(item)
{
  this.items[this.items.length] = item
  item.parent = this.id;
  this.children=true;
}

mMenu.prototype.addItem = mMenuAddItem;
function mtoout(src){

src.style.borderLeftColor=mmenuoutbordercolor;
src.style.borderRightColor=mmenuinbordercolor;
src.style.borderTopColor=mmenuoutbordercolor;
src.style.borderBottomColor=mmenuinbordercolor;
src.style.backgroundColor=mmenuoutcolor;
src.style.color=mmenuovercolor;
}
function mtoin(src){

src.style.borderLeftColor=mmenuinbordercolor;
src.style.borderRightColor=mmenuoutbordercolor;
src.style.borderTopColor=mmenuinbordercolor;
src.style.borderBottomColor=mmenuoutbordercolor;
src.style.backgroundColor=mmenuincolor;
src.style.color=mmenuovercolor;
}
function mnochange(src){
src.style.borderLeftColor=mmenucolor;
src.style.borderRightColor=mmenucolor;
src.style.borderTopColor=mmenucolor;
src.style.borderBottomColor=mmenucolor;
src.style.backgroundColor='';
src.style.color=mfontcolor;

}
function mallhide(){
	for(var nummenu=0;nummenu<mmenus.length;nummenu++){
		var themenu=document.all['mMenu'+nummenu]
		var themenudiv=document.all['mmenudiv'+nummenu]
                mnochange(themenu);
                mmenuhide(themenudiv);
                }
}
function mmenuhide(menuid){
menuid.style.filter='Alpha(Opacity=100)';
fade(menuid,false,6);
misShow=false;
}
function mmenushow(menuid,pid){
menuid.style.filter='Alpha(Opacity=80)';
menuid.style.left=mposflag.offsetLeft+pid.offsetLeft+mmenuadjust+15;menuid.style.top=mmenutable.offsetTop+mmenutable.offsetHeight+mmenuadjustV+111;
if(mmenuitemwidth+parseInt(menuid.style.left)>document.body.clientWidth+document.body.scrollLeft)
menuid.style.left=document.body.clientWidth+document.body.scrollLeft-mmenuitemwidth;
fade(menuid,true,6);
misShow=true;
}
function mmenu_over(menuid,x){
toel = getReal(window.event.toElement, "className", "coolButton");
fromel = getReal(window.event.fromElement, "className", "coolButton");
if (toel == fromel) return;
if(x<0){
  misShow = false;
  mallhide();
  mtoout(eval("mMenu"+x));
}else{

  mallhide();
  mtoin(eval("mMenu"+x));
  mmenushow(menuid,eval("mMenu"+x));

}
clearTimeout(mpopTimer);
}
function mmenu_out(x){
toel = getReal(window.event.toElement, "className", "coolButton");
fromel = getReal(window.event.fromElement, "className", "coolButton");
if (toel == fromel) return;
if (misShow){
mtoin(eval("mMenu"+x));
}else{
mnochange(eval("mMenu"+x));
}
mpopOut()
}
function mmenu_down(menuid,x){
  if(misShow){
  mmenuhide(menuid);
  mtoout(eval("mMenu"+x));
  }
  else{
  mtoin(eval("mMenu"+x));
  mmenushow(menuid,eval("mMenu"+x));
  misdown=true;
  }
}
function mmenu_up(){
  misdown=false;
}
function mmenuitem_over(x,i){
srcel = getReal(window.event.srcElement, "className", "coolButton");
if(misdown){
	mtoin(srcel);
}
else{
mtoout(srcel);
}
mthestatus = mmenus[x].items[i].statustxt;
if(mthestatus!=""){
	musestatus=true;
	window.status=mthestatus;
}
clearTimeout(mpopTimer);
}
function mmenuitem_out(){
srcel = getReal(window.event.srcElement, "className", "coolButton");
mnochange(srcel);
if(musestatus)window.status="";
mpopOut()
}
function mmenuitem_down(){
srcel = getReal(window.event.srcElement, "className", "coolButton");
mtoin(srcel)
misdown=true;
}
function mmenuitem_up(){
srcel = getReal(window.event.srcElement, "className", "coolButton");
mtoout(srcel)
misdown=false;
}
function mexec2(x){
var cmd;
if(mmenus[x].target=="blank"){
  cmd = "window.open('"+mmenus[x].command+"')";
}else{
  cmd = mmenus[x].target+".location=\""+mmenus[x].command+"\"";
}
eval(cmd);
}
function mexec(x,i){
var cmd;
if(mmenus[x].items[i].target=="blank"){
  cmd = "window.open('"+mmenus[x].items[i].command+"')";
}else{
  cmd = mmenus[x].items[i].target+".location=\""+mmenus[x].items[i].command+"\"";
}
eval(cmd);
}
function mbody_click(){

if (misShow){
	srcel = getReal(window.event.srcElement, "className", "coolButton");
	for(var x=0;x<=mmenus.length;x++){
		if(srcel.id=="mMenu"+x)
		return;
	}
	mallhide();
}
}
document.onclick=mbody_click;
function mwritetodocument(){
      var mwb=0;
                     var stringx='<div id="mposflag" align=center style="position:absolute;"></div><table  id=mmenutable border=0 cellpadding=0 cellspacing=2 height='+mmenuheight+' onselectstart="event.returnValue=false"'+
                     ' style="filter:Alpha(Opacity=80);cursor:'+mcursor+';'+mfonts+
                     ' border-left: '+mwb+'px solid '+mmenuoutbordercolor+';'+
                     ' border-right: '+mwb+'px solid '+mmenuinbordercolor+'; '+
                     'border-top: '+mwb+'px solid '+mmenuoutbordercolor+'; '+
                     'border-bottom: '+mwb+'px solid '+mmenuinbordercolor+'; padding:0px"><tr>'
                     for(var x=0;x<mmenus.length;x++){
                     	var thismenu=mmenus[x];
                     	var imgsize="";
                     	if(thismenu.sizex!="0"||thismenu.sizey!="0")imgsize=" width="+thismenu.sizex+" height="+thismenu.sizey;
                     	var ifspace="";
                     	if(thismenu.caption!="")ifspace="&nbsp;";
                     	stringx += "<td nowrap class=coolButton id=mMenu"+x+" style='border: "+mitemedge+"px solid "+mmenucolor+
                     	"' onmouseover=mmenu_over(mmenudiv"+x+
                     	","+x+") onmouseout=mmenu_out("+x+
                     	") onmousedown=mmenu_down(mmenudiv"+x+","+x+")";
                     	      if(thismenu.command!=""){
                     	          stringx += " onmouseup=mmenu_up();mexec2("+x+");";
                     	      }else{
                     	      	  stringx += " onmouseup=mmenu_up()";
                     	      }
                     	      if(thismenu.pos=="0"){
                     	          stringx += " align=center><img align=absmiddle src='"+thismenu.img+"'"+imgsize+">"+ifspace+thismenu.caption+"</td>";	
                     	      }else if(thismenu.pos=="1"){
                     	          stringx += " align=center>"+thismenu.caption+ifspace+"<img align=absmiddle src='"+thismenu.img+"'"+imgsize+"></td>";	
                     	      }else if(thismenu.pos=="2"){
                     	          stringx += " align=center background='"+thismenu.img+"'>"+thismenu.caption+"</td>";	
                     	      }else{
                     	          stringx += " align=center>"+thismenu.caption+"</td>";
                     	      }
                     	stringx += "";
                     }
                     stringx+="<td width=*></td></tr></table>";
                     
                     
                     for(var x=0;x<mmenus.length;x++){
                     	thismenu=mmenus[x];
                        if(x<0){
                        stringx+='<div id=mmenudiv'+x+' style="visiable:none"></div>';
                        }else{
                        stringx+='<div id=mmenudiv'+x+
                        ' style="filter:Alpha(Opacity=80);cursor:'+mcursor+';position:absolute;'+
                        'width:'+mmenuitemwidth+'px; z-index:'+(x+100);
                        if(mmenuinbordercolor!=mmenuoutbordercolor&&msubedge=="0"){
                        stringx+=';border-left: 1px solid '+mmidoutcolor+
                        ';border-top: 1px solid '+mmidoutcolor;}
                        stringx+=';border-right: 1px solid '+mmenuinbordercolor+
                        ';border-bottom: 1px solid '+mmenuinbordercolor+';visibility:hidden" onselectstart="event.returnValue=false">\n'+
                     	'<table  width="100%" border="0" height="100%" align="center" cellpadding="0" cellspacing="0" '+
                     	'style="'+mfonts+' border-left: 0px solid '+mmenuoutbordercolor;
                        if(mmenuinbordercolor!=mmenuoutbordercolor&&msubedge=="0"){
                     	stringx+=';border-right: 0px solid '+mmidincolor+
                     	';border-bottom: 1px solid '+mmidincolor;}
                     	stringx+=';border-top: 0px solid '+mmenuoutbordercolor+
                     	';padding: 0px" bgcolor='+mmenucolor+'>\n'
                     	for(var i=0;i<thismenu.items.length;i++){
                     		var thismenuitem=thismenu.items[i];
                     		var imgsize="";
                     	        if(thismenuitem.sizex!="0"||thismenuitem.sizey!="0")imgsize=" width="+thismenuitem.sizex+" height="+thismenuitem.sizey;
                     	        var ifspace="";
                     	        if(thismenu.caption!="")ifspace="&nbsp;";
                     		if(!thismenuitem.isline){
                     		stringx += "<tr><td class=coolButton style='border: "+mitemedge+"px solid "+mmenucolor+
                     		"' width=100% height=15px onmouseover=\"mmenuitem_over("+x+","+i+
                     		");\" onmouseout=mmenuitem_out() onmousedown=mmenuitem_down() onmouseup=";
 				stringx += "mmenuitem_up();mexec("+x+","+i+"); ";
 				if(thismenuitem.pos=="0"){
                     	            stringx += "><img align=absmiddle src='"+thismenuitem.img+"'"+imgsize+">"+ifspace+thismenuitem.caption+"</td></tr>";	
                     	          }else if(thismenuitem.pos=="1"){
                     	            stringx += ">"+thismenuitem.caption+ifspace+"<img align=absmiddle src='"+thismenuitem.img+"'"+imgsize+"></td></tr>";	
                     	          }else if(thismenuitem.pos=="2"){
                     	            stringx += "background='"+thismenuitem.img+"'>"+thismenuitem.caption+"</td></tr>";	
                     	          }else{
                     	            stringx += ">"+thismenuitem.caption+"</td></tr>";
                     	          }
 				}else{
                     		stringx+='<tr><td height="1" background="hr.gif" onmousemove="clearTimeout(mpopTimer);"><img height="1" width="1" src="none.gif" border="0"></td></tr>\n';
                     		}
                     	}stringx+='</table>\n</div>'
                     	}
                     	
                }
                
                     document.write("<div align='center'>"+stringx+"</div>");
}
function mcheckLocation(){
if(isNaN(mmenuwidth))mmenuwidth=document.body.clientWidth*parseInt(mmenuwidth.substring(0,3))/100;ym=eval(document.body.scrollTop)+0;xm=eval(document.body.scrollLeft)+0;y=mmenutable.style.pixelTop;x=mmenutable.style.pixelLeft;if(Math.abs(ym-y)>1)mmenutable.style.pixelTop=y+=(ym-y)/5;else mmenutable.style.pixelTop=y=ym;if(Math.abs(xm-x)>1)mmenutable.style.pixelLeft=x+=(xm-x)/5;else mmenutable.style.pixelLeft=x=xm;setTimeout("mcheckLocation()",10);}

<%
set rs=server.CreateObject("ADODB.RecordSet")
rs.Source="select * from type order by typeorder"
rs.Open rs.Source,conn,1,1
i=1
while not rs.EOF
RecordCount=rs.RecordCount
menuid=rs("typeid")
menuname=rs("typename")
menucontent=rs("typecontent")
menuview=rs("typeview")

%>
mpmenu<%=menuid%>=new mMenu('<span style="filter:dropshadow(color=#2A2A2A,offx=1,offy=1,positive=1); height:9px;">|&nbsp;<%=menuName%> </span>','type.asp?typeID=<%=menuid%>','self','','','','');
<%set rs1=server.CreateObject("ADODB.RecordSet")
    set rs1=conn.execute("SELECT * FROM bigclass where typeid="&menuid&" order by bigclassorder") 
    if not rs1.EOF then
	do while not rs1.eof

    %>
                  <%if not Rs1.eof then%>

mpmenu<%=menuid%>.addItem(new mMenuItem('&nbsp;<span style="line-height: 20px;"><%=Rs1("bigclassname")%>','bigclass.asp?typeid=<%=menuid%>&bigclassid=<%=Rs1("bigclassid")%>','self',false,'<%=Rs1("bigclassname")%>','','','',''));
<%
end if
rs1.movenext   
	
	  loop
	  end if
	rs1.close
%>
<%
i=i+1

rs.MoveNext
wend
rs.close
set rs=nothing
%>
mwritetodocument();
mcheckLocation();

           </script>