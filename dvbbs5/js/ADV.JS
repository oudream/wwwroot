function Adv()
{	var nowTime = new Date();
	var BasePath = IsRoot ? "" : "../";
	
	bust = Math.floor(1000000*Math.random());

	if(arguments.length>2)
	{
	sImg=arguments[3];
	if (document.all)document.write("<div id=waiting style=position:absolute;top:0px;left:0px;z-index:1;visibility:hidden>");
	else document.write("<layer name=waiting top=0 left=0 visibility=hide zIndex=2>");
	document.write("<table border=2 cellspacing=1 cellpadding=0 bordercolorlight=#FFFFFF bordercolordark=#C0C0C0 bgcolor=#E0E0E0><tr><td bgcolor=#E0E0E0>"+sImg+"</td></tr><tr><td bgcolor=#E0E0E0>");
	if(document.all)document.write("<img src=images/space.gif width=1 height=10 name=sbar style=background-color:#6699cc></td></tr></table></div>");
	else {
		document.write("<img src=images/space.gif width=1 height=10></td></tr></table></layer>");
		document.write("<layer name=rating top=0 left=0 visiblity=hide zIndex=1 bgcolor=#6699CC height=10></layer>");
		}
	window.onerror = null;
	bwidth = 0;
	if (document.all) swidth = document.all.waiting.clientWidth;
	else swidth = document.waiting.clip.width;
		if (arguments[0].search("http://")==0){
			if(document.all) document.write("<IFRAME STYLE=\"z-index:1\" WIDTH=468 HEIGHT=60 MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=no SRC="+arguments[0]+"&number="+nowTime.getTime()+bust+"><\/IFRAME>");
			else document.write("<SC"+"RIPT LANGUAGE=JavaScript SRC="+arguments[1]+"&number="+nowTime.getTime()+bust+"></SC"+"RIPT></td>");}
		else document.write(arguments[0]);
		if (arguments.length>4){
			}
		}
	else{
		if(document.all)document.onclick = CheckClick;
		else{
			window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
			window.onmousedown = CheckClick;}
	}
	eval(AutoRegStr);
}