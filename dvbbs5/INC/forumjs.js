function openScript(url, width, height){
	var Win = window.open(url,"openScript",'width=' + width + ',height=' + height + ',resizable=1,scrollbars=yes,menubar=no,status=yes' );
}

function gotourl(gourl,mtitle,mwidth,mheight)
{
if(window.parent.mmain){
window.parent.mmain.document.all.move1.innerHTML=mtitle;
window.parent.document.all.mmain.style.width=mwidth;
window.parent.document.all.mmain.style.height=mheight;
window.parent.mmain.wmain1.location.replace(gourl);
window.parent.document.all.mmain.style.left=((parseInt(window.document.body.clientWidth)-parseInt(window.parent.document.all.mmain.style.width))/2);
window.parent.document.all.mmain.style.top=((parseInt(window.document.body.clientHeight)-parseInt(window.parent.document.all.mmain.style.height))/2);

if(parseInt(window.parent.document.all.mmain.style.width)>=parseInt(window.document.body.clientWidth)){
window.parent.document.all.mmain.style.width=parseInt(window.document.body.clientWidth)-10;
window.parent.document.all.mmain.style.left=5;
}

if(parseInt(window.parent.document.all.mmain.style.height)>=parseInt(window.document.body.clientHeight)){
window.parent.document.all.mmain.style.height=parseInt(window.document.body.clientHeight)-10;
window.parent.document.all.mmain.style.top=5;
}
window.parent.mmain.document.all.maxfont.innerHTML='1'
window.parent.document.all.mmain.style.display='';
}

}