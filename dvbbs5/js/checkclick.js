function CheckClick(e)
{
if (e == 1)
{
	if (bwidth<swidth*0.98){
	bwidth += (swidth - bwidth) * 0.025;
	if (document.all)document.sbar.width = bwidth;
	else document.rating.clip.width = bwidth;
	setTimeout('CheckClick(1);',150);}
}
else
{
	if(document.all)
	{
		if(document.all.waiting.style.visibility == 'visible')
			{document.all.waiting.style.visibility = 'hidden';
			bwidth = 1;}
		whichIt = event.srcElement;
		while (whichIt.tagName != "A") {
			whichIt = whichIt.parentElement;
			if (whichIt == null)return true;
		}
		if(whichIt.href.substring(0,5)=="http:"){
		document.all.waiting.style.pixelTop = (document.body.offsetHeight - document.all.waiting.clientHeight) / 2 + document.body.scrollTop;
		document.all.waiting.style.pixelLeft = (document.body.offsetWidth - document.all.waiting.clientWidth) / 2 + document.body.scrollLeft;
		document.all.waiting.style.visibility = 'visible';
		if(!bwidth)CheckClick(1);
		bwidth = 1;}
	}
	else
	{
	if(document.waiting.visibility == 'show')
			{document.waiting.visibility = 'hide';
			document.rating.visibility = 'hide';
			bwidth = 1;}
		if(e.target.href.toString() != '')
		{
			document.waiting.top = (window.innerHeight - document.waiting.clip.height) / 2 + self.pageYOffset;
			document.waiting.left = (window.innerWidth - document.waiting.clip.width) / 2 + self.pageXOffset;
			document.waiting.visibility = 'show';
			document.rating.top = (window.innerHeight - document.waiting.clip.height) / 2 + self.pageYOffset+document.waiting.clip.height-10;
			document.rating.left = (window.innerWidth - document.waiting.clip.width) / 2 + self.pageXOffset;
			document.rating.visibility = 'show';
			if(!bwidth)CheckClick(1);
			bwidth = 1;
		}
	}
	return true;
}
}