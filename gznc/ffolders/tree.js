function showhide(isend,outline,sign)
{
	if (outline.style.display=='none')
	{
		outline.style.display='';
		if (isend)
			sign.src = imgsrc_midminus;
		else
			sign.src = imgsrc_lastminus;
	}
	else
	{
		outline.style.display='none';
		if (isend)
			sign.src = imgsrc_midplus;
		else
			sign.src = imgsrc_lastplus;
	}
}

//	  Ù–‘£∫id, parentid, level, name, info, link, type
function fdrItem(index, id, parentid, level, name, info, type, haschild, is_begin_in_this_level, is_end_in_this_level)
{
	//  attr
	this.index = index;
	this.id = id;
	this.parentid = parentid;
	this.level = level;
	this.name = name;
	this.info = info;
	this.link = linkHeader + this.id;
	this.type = type;
	this.haschild = haschild;
	this.isbegin = is_begin_in_this_level;
	this.isend = is_end_in_this_level;
	this.dealed = false;
	this.opened = false;
	
	// methord
	// show span
	this.showspan = fdrItem_showSpan;
	this.showchildspan = fdrItem_showChildSpan;
}
//methord
function fdrItem_showSpan()
{
	if (this.dealed) return;

	this.dealed = true;
	var i=0;
	//alert(this.parentid)
	if (this.parentid != -1)
	{
		if (this.haschild)
		{
			spanstr += "<span id='menu" + this.id + "' style=' text-decoration:underline;'>";
			for (i=0; i<(this.level); i++)
				spanstr += "<img src='" + imgsrc_line + "' align=absMiddle>";
				spanstr += "</span>";
				spanstr += '<img src="' + imgsrc_midblk + '" align=absMiddle>';
				spanstr += "<a href='" + this.link + "' target='" + linkTarget + "'>"+ this.name + "</a>";
			spanstr += "<img id='menu" + this.id + "sign' src='";
			if (this.isend)
				spanstr += imgsrc_lastplus;
			else
				spanstr += imgsrc_midplus;
			spanstr += "' onClick='showhide(this.isend, menu" + this.id + "outline, menu" + this.id + "sign)' align=absMiddle style='cursor:hand;'><br>";		
			spanstr += "<span id='menu" + this.id + "outline' ";
			spanstr += 'style="display:';
			if (this.level > 1)
				spanstr += "'none'";
			spanstr += '">';
			this.showchildspan();
			spanstr += "</span>";
		}
		else
		{
			for (i=0; i<(this.level); i++)
			if (this.id==4)
				spanstr += '<img src="' + imgsrc_lastblk + '" align=absMiddle>';
			else
				spanstr += '<img src="' + imgsrc_line + '" align=absMiddle>';
			spanstr += '<img src="' + imgsrc_midblk + '" align=absMiddle>';
			spanstr += '<a href="' + this.link + '" target="' + linkTarget + '">' + this.name + '</a><br>';
		}
	}
	else
	{
			spanstr += "<span id='menu" + this.id + "' style=' text-decoration:underline;'>";
			for (i=0; i<(this.level); i++)
				spanstr += "</span>";
			spanstr += "<span id='menu" + this.id + "outline' ";
			spanstr += 'style="display:';
			if (this.level > 1)
				spanstr += "'none'";
			spanstr += '">';
			this.showchildspan();
			spanstr += "</span>";
	}
}
function InitFdrInfo()
{
	fdritems = fdritems1;
	var i=0;
	for (i=0; i<(fdritems.length); i++)
	{
		if (fdritems[i].id == 0)
		{
			fdritems[i].name=fdrname_root;
			fdritems[i].info=fdrname_root;
			fdritems[i].link=linkFdrRoot;
			break;
		}
			
	}
}
function fdrItem_showChildSpan()
{
	var i=0;
	for (i=0; i<(fdritems.length); i++)
	{
		if (fdritems[i].parentid == this.id)
			fdritems[i].showspan();
	}
}
function InitSetupInfo()
{
	fdritems = setupitems;
}

//	this block show upstair fdrItems.
function ShowFdrItems()
{
	var i=0;
	for (i=0; i<fdritems.length; i++)
	{
		if (!fdritems[i].dealed)
			fdritems[i].showspan();
	}
//	prompt("xxx", spanstr);
	document.write(spanstr);
}