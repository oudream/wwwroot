<%
function DvBCode(strContent,PostUserGroup)
if Cint(Forum_Setting(35)) <> 1 then
	strContent = dvHTMLEncode(strContent)
else
	strContent = HTMLcode(strContent)
end if
if Cint(Forum_Setting(34))<>1 and PostUserGroup>2 then
	DvBCode=strContent
	exit function
end if
dim re
dim po,ii
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
po=0
ii=0

re.Pattern="(javascript)"
strContent=re.Replace(strContent,"&#106avascript")
re.Pattern="(jscript:)"
strContent=re.Replace(strContent,"&#106script:")
re.Pattern="(js:)"
strContent=re.Replace(strContent,"&#106s:")
re.Pattern="(value)"
strContent=re.Replace(strContent,"&#118alue")
re.Pattern="(about:)"
strContent=re.Replace(strContent,"about&#58")
re.Pattern="(file:)"
strContent=re.Replace(strContent,"file&#58")
re.Pattern="(document.cookie)"
strContent=re.Replace(strContent,"documents&#46cookie")
re.Pattern="(vbscript:)"
strContent=re.Replace(strContent,"&#118bscript:")
re.Pattern="(vbs:)"
strContent=re.Replace(strContent,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
strContent=re.Replace(strContent,"&#111n$2")
	
if Cint(Forum_Setting(36)) = 1 or PostUserGroup<4 then
	re.Pattern="\[IMG\](http|https|ftp):\/\/(.[^\[]*)\[\/IMG\]"
	strContent=re.Replace(strContent,"<a onfocus=this.blur() href=""$1://$2"" target=_blank><IMG SRC=""$1://$2"" border=0 alt=�������´������ͼƬ onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>")
	re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
	strContent= re.Replace(strContent,"<br><IMG SRC="""&Forum_info(7)&"$1.gif"" border=0>���������ͼƬ���£�<br><A HREF=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=�������´������ͼƬ onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
end if

re.Pattern="\[UPLOAD=(.[^\[]*)\](.[^\[]*)\[\/UPLOAD\]"
strContent= re.Replace(strContent,"<br><IMG SRC="""&Forum_info(7)&"$1.gif"" border=0> <a href=""$2"">���������ļ�</a>")

if cint(Forum_Setting(67))=1 or PostUserGroup<4 then
	re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
	strContent=re.Replace(strContent,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
	re.Pattern="\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
	strContent=re.Replace(strContent,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
	strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
	strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")

	re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$2"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")

	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
end if

'���ֿɼ�
re.Pattern="(^.*)(\[point=*([0-9]*)\])(.[^\[]*)(\[\/point\])(.*)"
po=re.Replace(strContent,"$3")
if IsNumeric(po) then
	ii=int(po)
else
	ii=0
end if
if membername=username or myuserep>ii then
	strContent=re.Replace(strContent,"$1<hr noshade size=1>����Ϊ��Ҫ���ִﵽ<B>$3</B>�������������<BR>$4<hr noshade size=1>$6")
else
	strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">������Ҫ����ִﵽ$3���ϲſ������</font><hr noshade size=1>$6")
end if
'�����ɼ�
re.Pattern="(^.*)(\[UserCP=*([0-9]*)\])(.[^\[]*)(\[\/UserCP\])(.*)"
po=re.Replace(strContent,"$3")
if IsNumeric(po) then
	ii=int(po)
else
	ii=0
end if
if membername=username or myusercp>ii then
	strContent=re.Replace(strContent,"$1<hr noshade size=1>����Ϊ��Ҫ�����ﵽ<B>$3</B>�������������<BR>$4<hr noshade size=1>$6")
else
	strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">������Ҫ�������ﵽ$3���ϲſ������</font><hr noshade size=1>$6")
end if
'�����ɼ�
re.Pattern="(^.*)(\[Power=*([0-9]*)\])(.[^\[]*)(\[\/Power\])(.*)"
po=re.Replace(strContent,"$3")
if IsNumeric(po) then
	ii=int(po) 
else
	ii=0
end if
if membername=username or mypower>ii then
	strContent=re.Replace(strContent,"$1<hr noshade size=1>����Ϊ��Ҫ�����ﵽ<B>$3</B>�������������<BR>$4<hr noshade size=1>$6")
else
	strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">������Ҫ�������ﵽ$3���ϲſ������</font><hr noshade size=1>$6")
end if
'���¿ɼ�
re.Pattern="(^.*)(\[Post=*([0-9]*)\])(.[^\[]*)(\[\/Post\])(.*)"
po=re.Replace(strContent,"$3")
if IsNumeric(po) then
	ii=int(po) 
else
	ii=0
end if
if membername=username or myArticle>ii then
	strContent=re.Replace(strContent,"$1<hr noshade size=1>����Ϊ��Ҫ�������ﵽ<B>$3</B>�������������<BR>$4<hr noshade size=1>$6")
else
	strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">������Ҫ���������ﵽ$3���ϲſ������</font><hr noshade size=1>$6")
end if
'��Ǯ�ɼ�
re.Pattern="(^.*)(\[money=*([0-9]*)\])(.[^\[]*)(\[\/money\])(.*)"
po=re.Replace(strContent,"$3")
if IsNumeric(po) then
	ii=int(po) 
else
	ii=0
end if
if membername=username or mymoney>ii then
	strContent=re.Replace(strContent,"$1<hr noshade size=1>����Ϊ��Ҫ��Ǯ���ﵽ<B>$3</B>�������������<BR>$4<hr noshade size=1>$6")
else
	strContent=re.Replace(strContent,"$1<hr noshade size=1><font color="&Forum_body(8)&">������Ҫ���Ǯ���ﵽ$3���ϲſ������</font><hr noshade size=1>$6")
end if
'�ظ��ɼ�
if instr(lcase(strContent),"[replyview]")>0 and instr(lcase(strContent),"[/replyview]")>0 then
	dim vrs
	set vrs=conn.execute("select AnnounceID from "&TotalUseTable&" where rootid="&Announceid&" and PostUserID="&UserID)
	re.Pattern="(\[replyview\])(.[^\[]*)(\[\/replyview\])"
	if vrs.eof and vrs.bof then
	strContent=re.Replace(strContent,"<hr noshade size=1><font color="&Forum_body(8)&">��������Ҫ�ظ��������</font><hr noshade size=1>")
	else
	strContent=re.Replace(strContent,"<hr noshade size=1>��������Ϊ��Ҫ�ظ��������<BR>$2<hr noshade size=1>")
	end if
	set vrs=nothing
end if

re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"">$2</A>")
re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

'�Զ�ʶ����ַ
re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")

'�Զ�ʶ��www�ȿ�ͷ����ַ
re.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=http://$2>$2</a>")

'�Զ�ʶ��Email��ַ����򿪱�������������ݺܶ�����ӻ����������ͣ��
're.Pattern = "([^(=)])((\w)+[@]{1}((\w)+[.]){1,3}(\w)+)"
'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=""mailto:$2"">$2</a>")

if Forum_Setting(37) = "1" or PostUserGroup<4 then
	re.Pattern="\[em(.[^\[]*)\]"
	strContent=re.Replace(strContent,"<img src="&Forum_info(7)&"em$1.gif border=0 align=middle>")
end if

re.Pattern="\[HTML\](.[^\[]*)\[\/HTML\]"
strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' class='"&abgcolor&"'><td><b>��������Ϊ�������:</b><br>$1</td></table>")
re.Pattern="\[code\](.[^\[]*)\[\/code\]"
strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' class='"&abgcolor&"'><td><b>��������Ϊ�������:</b><br>$1</td></table>")

re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
strContent=re.Replace(strContent,"<font color=$1>$2</font>")
re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
strContent=re.Replace(strContent,"<font face=$1>$2</font>")
re.Pattern="\[align=(center|left|right)\](.*)\[\/align\]"
strContent=re.Replace(strContent,"<div align=$1>$2</div>")

re.Pattern="\[QUOTE\](.*)\[\/QUOTE\]"
strContent=re.Replace(strContent,"<table style=""width:80%"" cellpadding=5 cellspacing=1 class=tableborder1><TR><TD class="&abgcolor&">$1</td></tr></table><br>")
re.Pattern="\[fly\](.*)\[\/fly\]"
strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
re.Pattern="\[move\](.*)\[\/move\]"
strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$1</marquee>")	
re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")

re.Pattern="\[i\](.[^\[]*)\[\/i\]"
strContent=re.Replace(strContent,"<i>$1</i>")
re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
strContent=re.Replace(strContent,"<u>$1</u>")
re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
strContent=re.Replace(strContent,"<b>$1</b>")
re.Pattern="\[size=([1-4])\](.[^\[]*)\[\/size\]"
strContent=re.Replace(strContent,"<font size=$1>$2</font>")
strContent=replace(strContent,"<I></I>","")
set re=Nothing
DvBcode=strContent
end function

function DvSignCode(strContent,PostUserGroup)
if Forum_Setting(66)<>1 then
	strContent = dvHTMLEncode(strContent)
else
	strContent = HTMLcode(strContent)
end if
if Forum_Setting(65)<>1 and PostUserGroup>2 then
	DvSignCode=strContent
	exit function
end if
dim re
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(javascript)"
strContent=re.Replace(strContent,"&#106avascript")
re.Pattern="(jscript:)"
strContent=re.Replace(strContent,"&#106script:")
re.Pattern="(js:)"
strContent=re.Replace(strContent,"&#106s:")
re.Pattern="(value)"
strContent=re.Replace(strContent,"&#118alue")
re.Pattern="(about:)"
strContent=re.Replace(strContent,"about&#58")
re.Pattern="(file:)"
strContent=re.Replace(strContent,"file&#58")
re.Pattern="(document.cookie)"
strContent=re.Replace(strContent,"documents&#46cookie")
re.Pattern="(vbscript:)"
strContent=re.Replace(strContent,"&#118bscript:")
re.Pattern="(vbs:)"
strContent=re.Replace(strContent,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
strContent=re.Replace(strContent,"&#111n$2")

if Forum_Setting(67)=1 or PostUserGroup<4 then
	re.Pattern="\[IMG\](http|https|ftp):\/\/(.[^\[]*)\[\/IMG\]"
	strContent=re.Replace(strContent,"<IMG SRC=""$1://$2"" border=0 onload=""javascript:if(this.width>screen.width-500)this.width=screen.width-500"">")
end if

if Forum_Setting(67)=1 or PostUserGroup<4 then
	re.Pattern="(\[FLASH\])(http://.[^\[]*(.swf))(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$2"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")

	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(http://.[^\[]*(.swf))(\[\/FLASH\])"
	strContent= re.Replace(strContent,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
end if

re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")

re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"">$2</A>")
re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
strContent= re.Replace(strContent,"<img align=absmiddle src=pic/email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

'�Զ�ʶ����ַ
re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")

'�Զ�ʶ��www�ȿ�ͷ����ַ
re.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=http://$2>$2</a>")

'�Զ�ʶ��Email��ַ
're.Pattern = "([^(=)])((\w)+[@]{1}((\w)+[.]){1,3}(\w)+)"
'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=""mailto:$2"">$2</a>")

if Forum_Setting(37) = "1" or PostUserGroup<4 then
	re.Pattern="\[em(.[^\[]*)\]"
	strContent=re.Replace(strContent,"<img src="&Forum_info(7)&"em$1.gif border=0 align=middle>")
end if

re.Pattern="\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
strContent=re.Replace(strContent,"<font color=$1>$2</font>")
re.Pattern="\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
strContent=re.Replace(strContent,"<font face=$1>$2</font>")
re.Pattern="\[align=(center|left|right)\](.[^\[]*)\[\/align\]"
strContent=re.Replace(strContent,"<div align=$1>$2</div>")

re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")

re.Pattern="\[i\](.[^\[]*)\[\/i\]"
strContent=re.Replace(strContent,"<i>$1</i>")
re.Pattern="\[u\](.[^\[]*)(\[\/u\])"
strContent=re.Replace(strContent,"<u>$1</u>")
re.Pattern="\[b\](.[^\[]*)(\[\/b\])"
strContent=re.Replace(strContent,"<b>$1</b>")
re.Pattern="\[size=([1-4])\](.[^\[]*)\[\/size\]"
strContent=re.Replace(strContent,"<font size=$1>$2</font>")
strContent=replace(strContent,"<I></I>","")
set re=Nothing
DvSignCode=strContent
end function

function reUBBCode(strContent)
strContent = dvHTMLEncode(strContent)
dim re
Set re=new RegExp
re.IgnoreCase =true
re.Global=True

re.Pattern="(\[QUOTE\])(.[^\[]*)(\[\/QUOTE\])"
strContent=re.Replace(strContent,"")
re.Pattern="(\[point=*([0-9]*)\])(.*)(\[\/point\])"
strContent=re.Replace(strContent,"")
re.Pattern="(\[post=*([0-9]*)\])(.*)(\[\/post\])"
strContent=re.Replace(strContent,"")
re.Pattern="(\[power=*([0-9]*)\])(.*)(\[\/power\])"
strContent=re.Replace(strContent,"")
re.Pattern="(\[usercp=*([0-9]*)\])(.*)(\[\/usercp\])"
strContent=re.Replace(strContent,"")
re.Pattern="(\[replyview=*([0-9]*)\])(.*)(\[\/replyview\])"
strContent=re.Replace(strContent,"")
strContent=replace(strContent,"<I></I>","")

set re=Nothing
reUBBCode=strContent
end function

Function FilterJS(v)
if not isnull(v) then
dim t
dim re
dim reContent
Set re=new RegExp
re.IgnoreCase =true
re.Global=True
re.Pattern="(javascript)"
t=re.Replace(v,"&#106avascript")
re.Pattern="(jscript:)"
t=re.Replace(t,"&#106script:")
re.Pattern="(js:)"
t=re.Replace(t,"&#106s:")
re.Pattern="(value)"
t=re.Replace(t,"&#118alue")
re.Pattern="(about:)"
t=re.Replace(t,"about&#58")
re.Pattern="(file:)"
t=re.Replace(t,"file&#58")
re.Pattern="(document.cookie)"
t=re.Replace(t,"documents&#46cookie")
re.Pattern="(vbscript:)"
t=re.Replace(t,"&#118bscript:")
re.Pattern="(vbs:)"
t=re.Replace(t,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
t=re.Replace(t,"&#111n$2")
re.Pattern="(&#)"
t=re.Replace(t,"��#")
FilterJS=t
set re=nothing
end if
End Function

function dvHTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "<I></I>&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    dvHTMLEncode = fString
end if
end function
%>
