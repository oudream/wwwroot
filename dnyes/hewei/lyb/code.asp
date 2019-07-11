<%

Function icode2Html(str, showimg, nonewwindow)
ON ERROR RESUME NEXT
if not str<>"" then exit function
  tmpstr="icode"
  str=icodeStr(str,"url",nonewwindow)
  str=icodeStr(str,"email",nonewwindow)
  if showimg then
    str=icodeStr(str,"img",nonewwindow)
  end if
  str=replace(str,"[b]","<b>",1,-1,1)
  str=replace(str,"[/b]","</b>",1,-1,1)
  str=replace(str,"[br]","<br>",1,-1,1)
  str=replace(str,"["&tmpstr,"[",1,-1,1)
  str=replace(str,tmpstr&"]","]",1,-1,1)
  str=replace(str,"/"&tmpstr,"/",1,-1,1)
  icode2Html=str
End Function

function icodestr(icode_str,icodeKeyWord,nonewwindow)
ON ERROR RESUME NEXT
tmpstr="icode"
beginstr=1
endstr=1
do while icodeKeyWord="url" or icodeKeyWord="email"
  beginstr=instr(beginstr,icode_str,"["&icodeKeyWord&"=",1)
  if beginstr=0 then exit do
  endstr=instr(beginstr,icode_str,"]",1)
  if endstr=0 then exit do
  LicodeKeyWord=icodeKeyWord
  beginstr=beginstr+len(licodeKeyWord)+2
  text=mid(icode_str,beginstr,endstr-beginstr)
  codetext=replace(text,"[","["&tmpstr,1,-1,1)
  codetext=replace(codetext,"]",tmpstr&"]",1,-1,1)
  codetext=replace(codetext,"/","/"&tmpstr,1,-1,1)
  select case icodeKeyWord
    case "url"
	icode_str=replace(icode_str,"[url="&text&"]","<a href='"&codetext&"' target=_blank>",1,1,1)
	icode_str=replace(icode_str,"[/url]","</a>",1,1,1)
	case "email"
	icode_str=replace(icode_str,"[email="&text&"]","<a href='mailto:"&codetext&"'>",1,1,1)
	icode_str=replace(icode_str,"[/email]","</a>",1,1,1)
  end select
loop

beginstr=1
do
  beginstr=instr(beginstr,icode_str,"["&icodeKeyWord&"]",1)
  if beginstr=0 then exit do
  endstr=instr(beginstr,icode_str,"[/"&icodeKeyWord&"]",1)
  if endstr=0 then exit do
  LicodeKeyWord=icodeKeyWord
  beginstr=beginstr+len(licodeKeyWord)+2
  text=mid(icode_str,beginstr,endstr-beginstr)
  codetext=replace(text,"[","["&tmpstr,1,-1,1)
  codetext=replace(codetext,"]",tmpstr&"]",1,-1,1)
  codetext=replace(codetext,"/","/"&tmpstr,1,-1,1)
  select case icodeKeyWord
    case "url"
	icode_str=replace(icode_str,"["&icodeKeyWord&"]"&text,"<a href='"&codetext&"' target=_blank>"&codetext,1,1,1)
	icode_str=replace(icode_str,"<a href='"&codetext&"' target=_blank>"&codetext&"[/"&icodeKeyWord&"]","<a href="&codetext&" target=_blank>"&codetext&"</a>",1,1,1)
    case "email"
	icode_str=replace(icode_str,"["&icodeKeyWord&"]"&text,"<a href='mailto:"&codetext&"'>"&codetext,1,1,1)
	icode_str=replace(icode_str,"<a href='mailto:"&codetext&"'>"&codetext&"[/"&icodeKeyWord&"]","<a href=mailto:"&codetext&" target=_blank>"&codetext&"</a>",1,1,1)
    case "img"
	if nonewwindow then
	  icode_str=replace(icode_str,"[img]"&text,"<img src="&codetext,1,1,1)
	  icode_str=replace(icode_str,"[/img]"," border=0>",1,1,1)
	else
	  icode_str=replace(icode_str,"[img]"&text,"<table width='100%' align=center border='0' cellspacing='0' cellpadding='0' style='TABLE-LAYOUT: fixed'><tr><td><a href='"&codetext&"' target=_blank><img src="&codetext,1,1,1)
	  icode_str=replace(icode_str,"[/img]"," border=0 alt='点击打开新窗口'></a></td></tr></table>",1,1,1)
	end if
  end select
loop
icodestr=icode_str
end function
%>