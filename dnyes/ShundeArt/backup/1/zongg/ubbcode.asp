
<%
function doCode(fString, fOTag, fCTag, fROTag, fRCTag)
	fOTagPos = Instr(1, fString, fOTag, 1)
	fCTagPos = Instr(1, fString, fCTag, 1)
	while (fCTagPos > 0 and fOTagPos > 0)
		fString = replace(fString, fOTag, fROTag, 1, 1, 1)
		fString = replace(fString, fCTag, fRCTag, 1, 1, 1)
		fOTagPos = Instr(1, fString, fOTag, 1)
		fCTagPos = Instr(1, fString, fCTag, 1)
	wend
	doCode = fString
end function

function HTMLEncode(fString)

	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")

	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode = fString
end function

function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function

function UBBCode(strContent)
	on error resume next
	strContent = HTMLEncode(strContent)
	dim objRegExp
	Set objRegExp=new RegExp
	objRegExp.IgnoreCase =true
	objRegExp.Global=True

   
	objRegExp.Pattern="(\[color=(.*)\])(.*)(\[\/color\])"
	strContent=objRegExp.Replace(strContent,"<font color=$2>$3</font>")
	objRegExp.Pattern="(\[face=(.*)\])(.*)(\[\/face\])"
	strContent=objRegExp.Replace(strContent,"<font face=$2>$3</font>")
	objRegExp.Pattern="(\[align=(.*)\])(.*)(\[\/align\])"
	strContent=objRegExp.Replace(strContent,"<div align=$2>$3</div>")

	objRegExp.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
	strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=1 face=""Verdana, Arial"">quote:</font><HR>$2<HR></BLOCKQUOTE>")

    
	objRegExp.Pattern="(\[i\])(.*)(\[\/i\])"
	strContent=objRegExp.Replace(strContent,"<i>$2</i>")
	objRegExp.Pattern="(\[u\])(.*)(\[\/u\])"
	strContent=objRegExp.Replace(strContent,"<u>$2</u>")
	objRegExp.Pattern="(\[b\])(.*)(\[\/b\])"
	strContent=objRegExp.Replace(strContent,"<b>$2</b>")


	objRegExp.Pattern="(\[size=1\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=1>$2</font>")
	objRegExp.Pattern="(\[size=2\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=2>$2</font>")
	objRegExp.Pattern="(\[size=3\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=3>$2</font>")
	objRegExp.Pattern="(\[size=4\])(.*)(\[\/size\])"
	strContent=objRegExp.Replace(strContent,"<font size=4>$2</font>")

	strContent = doCode(strContent, "[list]", "[/list]", "<ul>", "</ul>")
	strContent = doCode(strContent, "[list=1]", "[/list]", "<ol type=1>", "</ol id=1>")
	strContent = doCode(strContent, "[list=a]", "[/list]", "<ol type=a>", "</ol id=a>")
	strContent = doCode(strContent, "[*]", "[/*]", "<li>", "</li>")
	strContent = doCode(strContent, "[code]", "[/code]", "<pre id=code><font size=1 face=""Verdana, Arial"" id=code>", "</font id=code></pre id=code>")

	set objRegExp=Nothing
	UBBCode=strContent
end function


%>