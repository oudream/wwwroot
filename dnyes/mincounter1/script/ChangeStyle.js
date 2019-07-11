// Example:
// alert( readCookie("myCookie") );
function readCookie(name)
{
  var cookieValue = "";
  var search = name + "=";
  if(document.cookie.length > 0)
  { 
    offset = document.cookie.indexOf(search);
    if (offset != -1)
    { 
      offset += search.length;
      end = document.cookie.indexOf(";", offset);
      if (end == -1) end = document.cookie.length;
      cookieValue = unescape(document.cookie.substring(offset, end))
    }
  }
  return cookieValue;
}
// Example:
// writeCookie("myCookie", "my name", 24);
// Stores the string "my name" in the cookie "myCookie" which expires after 24 hours.

function writeCookie(name, value)
{
  var expire = "";
    expire = new Date((new Date()).getTime() + 24 * 365 * 3600000);
    expire = "; expires=" + expire.toGMTString();
  document.cookie = name + "=" + escape(value) + expire;
}

function changeStyle(n){
	for(var i=1;i<document.styleSheets.length;++i){
			if(document.styleSheets[i].href)
				document.styleSheets[i].disabled=true;
		}
	/*
	var oCss = document.createElement("LINK");
	oCss.rel="stylesheet";
	oCss.type="text/css";
	oCss.href=(n?"../style/Sty0"+n+".css":"../style/main.css");
	document.body.appendChild(oCss)
	*/
	try{
		document.createStyleSheet(n?"../style/Sty0"+n+".css":"../style/main.css");
	}catch(e){
		//not ie?
	}
	writeCookie("cc_count_6_style",n);
}

var sCookiesStyle = readCookie("cc_count_6_style");
if(!parseInt(sCookiesStyle)) sCookiesStyle=1; 
changeStyle(sCookiesStyle);