//时间控制
var now = new Date();
var months = now.getMonth()+1;
var dates = now.getDate();
var hours = now.getHours();
var minutes = now.getMinutes();
var seconds = now.getSeconds();

var days = 100;
var expdate = new Date();
expdate.setTime (expdate.getTime() + (86400 * 1000 * days));

function getCookieVal (offset) {
        var endstr = document.cookie.indexOf (";", offset);
        if (endstr == -1)
                endstr = document.cookie.length;
        return unescape(document.cookie.substring(offset, endstr));
}

function GetCookie (name) {
        var arg = name + "=";
        var alen = arg.length;
        var clen = document.cookie.length;
        var i = 0;
        
        while (i < clen) {
                var j = i + alen;
                if (document.cookie.substring(i, j) == arg)
                return getCookieVal (j);
                i = document.cookie.indexOf(" ", i) + 1;
                if (i == 0) 
                break; 
        }
        return -1;
}

function SetCookie (name, value) {

        var argv = SetCookie.arguments;
        var argc = SetCookie.arguments.length;
        var expires = (2 < argc) ? argv[2] : null;
        var path = (3 < argc) ? argv[3] : null;
        var domain = (4 < argc) ? argv[4] : null;
        var secure = (5 < argc) ? argv[5] : false;
        
        document.cookie = name + "=" + escape (value) +
                ((expires == null) ? "" : ("; expires=" + expires.toGMTString())) +
                ((path == null) ? "" : ("; path=" + path)) +
                ((domain == null) ? "" : ("; domain=" + domain)) +
                ((secure == true) ? "; secure" : "");
}

function SaveChangeStyle(frm )
{
	var style = frm.style;
	var i;
	for( i=0; i < style.options.length; i++ ){
		if ( true == style.options[i].selected )
		{
			SetCookie( "style", style.options[i].value, expdate, "", "163.net" );
			SetCookie( "style", style.options[i].value, expdate, "", "mail.tom.com" );
		}
	}
	return true;
}

function SaveStyle(frm )
{
        var style = frm.style;
        var user = frm.user;
        var i;
        for( i=0; i < style.options.length; i++ ){
                if ( true == style.options[i].selected )
                {
                        SetCookie( "style", style.options[i].value, expdate, "", "163.net" );
                        SetCookie( "style", style.options[i].value, expdate, "", "mail.tom.com" );
                        SetCookie( "user", user.value, expdate, "", "163.net" );
                        SetCookie( "user", user.value, expdate, "", "mail.tom.com" );
                }
        }
        return true;
}

function LoadStyle()
{
        var i;
        var styleIndex = GetCookie( "style" );
        var username = GetCookie( "user" );

try{        
        if ( styleIndex > 1 || styleIndex == -1 ){
           loginfrm.type[0].click();
        }else{
           loginfrm.type[1].click();
        }
}catch(e){}

        if ( null != styleIndex ){
                for( i=0; i < loginfrm.style.options.length; i++ ){
                        if ( styleIndex == loginfrm.style.options[i].value )
                        {
                                loginfrm.style.options[i].selected = true;
                        }
                }
        }

        loginfrm.user.focus();

        if ( -1 != username ){
            loginfrm.user.value = username;
            loginfrm.pass.focus();
        }
        
        
}

function update_subtype(type,subTypeForm)
{
    subTypeForm.length = 0;
    var styleIndex = GetCookie( "style" );
    
    switch(type)
    {
    case "0":
        subTypeForm.options[0] = new Option("默认风格","0");
        subTypeForm.options[1] = new Option("中国古典","1");
        subTypeForm.options[0].selected = true;
        break;
    case "1":
        subTypeForm.options[0] = new Option("纵横本色","2");
        subTypeForm.options[1] = new Option("古典神话","3");
        subTypeForm.options[2] = new Option("潮流科技","4");
        subTypeForm.options[3] = new Option("少女情怀","5");
    //  subTypeForm.options[4] = new Option("新新人类","6");
        subTypeForm.options[0].selected = true;
        break;
    }
    
    if ( null != styleIndex ){
        for( i=0; i < subTypeForm.options.length; i++ ){
           if ( styleIndex == subTypeForm.options[i].value )
           {
                subTypeForm.options[i].selected = true;
           }
        }
    }
    
}

var bkgs = new Array;
bkgs[0] = "http://bjpic.163.net/images/style_cm25_company_tom/folder_bg.gif";

function SetLeftBackGround()
{
    var background;
    
    try{
        background = GetCookie( "left_bkground" );
        if ( -1 == background ){
            background = bkgs[0];
        }
        htmlbody.background = background;
    }
    catch(e){
    }
}

function SetRightBackGround()
{
    var background;
    
    try{
        background = GetCookie( "right_bkground" );
        if ( -1 == background ){
            background = bkgs[0];
        }
        htmlbody.background = background;
    }
    catch(e){
    }
}

function ckma( cked, addrs )
{
     var xn=0; var p_=0;
     p_=addrs.indexOf("@",p_);  while(p_>0)
     {xn++; p_=addrs.indexOf("@",p_+1);}
     if (xn>1){ alert("你输入的邮件地址格式不对"); return false; }
     p_=addres.indexOf(" ");
     if (p_>0){ alert("你输入的邮件地址格式不对"); return false; }
     p_=addres.indexOf(",");
     if (p_>0){ alert("你输入的邮件地址格式不对"); return false; }
     p_=addres.indexOf(";");
     if (p_>0){ alert("你输入的邮件地址格式不对"); return false; }
     return true;
}
