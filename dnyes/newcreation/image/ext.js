  if (navigator.appName == "Netscape" && parseInt(navigator.appVersion) == 4) {
    // NS4
	var ns = true;
  }
  else if (document.all) {
    // IE4+
	var ie = true;
  }
  else if (document.getElementById) 
  //  else if (navigator.product == ("Gecko"))
//  else if (navigator.userAgent.indexOf("Gecko") != -1)
 {
	var ns6 = true;
  }

if (ns) document.writeln("<link rel=\"stylesheet\" type=\"text/css\" href=../1.files/intro.files//%22wdstylenn.css/%22>");
else if (ie) document.writeln("<link rel=\"stylesheet\" type=\"text/css\" href=../1.files/intro.files//%22wdstyleie.css/%22>"); 
else if (ns6)  document.writeln("<link rel=\"stylesheet\" type=\"text/css\" href=../1.files/intro.files//%22wdstylens6.css/%22>"); 
else document.writeln("<link rel=\"stylesheet\" type=\"text/css\" href=../1.files/intro.files//%22wdstylenn.css/%22>"); 

if (top==self) location.href = "start.asp"

function Act(imgName) {
		if (document.images) document[imgName].src = eval(imgName + '_on.src') 
}

function InAct(imgName) {
		if (document.images) document[imgName].src = eval(imgName + '_off.src') 
}

function imgon(imgname) {
        if (document.images) {
                document[imgname].src = eval(imgname + "on.src");
        }
}

function imgoff(imgname) {
        if (document.images) {
                document[imgname].src = eval(imgname + "off.src");
        }
}

function MM_reloadPage(init) {  
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);


