document.onkeydown=look;
function look(){ 
//if(event.shiftKey) 
//alert("��ֹ��Shift��!"); //���Ի���ALT��CTRL
if(event.keyCode==116){
event.keyCode=0;
event.returnValue=false;
}
} 

function maxWin()
{
  try
  {
    var b = top.screenLeft == 0;
    var b = b && top.screen.availHeight - top.screenTop - top.body.offsetHeight - 20 == 0;
    if(!b)
    {
      var str  = '<object id=Max classid="clsid:ADB880A6-D8FF-11CF-9377-00AA003B7A11">'
          str += '<param name="Command" value="Maximize"></object>';
      document.body.insertAdjacentHTML("beforeEnd", str);
      document.getElementById("Max").Click();
    }
  }catch(e){}
}


function checkclick(msg){if(confirm(msg)){event.returnValue=true;}else{event.returnValue=false;}}


function menuShow(obj,maxh,obj2)
{
  if(obj.style.pixelHeight<maxh)
  {
    obj.style.pixelHeight+=maxh/20;
	obj.filters.alpha.opacity+=5;
	obj2.background="images/title_bg_show.gif";
    if(obj.style.pixelHeight==maxh/10)
	  obj.style.display='block';
	myObj=obj;
	myMaxh=maxh;
	myObj2=obj2;
	setTimeout('menuShow(myObj,myMaxh,myObj2)','5');
  }
}
function menuHide(obj,maxh,obj2)
{
  if(obj.style.pixelHeight>0)
  {
    if(obj.style.pixelHeight==maxh/20)
	  obj.style.display='none';
    obj.style.pixelHeight-=maxh/20;
	obj.filters.alpha.opacity-=5;
	obj2.background="images/title_bg_hide.gif";
	myObj=obj;
	myMaxh=maxh
	myObj2=obj2;
	setTimeout('menuHide(myObj,myMaxh,myObj2)','5');
  }
  else
    if(whichContinue)
	  whichContinue.click();
}
function menuChange(obj,maxh,obj2)
{
  if(obj.style.pixelHeight)
  {
    menuHide(obj,maxh,obj2);
	whichOpen='';
	whichcontinue='';
  }
  else
    if(whichOpen)
	{
	  whichContinue=obj2;
      whichOpen.click();
	}
	else
	{
	  menuShow(obj,maxh,obj2);
	  whichOpen=obj2;
	  whichContinue='';
	}
}


var tt='start';
var ii='start';
function turnit(ss,bb) {

  if (ss.style.display=="none") {
    if(tt!='start') tt.style.display="none";
    if(ii!='start') ii.src="dian.imgaes";
    ss.style.display="";
    tt=ss;
    ii=bb;
    bb.src="ball.imgaes";
  }
  else {
    ss.style.display="none"; 
    bb.src="dian.imgaes";
  }
}
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
function about()
{
window.showModalDialog("about.htm","","dialogWidth:420px;dialogHeight:340px;scroll:no;status:no;help:no");
}


function runClock() {
theTime = window.setTimeout("runClock()", 1000);
var today = new Date();
var display= today.toLocaleString();
status="�����ǣ�"+display+"����ӭ���Ĺ��٣�";
}


var version = "other"
browserName = navigator.appName;   
browserVer = parseInt(navigator.appVersion);

if (browserName == "Netscape" && browserVer >= 3) version = "n3";
else if (browserName == "Netscape" && browserVer < 3) version = "n2";
else if (browserName == "Microsoft Internet Explorer" && browserVer >= 4) version = "e4";
else if (browserName == "Microsoft Internet Explorer" && browserVer < 4) version = "e3";


function marquee1()
{
	if (version == "e4")
	{
		document.write("<marquee style='BOTTOM: 0px; FONT-WEIGHT: 100px; HEIGHT:20px;  TEXT-ALIGN: left; TOP: 3px' id='news' scrollamount='1' scrolldelay='10' behavior='loop' direction='left' border='0' onmouseover='this.stop()' onmouseout='this.start()'>")
	}
}

function marquee2()
{
	if (version == "e4")
	{
		document.write("</marquee>")
	}
}

function userreg_onsubmit() {
var i, n;
if (document.userreg.username.value=="")
	{
	  alert("�Բ��������������û�����")
	  document.userreg.username.focus()
	  return false
	 }
else if (document.userreg.username.value.length < 2)
	{
	  alert("�����û����ܲ��ܳ�һ�㣡")
	  document.userreg.username.focus()
	  return false
	 }
else if (document.userreg.username.value.length > 30)
	{
	  alert("�����û���̫���˰ɣ�")
	  document.userreg.username.focus()
	  return false
	 }
else if (document.userreg.username.value.indexOf('`') >= 0 ||
          document.userreg.username.value.indexOf('~') >= 0 ||
          document.userreg.username.value.indexOf('!') >= 0 ||
          document.userreg.username.value.indexOf('@') >= 0 ||
          document.userreg.username.value.indexOf('#') >= 0 ||
          document.userreg.username.value.indexOf('$') >= 0 ||
          document.userreg.username.value.indexOf('%') >= 0 ||
          document.userreg.username.value.indexOf('^') >= 0 ||
          document.userreg.username.value.indexOf('&') >= 0 ||
          document.userreg.username.value.indexOf('*') >= 0 ||
          document.userreg.username.value.indexOf('(') >= 0 ||
          document.userreg.username.value.indexOf(')') >= 0 ||
          document.userreg.username.value.indexOf('+') >= 0 ||
          document.userreg.username.value.indexOf('{') >= 0 ||
          document.userreg.username.value.indexOf('}') >= 0 ||
          document.userreg.username.value.indexOf('|') >= 0 ||
          document.userreg.username.value.indexOf('[') >= 0 ||
          document.userreg.username.value.indexOf(']') >= 0 ||
          document.userreg.username.value.indexOf('\\') >= 0 ||
          document.userreg.username.value.indexOf(';') >= 0 ||
          document.userreg.username.value.indexOf(':') >= 0 ||
          document.userreg.username.value.indexOf('>') >= 0 ||
          document.userreg.username.value.indexOf('<') >= 0 ||
          document.userreg.username.value.indexOf(',') >= 0 ||
          document.userreg.username.value.indexOf('?') >= 0 ||
          document.userreg.username.value.indexOf('/') >= 0 || 
          document.userreg.username.value.indexOf('\'') >= 0 || 
          document.userreg.username.value.indexOf('"') >= 0 || 
          document.userreg.username.value.indexOf(' ') >= 0 || 
          document.userreg.username.value.indexOf('=') >= 0 || 
          document.userreg.username.value.indexOf('%') >= 0
         )  
          {
            alert("�û����а�����Ч�ַ���������ѡ���û�����"); 
            document.userreg.username.focus();
            return false;
          }
else if (document.userreg.passwd.value=="")
	{
	  alert("�Բ��������������룡")
	  document.userreg.passwd.focus()
	  return false
	 }
else if (document.userreg.passwd.value.length < 4)
	{
	  alert("Ϊ�˰�ȫ����������Ӧ�ó�һ�㣡")
	  document.userreg.passwd.focus()
	  return false
	 }
else if (document.userreg.passwd.value.length > 16)
	{
	  alert("��������̫���˰ɣ�")
	  document.userreg.passwd.focus()
	  return false
	 }
else if (document.userreg.passwd.value.indexOf('"') >= 0 || 
          document.userreg.passwd.value.indexOf(' ') >= 0 || 
          document.userreg.passwd.value.indexOf('=') >= 0)  
          {
            alert("�����а�����Ч�ַ���������ѡ���û�����"); 
            document.userreg.passwd.focus();
            return false;
          }
else if (document.userreg.username.value==document.userreg.passwd.value)
	{
	  alert("Ϊ�˰�ȫ���û��������벻Ӧ����ͬ��")
	  document.userreg.passwd.focus()
	  return false
	 }
else if (document.userreg.passwd2.value=="")
	{
	  alert("�Բ�������������֤���룡")
	  document.userreg.passwd2.focus()
	  return false
	 }
else if (document.userreg.passwd2.value !== document.userreg.passwd.value)
	{
	  alert("�Բ�����������������벻һ�£�")
	  document.userreg.passwd2.focus()
	  return false
	 }
else if (document.userreg.question.value=="")
	{
	  alert("�Բ�������������ʾ���⣡")
	  document.userreg.question.focus()
	  return false
	 }
else if (document.userreg.answer.value=="")
	{
	  alert("�Բ���������������𰸣�")
	  document.userreg.answer.focus()
	  return false
	 }
else if (document.userreg.question.value==document.userreg.answer.value)
	{
	  alert("Ϊ�˰�ȫ����ʾ����������𰸲�Ӧ����ͬ��")
	  document.userreg.answer.focus()
	  return false
	 }
else if (document.userreg.fullname.value=="")
	{
	  alert("�Բ���������������ʵ������")
	  document.userreg.fullname.focus()
	  return false
	 }
else if (document.userreg.depid.value=="")
	{
	  alert("�Բ�����ѡ�����Ĺ�����λ��")
	  document.userreg.depid.focus()
	  return false
	 }
else if (document.userreg.sex.value=="")
	{
	  alert("�Բ�����ѡ�������Ա�")
	  document.userreg.sex.focus()
	  return false
	 }
else if (document.userreg.tel.value=="")
	{
	  alert("�Բ���������������ϵ�绰��")
	  document.userreg.tel.focus()
	  return false
	 }
else if (document.userreg.email.value=="")
	{
	  alert("�Բ������������ĵ����ʼ���")
	  document.userreg.email.focus()
	  return false
	 }
else if (document.userreg.email.value.indexOf("@",0)== -1||document.userreg.email.value.indexOf(".",0)==-1)
	  {
	  alert("�Բ���������ĵ����ʼ�����")
	  document.userreg.email.focus()
	  return false
	 }
}

function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}


//Function to format text in the text box
function FormatText(command, option){
	
  	frames.message.document.execCommand(command, true, option);
  	frames.message.focus();
}

//Function to clear form
function ResetForm(){

	if (window.confirm('ȷ��Ҫ��նԻ�������?')){
	 	frames.message.document.body.innerHTML = ''; 
	 	return true;
	 } 
	 return false;		
}

//Function to open pop up window
function openWin(theURL,winName,features) {
  	window.open(theURL,winName,features);
}

function setMode(newMode)
{
  bTextMode = newMode;
  var cont;
  if (bTextMode) {
    cleanHtml();
    cleanHtml();

    cont=message.document.body.innerHTML;
    message.document.body.innerText=cont;
  } else {
    cont=message.document.body.innerText;
    message.document.body.innerHTML=cont;
  }
  message.focus();
}

function cleanHtml()
{
  var fonts = message.document.body.all.tags("FONT");
  var curr;
  for (var i = fonts.length - 1; i >= 0; i--) {
    curr = fonts[i];
    if (curr.style.backgroundColor == "#ffffff") curr.outerHTML	= curr.innerHTML;
  }

}