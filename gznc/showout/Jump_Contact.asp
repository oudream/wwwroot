<html>
<head>
<title>Jump...</title>
<LINK href="CSS.css" type=text/css rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<FORM name=loading> 
<DIV align=center> 
<P> 
<INPUT 
style="font-weight: bolder; color: #FF5A00; font-family: Arial; background-color: #4C5155; border-style: solid; border-color: #4C5155; 

padding: 0px" 
size=46 name=chart> 
<BR> 
      <input 
style="color: #FF5A00; text-align: center; background-color: #4C5155; border-style: solid; border-color: #4C5155" 
size=43 name=percent> 
      <SCRIPT language="">
 

var bar = 0 
var line = "||" 
var amount ="||" 
count() 
function count(){ 
bar= bar+2 
amount =amount + line 
document.loading.chart.value=amount 
document.loading.percent.value=bar+"%" 
if (bar<99) 
{setTimeout("count()",4);} 
else 
{window.location = "View_Contact.asp";} 
} 
</SCRIPT> 
</P> 
</DIV></FORM>	</td>
  </tr>
</table>
 
