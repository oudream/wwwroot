<!--#include file="../conn/conn.asp"-->
<html>
<head>
<title>Delete</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="VBScript">
sub check()
if form1.selfield.value=empty then
msgbox "Please enter the field!"
focusto(0)
exit sub
end if 
if form1.selmonth.value=empty then
msgbox "Please enter the month!"
focusto(1)
exit sub
end if 
if form1.selday.value=empty then
msgbox "Please enter the day!"
focusto(1)
exit sub
end if 
if form1.selyear.value=empty then
msgbox "Please enter the year!"
focusto(2)
exit sub
end if 
if form1.selquantum.value=empty then
msgbox "Please enter the quantum!"
focusto(2)
exit sub
end if 
form1.submit
end sub
sub focusto(x)
document.form1.elements(x).focus()
end sub 
</script>
<body>
<form name="form1" method="post" action="scheduledel.asp">
  <p>Enter the field : &nbsp;&nbsp; 
    <select name="selfield" >
      <option selected>Select the field ...</option>
      <option value="1">Katong A</option>
      <option value="2">Katong B</option>
      <option value="3">New Town</option>
      <option value="4">Saint Andrew's</option>
      <option value="5">Sentosa East</option>
      <option value="6">Sentosa West Field</option>
      <option value="7">Serangonn</option>
      <option value="8">Tjc</option>
    </select>
  </p>
  <p>Enter the date : &nbsp;&nbsp;&nbsp;&nbsp; 
    <select name="selmonth">
      <option>NULL.</option>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      <option value="9">9</option>
      <option value="10">10</option>
      <option value="11">11</option>
      <option value="12">12</option>
    </select>
    Month 
    <select name="selday">
      <option>NULL.</option>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
      <option value="8">8</option>
      <option value="9">9</option>
      <option value="10">10</option>
      <option value="11">11</option>
      <option value="12">12</option>
      <option value="13">13</option>
      <option value="14">14</option>
      <option value="15">15</option>
      <option value="16">16</option>
      <option value="17">17</option>
      <option value="18">18</option>
      <option value="19">19</option>
      <option value="20">20</option>
      <option value="21">21</option>
      <option value="22">22</option>
      <option value="23">23</option>
      <option value="24">24</option>
      <option value="25">25</option>
      <option value="26">26</option>
      <option value="27">27</option>
      <option value="28">28</option>
      <option value="29">29</option>
      <option value="30">30</option>
      <option value="31">31</option>
    </select>
    Day
    <select name="selyear">
      <option>NULL...</option>
      <option value="2002">2002</option>
      <option value="2003">2003</option>
      <option value="2004">2004</option>
      <option value="2005">2005</option>
      <option value="2006">2006</option>
      <option value="2007">2007</option>
      <option value="2008">2008</option>
      <option value="2009">2009</option>
      <option value="2010">2010</option>
    </select>
    Year</p>
  <p>Enter the quantum : 
    <select name="selquantum">
      <option>NULL.</option>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="5">5</option>
      <option value="6">6</option>
      <option value="7">7</option>
    </select>
  </p>
  <p align="right"> 
    <input name="delsubmit" type="submit" id="delsubmit" value="Submit" onClick="check()">&nbsp;&nbsp;&nbsp;
    <input name="delreset" type="reset" id="delreset" value="Reset"  >
  </p>
</form>
<%
delfield=request("selfield")
delmonth=request("selmonth")
delday=request("selday")
delyear=request("selyear")
delquantum=request("selquantum")

delete="DELETE * FROM main where field="&"'"&cstr(delfield)&"'"&" and "&"year="&"'"&cstr(delyear)&"'"&" and "&"month="&"'"&cstr(delmonth)&"'"&" and "&"day="&"'"&cstr(delday)&"'"&" and "&"quantum="&"'"&cstr(delquantum)&"'"
conn.Execute delete

%>
</body>

</html>























