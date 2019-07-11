<!--#include file="conn/conn.asp"-->

<html>
<head>
<title>Ocean12 ASP Calendar Manager</title>

<script language="javascript">

function open_Event_Name() {
  popupWin = window.open('help.asp?ID=Event_Name', 'remote', 'width=300,height=200')
}
function open_Event_Date() {
  popupWin = window.open('help.asp?ID=Event_Date', 'remote', 'width=300,height=200')
}
function open_Event_Time() {
  popupWin = window.open('help.asp?ID=Event_Time', 'remote', 'width=300,height=200')
}
function open_Event_Category() {
  popupWin = window.open('help.asp?ID=Event_Category', 'remote', 'width=300,height=200')
}
function open_Event_Location() {
  popupWin = window.open('help.asp?ID=Event_Location', 'remote', 'width=300,height=200')
}
function open_Event_Description() {
  popupWin = window.open('help.asp?ID=Event_Description', 'remote', 'width=300,height=200')
}

function VerifyData() {
	if (document.frmUser.Event_Name.value == "") {
		alert("请输入日程标题.");
		return false;
			} else if (
				(document.frmUser.Time_Hour.value == "Blank") ||
				(document.frmUser.Time_Minute.value == "Blank") ||
				(document.frmUser.Time_AMPM.value == "Blank")) {
		alert("请你选择日程时间.");
		return false;
	} else if (document.frmUser.Category.value == "") {
		alert("请选择日程类型.");
		return false;

	} else
		return true;
}
</script>
<!--#include file="schedule/style.asp"-->
<link rel="stylesheet" href="schedule/css/css1.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" vlink="#48576C" link="#48576C" alink="#000000" topmargin="3">
<table width=100% border=0 cellspacing=0 cellpadding=0>
  <tr> 
    <td height=22 class=a><font face="Arial"><b><font size="3">欢迎<font color="#FF6600"><%=session("ename")%></font>登录日程安排系统</font></b></font></td>
    <td align=right><a href="schedule/main.asp"><font color="#000000">返回首页</font></a></td>
  </tr>
  <tr> 
    <td  colspan=2></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td align=center background="schedule/images/v.gif"></td>
  </tr>
</table>
<br>
当前位置：日程安排 &gt; 修改日程<font face="Verdana" size="1"><br>
<%
dim RSFORM
Set RSFORM = Server.CreateObject("ADODB.Recordset")
RSFORM.Open "schedule_events", Conn, 2, 2

RSFORM.Find "ID='" & request("ID") & "'"
%>
</font>
<form action="schedule/scheduleeditsave.asp" name="frmUser" method="Post" onSubmit="return VerifyData()">
  <div align="center"><font face="Verdana" size="1"> 
    <input type="hidden" name="ID" value="<%=request("ID")%>">
    <input type="hidden" name="orderby" value="<%=request("orderby")%>">
    <input type="hidden" name="page" value="<%=request("page")%>">
    <input type="hidden" name="SearchFor" value="<%=request("SearchFor")%>">
    <input type="hidden" name="SearchWhere" value="<%=request("SearchWhere")%>">
    </font> </div>
  <div align="center">
    <table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#C0C0C0" width="80%" id="AutoNumber3" height="20">
      <tr> 
        <td bgcolor="#E6EEFF" height="24"> 
          <div align="center"><font color="#000000">日程标题</font></div>
        </td>
        <td bgcolor="#FFFFFF" height="24"> 　 
          <input type="text" name="Event_Name" value="<%=RSFORM("Event_Name")%>" onBlur="this.className='inputnormal'" onFocus="this.className='inputedit';this.select()" class="inputnormal" size="30" maxlength="255">
        </td>
      </tr>
      <tr> 
        <td bgcolor="#E6EEFF" height="24"> 
          <div align="center"><font color="#000000">日程日期</font></div>
        </td>
        <td bgcolor="#FFFFFF" height="24"> 　 
          <%date1=request("date")%>
          <%=year(date1)%>年<%=month(date1)%>月<%=day(date1)%>日 
          <input type="hidden" name="date" value="<%=date1%>">
        </td>
      </tr>
      <tr> 
        <td bgcolor="#E6EEFF" height="24"> 
          <div align="center"><font color="#000000">日程时间</font></div>
        </td>
        <td bgcolor="#FFFFFF" height="24"> 　 
          <select name="Time_Hour" class="stedit">
            
            <% 
hourCounter = 1
hourSelected = CInt(Hour(RSFORM("Date")))

If hourSelected >=12 Then
	hourSelected = hourSelected - 12
End If

Do WHILE hourCounter <= 12
If (hourSelected <> hourCounter) Then %>
            <option value="<%=hourCounter%>"><%=hourCounter%></option>
            <% Else %>
            <option value="<%=hourCounter%>" SELECTED><%=hourCounter%></option>
            <% End If
hourCounter = hourCounter + 1
Loop
%>
          </select>
          <select name="Time_Minute" class="stedit">
           
            <% 
minuteCounter = 0
minuteSelected = CInt(Minute(RSFORM("Date")))

Do WHILE minuteCounter <= 55
If (minuteSelected <> minuteCounter) Then %>
            <option value="<%=minuteCounter%>"> 
            <%If (minuteCounter < 10) Then%>
            0 
            <%End If%>
            <%=minuteCounter%></option>
            <% Else %>
            <option value="<%=minuteCounter%>" SELECTED> 
            <%If (minuteCounter < 10) Then%>
            0 
            <%End If%>
            <%=minuteCounter%></option>
            <% End If
minuteCounter = minuteCounter + 5
Loop
%>
          </select>
          <select name="Time_AMPM" class="stedit">
          
            <%
AMPMSelected = CInt(Hour(RSFORM("Date")))
%>
            <option value="AM" <%If (AMPMSelected <= 11) Then%>SELECTED<%End If%>>AM</option>
            <option value="PM" <%If (AMPMSelected >= 12) Then%>SELECTED<%End If%>>PM</option>
          </select>
        </td>
      </tr>
      <tr> 
        <td bgcolor="#E6EEFF" height="24"> 
          <div align="center"><font color="#000000">日程分类</font></div>
        </td>
        <td bgcolor="#FFFFFF" height="24"> 　 
          <select name="Category" class="stedit">
            
            <%
Set RSCATEGORY = Server.CreateObject("ADODB.Recordset")
RSCATEGORY.Open "schedule_Categories", Conn, 2, 2

Do While Not RSCATEGORY.eof

If (RSCATEGORY("Category") <> RSFORM("Category")) Then %>
            <option value="<%=RSCATEGORY("Category")%>"><%=RSCATEGORY("Category")%></option>
            <% Else %>
            <option value="<%=RSCATEGORY("Category")%>" SELECTED><%=RSCATEGORY("Category")%></option>
            <% End If

RSCATEGORY.movenext
Loop
RSCATEGORY.close
set RSCATEGORY = nothing
%>
          </select>
        </td>
      </tr>
      <tr> 
        <td bgcolor="#E6EEFF" height="24"> 
          <div align="center"><font color="#000000">日程地点</font></div>
        </td>
        <td bgcolor="#FFFFFF" height="24"> 　 
          <input type="text" name="Location" value="<%=RSFORM("Location")%>" onBlur="this.className='inputnormal'" onFocus="this.className='inputedit';this.select()" class="inputnormal" size="30" maxlength="255">
        </td>
      </tr>
      <tr> 
        <td bgcolor="#E6EEFF"> 
          <div align="center"><font color="#000000">具体情况</font></div>
        </td>
        <td bgcolor="#FFFFFF"> 　 
          <textarea name="Description" onBlur="this.className='inputnormal'" onFocus="this.className='inputedit';this.select()" class="inputnormal" cols="45" rows="5"><%=RSFORM("Description")%></textarea>
        </td>
      </tr>
      <tr> 
        <td align="right" valign="bottom" bgcolor="#E6EEFF">&nbsp; </td>
        <td valign="bottom" bgcolor="#FFFFFF"> <font face="Verdana" size="1"><br>
          　 
          <input type="submit" name="Submit" value=" 修改日程 " class="s02">
          　　 
          <input type="reset" name="Reset" value=" 取消 " class="s02">
          <br>
          <br>
          </font></td>
      </tr>
    </table>
  </div>
  <div align="left">
    <div align="center"></div>
    <font size="1" face="Verdana"> </font> </div>
</form>
<font face="Verdana" size="1">
<%
RSFORM.close
set RSFORM = nothing
%>
</font> 
</body>
</html>
