<style type="text/css">
<!--
body           {font-size: 9pt;}
table          {font-size: 9pt; }
tr             {height: 20;}
tr.over        {font-size: 9pt; color: #ffffff; background-color: #66aadd; cursor: hand;}
tr.out         {font-size: 9pt; color: #ffffff; background-color: #336699; cursor: default;}
div.rm_div     {position: absolute; filter: Alpha(Opacity='95'); display: none; border: 2px outset #3377aa; background-color: #336699; width: 0; height: 0;}
hr.sperator    {border: 1px inset #3377aa;}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
-->
</style>
<SCRIPT src="inc/js.js" type=text/javascript></SCRIPT>
<SCRIPT src="inc/exit.js" type=text/javascript></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
function runClock() {
	theTime = window.setTimeout("runClock()", 1000);
	var today = new Date();
	var display= today.toLocaleString();
	status="�����ǣ�"+display+"����ӭ���Ĺ��٣�";
}
</SCRIPT>
<body oncontextmenu="self.event.returnValue=false" onLoad="runClock()">
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
	<td background="IMAGES/top2.gif" align="center">
<%if Request.cookies(Forcast_SN)("ManageUserName")<>"" then%>
	<FONT color=#428eff>ϵͳ���ù���Ա [</FONT><font color=#000000><%=(Request.cookies(Forcast_SN)("ManageUserName"))%></font><FONT color=#428eff>]�ѵ�¼��������</font><a href=ExitManage.asp><font color=#000000>�˳�����</FONT></a>&nbsp;&nbsp;
<%end if%>
	</td>
</tr>
</table>
</BODY>