<%
dim curid
curid=request("curid")
%>
<title>实例　天气预报</title>
<script language="javascript">
<!--
function showdata(IID)
{
window.location.href = "index.asp?curid=" + IID
}
-->
</script>
<body>
<p align="center"><font face="隶书" size="6" color="#ff0000"><strong>天气预报</strong></font></p>
<div style="LEFT: 5%; POSITION: absolute; TOP: 20%; WIDTH: 50%">
<MAP NAME="map1">
<AREA ID=a1 SHAPE="CIRCLE" COORDS="211, 103, 16" ALT  ="北京" ONCLICK="showdata(1)">
<AREA ID=a2 SHAPE="CIRCLE" COORDS="256, 157, 13" ALT  ="上海" ONCLICK="showdata(2)">
<AREA ID=a3 SHAPE="CIRCLE" COORDS="149, 164, 15" ALT  ="成都" ONCLICK="showdata(3)">
<AREA ID=a4 SHAPE="CIRCLE" COORDS="242, 88, 14" ALT  ="沈阳" ONCLICK="showdata(4)">
<AREA ID=a5 SHAPE="CIRCLE" COORDS="233, 146, 13" ALT  ="武汉" ONCLICK="showdata(5)">
<AREA ID=a6 SHAPE="CIRCLE" COORDS="210, 204, 14" ALT  ="广州" ONCLICK="showdata(6)">
<AREA ID=a7 SHAPE="CIRCLE" COORDS="209, 158, 13" ALT  ="南京" ONCLICK="showdata(7)">
<AREA ID=a8 SHAPE="CIRCLE" COORDS="179, 145, 14" ALT  ="西安" ONCLICK="showdata(8)">
</MAP>
<IMG border=0 height=287 id=IMG1 src="picture/Map.jpg" useMap=#map1> 
</div>
<%
dim rs
dim sql
dim selected
dim conn
set rs=server.createobject("adodb.recordset")
conn = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
if curid<>"" then
	sql="select * from mytable where id="+cstr(curid)
else 
	sql="select * from mytable where id=1"
end if
rs.open sql,conn,1,1
%><div style="LEFT: 50%; POSITION: absolute; TOP: 20%; WIDTH: 45%">
<table border="1" cellspacing="1" width="328">
    <TBODY style="COLOR: Coral" bgcolor=Ivory>
  <tr align="center">
    <td width="107">城市名称</td><td width="199"><%=rs("city_name")%></td>
  </tr>
  <tr align="center">
    <td width="107">今日温度</td><td width="199"><%=rs("todaytmp")%></td>
  </tr>
  <tr align="center">
	<td width="107">今日天气</td><td width="199"><img src="picture/<%=rs("todayday")%>.gif"></td>
  </tr>
  <tr align="center">
    <td width="107">明日温度</td><td width="199"><%=rs("tomorrowtmp")%></td>
  </tr>
  <tr align="center"><td width="107">明日天气</td>
    <td width="199"><img src="picture/<%=rs("tomorrowtday")%>.gif"></td>
  </tr>
  <tr align="center">
    <td width="107">后日趋势</td><td width="199"><%=rs("dayend")%></td>
  </tr></TBODY>
</table>
</div> 
</body>
