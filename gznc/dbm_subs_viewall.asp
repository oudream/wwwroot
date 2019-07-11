<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="dbm_css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
  <!--#include file="dbm_adminconn.asp" -->
<%
subs_id=request("subs_id")
sql="select * from dbm_subs where subs_id=" & subs_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
memo = replace(rs("memo"),chr(13),"<br>")
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
  
<TABLE width=600 
                  border=0 align="center" cellPadding=5 cellSpacing=1 bordercolor="#000000" class="tabp">
  <TBODY>
    <TR align="center" class="tdp"> 
      <TD colspan="2" class="tdp">作品查看 </TD>
    </TR>
    <TR class="tdp"> 
      <TD width="20%" align="right" class="tdp"> 作品名称：</TD>
      <TD class="tdp"> <%=rs("sname")%> * </TD>
    </TR>
    <TR class="tdp"> 
      <TD colspan="2" align="right" class="tdp"> 
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
          <tr align="center"> 
            <td>图片 1</td>
            <td>图片 2</td>
            <td>图片 3</td>
            <td>图片 4</td>
          </tr>
          <tr align="center" valign="middle"> 
            <td><%if rs("sfile1")<>" " then %>
			<img align="middle" height="150" width="200" src="photo/<%=rs("sfile1")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile2")<>" " then %>
			<img align="middle" height="150" width="200" src="photo/<%=rs("sfile2")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile3")<>" " then %>
			<img align="middle" height="150" width="200" src="photo/<%=rs("sfile3")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile4")<>" " then %>
			<img align="middle" height="150" width="200" src="photo/<%=rs("sfile4")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
          </tr>
          <tr align="center"> 
            <td>图片 5</td>
            <td>图片 6</td>
            <td>图片 7</td>
            <td>图片 8</td>
          </tr>
          <tr align="center" valign="middle"> 
            <td><%if rs("sfile5")<>" " then %><img align="middle" height="150" width="200" src="photo/<%=rs("sfile5")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile6")<>" " then %><img align="middle" height="150" width="200" src="photo/<%=rs("sfile6")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile7")<>" " then %><img align="middle" height="150" width="200" src="photo/<%=rs("sfile7")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
            <td><%if rs("sfile8")<>" " then %><img align="middle" height="150" width="200" src="photo/<%=rs("sfile8")%>">
			<% 
			else
			response.Write("无图片")
			end if %>&nbsp;</td>
          </tr>
        </table></TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">简介：</TD>
      <TD class="tdp"> <%=memo%> </TD>
    </TR>
  </TBODY>
</TABLE>

<%
rs.close
set rs=nothing
%>
</body>
</html>
