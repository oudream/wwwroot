<!--#include file="dbm_adminconn.asp" -->
<%
sql="select * from dbm_subs where subs_id=" & request("subs_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td valign="top"> <table width="102%" border="0" cellspacing="0" cellpadding="0">
              <tr bgcolor="#FFFFFF"> 
                <td height="3" bgcolor="#FFFFFF"><img src="images/kb.gif" width="9" height="3"></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr valign="top"> 
    <td background="images/menu-guest-l.gif"> <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td> <table border="0" cellpadding="0" cellspacing="0" width="98%" align="center">
              <tr> 
                <td width="100%" align=center> <br> <font size="+2"><%=rs("sname")%></font> <HR> </td>
              </tr>
              <tr> 
                <td width="100%"  align="center"><img src="<%=rs("simg")%>"> 
                  <br> <br> <H2>&nbsp;</H2>
                  <!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
                  <table width="90%" border="0" cellpadding="1" cellspacing="2" bgcolor="f5f5f5">
                    <tr> 
                      <td width="68" bgcolor="#FFFFFF">作品名称:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("sname")%></font></td>
                      <td width="50" bgcolor="#FFFFFF">尺寸:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("sproperties")%></font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFFFFF">用户:</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
                        <%
psql="select * from dbm_guest where guest_id=" & rs("guest_id")
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
                        <%=prs("allname")%> 
                        <%
end if
prs.close
%>
                        </font></td>
                      <td bgcolor="#FFFFFF">价格</td>
                      <td bgcolor="#FFFFFF"><font 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("price")%></font></td>
                    </tr>
                    <tr> 
                      <td bgcolor="#FFFFFF">简介:</td>
                      <td colspan="3" bgcolor="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 ><%=rs("memo")%></font></td>
                    </tr>
                    <tr align="center" bgcolor="#FFFFFF"> 
                      <td colspan="4"><A 
      href="javascript:print()">〖 打印此页 〗</a>　　<A 
      href="javascript:close()">〖 关闭窗口 〗</a></td>
                    </tr>
                  </table></td>
              </tr>
              <tr> 
                <td width="100%"  height="25"> <div align="center"> </div></td>
              </tr>
              <tr> 
                <td width=100% ><hr size=1></td>
              </tr>
              <tr> 
                <td height=8></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
</table>

<%
set prs=nothing
rs.close
set rs=nothing
%>
