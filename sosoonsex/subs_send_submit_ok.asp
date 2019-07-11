<head>
<title>Control panel - design by sosoon.cn</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body bgcolor="#FFFFFF">
<!--#include file="adminconn.asp" -->
<TABLE width=600 border=0 align="center" cellPadding=0 cellSpacing=5>
  <TBODY>
    <TR align="center"> 
      <TD height="120" colSpan=2><br> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="45" align="right"><font size="6" face="宋体"><img src="sosoon_logo.jpg" width="45" height="45"></font></td>
            <td width="40%"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="220" height="25"><strong><font size="3" face="宋体">&nbsp;数 
                    信 科 技 有 限 公 司</font></strong> </td>
                </tr>
                <tr> 
                  <td width="220" height="25" valign="middle"><strong><font size="3" face="宋体">SOSOON 
                    TECHNOLOYG CO.LTD</font></strong></td>
                </tr>
              </table></td>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td height="24" align="right"><strong>追求永远的质量</strong></td>
                </tr>
                <tr> 
                  <td align="right"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td bgcolor="#000000"><img src="1.jpg" width="1" height="1"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="24" align="right"><strong>追求永远的服务</strong></td>
                </tr>
              </table></td>
          </tr>
          <tr align="center"> 
            <td height="50" colspan="3"><font size="6" face="宋体"><strong>整 机 
              报 价 单</strong></font> </td>
          </tr>
        </table></TD>
    </TR>
    <TR align="center"> 
      <TD height="50" colSpan=2 align="center"><table width="100%" border="1" cellspacing="0" cellpadding="2">
          <tr> 
            <td width="80"><strong>种类</strong></td>
            <td width="250" height="35" align="center"><strong>品牌/型号</strong></td>
            <td width="80" height="35" align="center"><strong>数量/单位</strong></td>
            <td width="80" height="35" align="center"><strong>金额</strong></td>
            <td width="100" height="35" align="center"><strong>备注</strong></td>
          </tr>
          <%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set brs=Server.CreateObject("ADODB.RecordSet")
Set prs=Server.CreateObject("ADODB.RecordSet")
if session("subs_id_1")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_1")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>CPU</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_1")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_1")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_1")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_2")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_2")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>主板</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_2")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_2")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_2")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_3")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_3")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>内存</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_3")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_3")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_3")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_4")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_4")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>硬盘</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_4")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_4")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_4")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_5")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_5")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>显示卡</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_5")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_5")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_5")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_6")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_6")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>显示器</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_6")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_6")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_6")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_7")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_7")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>光驱</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_7")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_7")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_7")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_8")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_8")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>软驱</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_8")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_8")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_8")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_9")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_9")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>键盘</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_9")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_9")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_9")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_10")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_10")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>鼠标</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_10")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_10")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_10")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_11")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_11")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>机箱电源</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_11")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_11")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_11")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_12")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_12")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong>音箱</strong></td>
            <td width="250" height="30"> <strong> 
              <%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_12")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_12")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_12")%>&nbsp;</strong></td>
          </tr>
          <%
end if
rs.close
end if
%>
          <%
if session("subs_id_13")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_13")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
psql="select * from sort where sort_id=" & rs("sort_id")
prs.Open psql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong><%=prs("sname")%></strong></td>
            <td width="250" height="30"> <strong> 
              <%
response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_13")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_13")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_13")%>&nbsp;</strong></td>
          </tr>
          <%
prs.close
brs.close
end if
rs.close
end if
%>
          <%
if session("subs_id_14")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_14")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
psql="select * from sort where sort_id=" & rs("sort_id")
prs.Open psql,conn,1,1
%>
          <tr> 
            <td width="80" bgcolor="#CCCCCC"><strong><%=prs("sname")%></strong></td>
            <td width="250" height="30"> <strong> 
              <%
response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
%>
              </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("mount_14")&"/"&rs("suit")%> </strong></td>
            <td width="80" height="30" align="center"><strong><%=request("counter_14")%></strong></td>
            <td width="100" height="30" align="center"><strong><%=request("memo_14")%>&nbsp;</strong></td>
          </tr>
          <%
prs.close
brs.close
end if
rs.close
end if
%>
          <tr> 
            <td height="35" colspan="2" align="center">&nbsp;</td>
            <td height="35" colspan="2" align="center"><strong>总金额</strong> <font color="#FF0000"><strong><%=request("counter_all")%></strong></font></td>
            <td width="100" height="35" align="center">&nbsp;</td>
          </tr>
        </table>
        <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></TD>
    </TR>
    <TR> 
      <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="14%" height="50"><strong><font size="3">客户</font></strong></td>
            <td width="43%" height="50"><strong><font size="3"><%=request("gid")%>&nbsp; </font></strong></td>
            <td width="43%" height="50" align="right"> <strong><%=request("pyear")%>年 <%=request("pmonth")%>月 <%=request("pday")%>日</strong></td>
          </tr>
          <tr> 
            <td width="14%" height="30"><strong>客户备注</strong></td>
            <td height="30" colspan="2"><strong><%=request("gmemo")%>&nbsp;</strong></td>
          </tr>
        </table></TD>
    </TR>
    <TR> 
      <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><strong>本公司经手人及联系电话 <%=request("operationer_id")%></strong></td>
          </tr>
        </table></TD>
    </TR>
    <TR> 
      <TD height="50" colSpan=2> <P><strong>*以上产品没有特殊说明，即从购买日起均享受三个月包换、一年免费保修。<br>
          *本公司诚诺为广大客户提供终身技术支持及免费软件升级。</strong></P>
        <p><strong>地 址 ：广东省顺德区大良顺伟电脑城（即肯德基对面）三楼C3010-C3011. <br>
          写字楼 ：广东省顺德区大良金榜下街16座301.<br>
          服务热线 ：0757-22205321、22221633 图文传真:0757-22205321<br>
          http://www.sosoon.cn E-mail:sales@sosoon.cn</strong></p></TD>
    </TR>
  </TBODY>
</TABLE>
</body>
</html>