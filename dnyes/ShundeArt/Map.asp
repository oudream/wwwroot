<%@ Language=VBScript%>
<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>网站地图__<%=jjgn%></title>
<LINK href=news.css rel=stylesheet>
</head>

<body topmargin="0">
<!--#include file="top.asp"-->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr valign="top"> 
    <td width="760" class="daohang"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="3" bgcolor="#FFFFFF"><img src="IMAGES/kb.gif" width="9" height="3"></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-guestbook.gif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font class=m_tittle>&nbsp;</font>栏目导航&nbsp;&nbsp;<b>当前位置：<a class="daohang" href="./" >网站首页</a>&gt;&gt;网站地图</b></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-guest-l.gif"><div align="center"> 
              <a class=daohang href="./" > 
              <script language=javascript src=./zongg/ad.asp?i=13></script> 
              </a> </div></td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-guestbook.gif">　</td>
        </tr>
        <tr> 
          <td height="25" background="IMAGES/menu-guest-l.gif">
              <table width="95%" height="156" border="0" align="center" cellpadding="0" cellspacing="0" id="AutoNumber3" style="border-collapse: collapse">
                <tr> 
                  <td height="10" valign=top>&nbsp;</td>
                </tr>
                <tr> 
                  <td height="10" valign=top><table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td width="20%" height="20" colspan="5"> 
			<%
dim menuid1
dim menuname1
dim menucontent1
dim menuview1
set rs22=server.CreateObject("ADODB.RecordSet")
rs22.Source="select * from " & db_Type_Table & " where typeview=1 order by typeorder"
rs22.Open rs22.Source,conn,1,1
i=1

while not rs22.EOF
RecordCount=rs22.RecordCount
menuid1=rs22("typeid")
menuname1=rs22("typename")
menucontent1=rs22("typecontent")
menuview1=rs22("typeview")

%> <img src="images/m.gif" width="13" height="11"><a class=middle target="_top" href='Type.asp?typeID=<%=menuid1%> ' title=<%=menucontent1%> ><%=menuName1 %></a><br> <%  
typeid=rs22("typeid")
	set rs21=conn.execute("SELECT * FROM " & db_BigClass_Table & " where typeid="&typeid&" order by bigclassorder") 
	do while not rs21.eof
%> <%if not Rs21.eof then%> <table width="100%" border="0" cellpadding="3" cellspacing="0">
                <tr> 
                  <td width="18%" background="images/h.gif" bgcolor="#CCCCCC">&nbsp;&nbsp;<img src="images/m1.gif" width="13" height="11"><a target="_top" class=middle href="BigClass.asp?typeid=<%=typeid%>&bigclassid=<%=Rs21("bigclassid")%>"><%=Rs21("bigclassname")%></a></td>
                  <td width="82%" background="images/h.gif" bgcolor="#FFFFFF"> <%set nrs=conn.execute("SELECT * FROM "& db_SmallClass_Table &" where bigclassid="&cstr(rs21("bigclassid"))&" order by smallclassorder")%> <%do while not nrs.eof%> <%if not nRs.eof then%> <a target="_top" class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a> <%nrs.movenext   
		     end if %> <%if not nRs.eof then%> <a target="_top" class=class href="SmallClass.asp?typeid=<%=typeid%>&BigClassID=<%=nrs("bigclassid")%>&smallclassid=<%=nrs("smallclassid")%>"><%=nRs("smallclassname")%></a> <%nRs.movenext
             end if%> <%loop
            nRs.Close
           set nRs=nothing
            %> </td>
                </tr>
              </table>
              <%rs21.movenext   
	end if
	loop  
	rs21.close
%> <br>
              
              <%


i=i+1

rs22.MoveNext
wend
rs22.close
set rs22=nothing
%>
                      </td>
                    </tr><tr> 
                  <td height="10" valign=top>&nbsp;</td>
                </tr>
              </table></td>
        </tr>
      </table>
	  </td>
  </tr>
  </table>
	  </td>
  </tr>
  <tr valign="top">
    <td height="20" background="IMAGES/menu-guest-t.gif" class="daohang">&nbsp;</td>
  </tr>
</table>
<!--#include file=bottom.asp -->
</body>
</html>
