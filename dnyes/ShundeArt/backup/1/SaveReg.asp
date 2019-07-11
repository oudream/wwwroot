<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="ChkURL.asp"-->
<!--#include file=User_Top.asp-->

<%
dim sql
dim rs
dim webname
dim weburl
dim logo
dim webmaster
dim content
dim email
dim linktype
dim pass

webname=checkstr(request.form("webname"))
weburl=checkstr(request.form("weburl"))
content=htmlencode(request.form("content"))
logo=checkstr(request.form("logo"))
webmaster=checkstr(request.form("webmaster"))
email=checkstr(request.form("email"))
linktype=checkstr(request.form("linktype"))
pass=request.form("pass")

set rs=server.createobject("adodb.recordset")
sql="select * from "& db_Link_Table &" where (id is null)" 

rs.open sql,conn,1,3
rs.addnew
rs("linktype")=linktype
rs("webname")=webname
rs("content")=content
rs("weburl")=weburl
rs("logo")=logo
rs("webmaster")=webmaster
rs("email")=email
rs("pass")=pass
rs("dateandtime")=now()
rs.update
rs.close
%>

<html>
<head>
<title>申请成功</title>
<LINK href=site.css rel=stylesheet>
<script language="Javascript">
     function closeit() {
     setTimeout("self.close()",4000)
     }
     </script>
</head>

<body body onload="closeit()" topmargin="0">
<div align="center">
  <table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760"  id="AutoNumber1">
    <tr height="120">
      <td width="100%"> 
        <p align="center"><br><font color=red><b>友情链接申请成功</b><br></font><br>请首先在您的站点加上本站的链接。如果合适，本站将加上您的站点的链接！<br><br>
      </td>
    </tr>
    <tr valign="middle" height="120"> 
      <td width="100%" > 
        <div align="center"><br>
          本窗口将在4秒钟后自动关闭，您也可以直接点击下面关闭窗口。<br>
         <br> <a href="javascript:window.close()">关闭窗口</a><br>
          </div>
      </td>
    </tr>
    </table>
</div>
</body>
</html>
<!--#include file=User_Bottom.asp-->