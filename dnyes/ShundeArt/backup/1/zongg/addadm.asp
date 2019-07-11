<meta http-equiv="Content-Language" content="zh-cn">
<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file=config.asp -->
<%
if trim(request("job"))="add" then
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from admin where aid<>"&session("aid")&" and admin='"&trim(request("admname"))&"'"
adsrs.open adssql,adsconn,1,1
if not adsrs.eof then
%>
<script language="VBScript" type="text/VBScript">
msgbox "管理员名已经存在，请重新输入！"
 history.back(1)
</script>
              <%
else
adsrs.close
adssql="select * from admin where aid="&session("aid")
adsrs.open adssql,adsconn,3,3
adsrs(1)=trim(request("admname"))
adsrs(2)=trim(request("pass"))
adsrs.update
response.redirect "?job=suc"
end if
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing

end if
%>

<script language="javascript">
<!--
function isok(theform)
{
    if (theform.admname.value=="")
    {
        alert("请填写管理员名！");
        theform.admname.focus();
        return (false);
    }
       if (theform.pass.value=="")
    {
        alert("请填写管理密码！");
        theform.pass.focus();
        return (false);
    }
           if (theform.pass1.value=="")
    {
        alert("请填写验证密码！");
        theform.pass1.focus();
        return (false);
    }
    if (theform.pass1.value!=theform.pass.value)
    {
        alert("两次输入的密码不相同！");
        theform.pass1.focus();
        return (false);
    }
}
-->
</script>
<script language=JavaScript>
<!--

function opw(url,name, width, height) { //v2.0
window.open(url,name,''+'width='+width+',height='+height+',scrollbars=yes'+'');
}
//-->
</script>
<!--#include file="top.asp"-->


<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="102">
<%
adsconn.open adsdata
set adsrs=server.createobject("adodb.recordset")
adssql="select * from admin where aid="&session("aid")
adsrs.open adssql,adsconn,1,1%>	
<form method="POST"  action="?job=add" onSubmit="return isok(this)"> <tr>
     <td width="100" align="right" height="25">管理员名：</td>
      <td height="25"><input type="text" name="admname" value="<%=adsrs(1)%>" size="20"></td>
      <td height="25" width="100">
      <p align="right">管理密码：</td>
      <td height="25">
      <input type="password" name="pass" value="<%=adsrs(2)%>" size="20"></td>
      <td colspan="3" align="center">
    
<%if request("job")="suc" then%><font color=red>管理信息修改成功！</font><%end if%>

</td>
    </tr> <tr>
     <td align="right" colspan="2">
     <p align="center"><font color="#808080">下一版本将推出多管理员权限分配</font></td>
      <td height="25" width="100">
      <p align="right">验证密码：</td>
      <td height="25">
      <input type="password" name="pass1" value="<%=adsrs(2)%>" size="20"></td>
      <td  colspan="3" align="center"> <input type="submit" value="确定修改" name="B1"></td>   
    </tr> <tr>
     <td align="right" colspan="2">　</td>
      <td height="25" width="100">
      　</td>
      <td height="25">　</td>
      <td width="100" height="25">
      　</td>
      <td height="25">　</td>
      <td> 　</td>   
    </tr></form>
<%
adsrs.close
set adsrs=nothing
adsconn.close
set adsconn=nothing
%>
  </table>
  </center>
</div>



  <!--#include file="boot.asp"-->