<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("请输入作品名称 ");
		form_administrator.sname.focus();
		return false;
		}
	if (form_administrator.artist_id.value.length == 0) {
		alert("请选取这个作品的作者 ");
		form_administrator.artist_id.focus();
		return false;
		}
	if (form_administrator.sproperties.value.length == 0) {
		alert("请输入规格/参数");
		form_administrator.sproperties.focus();
		return false;
		}
	if(! isNumber(form_administrator.price.value)) {
		alert('请用数字填写');
		form_administrator.price.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
	}
	return true;
}

function isEnglish(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}
</script>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--media=print 这个属性可以在打印时有效--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style> 

<style> 
.tdp 
{ 
border-bottom: 1 solid #000000; 
border-left: 1 solid #000000; 
border-right: 0 solid #ffffff; 
border-top: 0 solid #ffffff; 
} 
.tabp 
{ 
border-color: #000000 #000000 #000000 #000000; 
border-style: solid; 
border-top-width: 2px; 
border-right-width: 2px; 
border-bottom-width: 1px; 
border-left-width: 1px; 
} 
.NOPRINT { 
font-family: "宋体"; 
font-size: 9pt; 
} 

</style> 

</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
  <!--#include file="art_adminconn.asp" -->
  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0">
    <TR class="tdp">
      <td>
	  <center class="Noprint" >
    <OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0>
    </OBJECT>
    <input type=button value=打印 onclick=document.all.WebBrowser.ExecWB(6,1)>&nbsp;
    <input type=button value=直接打印 onclick=document.all.WebBrowser.ExecWB(6,6)>&nbsp;
    <input type=button value=页面设置 onclick=document.all.WebBrowser.ExecWB(8,1)>&nbsp;
    <input name="button" type=button onClick=document.all.WebBrowser.ExecWB(7,1) value=打印预览>
      </center> 
</td>
    </tr>
  </table>





<%
subs_id=request("subs_id")

if request("option")="edit" then

sname=trim(request("sname"))
sname=replace(sname,"'","''")
simg=trim(request("simg"))
simg=replace(simg,"'","''")
artist_id=request("artist_id")
sproperties=trim(request("sproperties"))
sproperties=replace(sproperties,"'","''")
price=trim(request("price"))
price=replace(price,"'","''")
sstate=trim(request("sstate"))
sstate=replace(sstate,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_subs where sname='" & sname & "' and subs_id<>" & subs_id
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此作品有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if price="" then price=0
if simg="" then simg=" "
if memo="" then memo=" "
conn.Execute("update art_subs set sname='"&sname&"',simg='"&simg&"',artist_id="&artist_id&",sproperties='"&sproperties&"',price="&price&",sstate='"&sstate&"',memo='"&memo&"' where subs_id=" & subs_id)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('修改成功');</SCRIPT>")
end if
end if
%>
<%
sql="select * from art_subs where subs_id=" & subs_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
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
      <TD align="right" class="tdp"> 作品图片：</TD>
      <TD class="tdp"> <%=rs("simg")%></TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">所属美术家：</TD>
      <TD class="tdp">  
        <%
psql="select * from art_artist where artist_id=" & rs("artist_id")
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
        <%=prs("allname")%> 
        <%
end if
prs.close
%>
        </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp"> 规格/参数：</TD>
      <TD class="tdp"> <%=rs("sproperties")%> * </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">市场价：</TD>
      <TD class="tdp"> <%=rs("price")%> </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">状态：</TD>
      <TD class="tdp"> <%=rs("sstate")%> </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">简介：</TD>
      <TD class="tdp"> <%=rs("brief")%> </TD>
    </TR>
    <TR class="tdp"> 
      <TD align="right" class="tdp">备注：</TD>
      <TD height="300" valign="top" class="tdp"> <%=rs("memo")%> </TD>
    </TR>
  </TBODY>
</TABLE>

<%
set prs=nothing
rs.close
set rs=nothing
%>
</body>
</html>
