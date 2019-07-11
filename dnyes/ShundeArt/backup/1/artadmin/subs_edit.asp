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
<LINK href="css.css" type=text/css rel=stylesheet>
</head>
<body text="#000000" leftmargin="8" topmargin="8" marginwidth="0" marginheight="0">
<!--#include file="adminconn.asp" -->
<%
artist_id=request("artist_id")

if request("option")="edit" then

sname=trim(request("sname"))
sname=replace(sname,"'","''")
simg=trim(request("simg"))
simg=replace(simg,"'","''")
sproperties=trim(request("sproperties"))
sproperties=replace(sproperties,"'","''")
price=trim(request("price"))
price=replace(price,"'","''")
sstate=trim(request("sstate"))
sstate=replace(sstate,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_subs where sname='" & sname & "' and subs_id<>" & request("subs_id")
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
conn.Execute("update art_subs set sname='"&sname&"',simg='"&simg&"',artist_id="&artist_id&",sproperties='"&sproperties&"',price="&price&",sstate='"&sstate&"',memo='"&memo&"' where subs_id=" & request("subs_id"))
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('修改成功');</SCRIPT>")
end if
end if
%>
<%
sql="select * from art_subs where subs_id=" & request("subs_id") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>作品修改</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="subs_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=262 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">作品名称</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="sname" name="sname" value="<%=rs("sname")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">作品图片</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="simg" name="simg" value="<%=rs("simg")%>">
          * (例如：123456.jpg)</FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">所属美术家</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <select name="artist_id" id="artist_id" class="display_dropdown">
<%
psql="select * from art_artist where artist_id=" & rs("artist_id")
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("artist_id")%> selected><%=prs("allname")%></option>
<%
end if
prs.close
%>
<%
psql="select * from art_artist where artist_id<>" & rs("artist_id")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("artist_id")%>><%=prs("allname")%></option>
<%
prs.movenext
loop
end if
prs.close
%>
          </select>
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">规格/参数</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=250 size=30 id="sproperties" name="sproperties" value="<%=rs("sproperties")%>">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">市场价</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=8 size=30 id="price" name="price" value=<%=rs("price")%>>
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">状态</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <select name="sstate" id="sstate">
            <option value="<%=rs("sstate")%>" selected><%=rs("sstate")%></option>
            <option value="已出售">已出售</option>
            <option value="未出售">未出售</option>
            <option value="收藏">收藏</option>
          </select>
          </FONT></TD>
      </TR>
      <TR> 
        <TD width="150" height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备注</font></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="edit">
            </FONT></STRONG></P></TD>
      </TR>
    </TBODY>
  </TABLE>
<input type="hidden" id="subs_id" name="subs_id" value="<%=request("subs_id")%>">
</FORM>
<%
set prs=nothing
rs.close
set rs=nothing
%>
<br>
<br>
</body>
</html>
