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
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=site.css rel=stylesheet>
<title></title>
</head>
<body topmargin="0">
<!--#include file="art_adminconn.asp" -->






<%
artist_id=request("artist_id")

if request("option")="添加" then

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
brief=trim(request("brief"))
brief=replace(brief,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_subs where sname='" & sname & "'"
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
if brief="" then brief=" "
if memo="" then memo=" "
sql="insert into art_subs (sname,simg,artist_id,sproperties,price,sstate,brief,memo,inserttime) values ('"&sname&"','"&simg&"',"&artist_id&",'"&sproperties&"',"&price&",'"&sstate&"','"&brief&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
  <TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
<FORM action="art_subs_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height=30 colspan="2"><font color="#FF0000" size="+1"><strong>作品添加</strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right">作品名称：</TD>
        <TD> 
          <input maxlength=100 size=30 id="sname" name="sname">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 作品图片：</TD>
        <TD>
<INPUT maxLength=100 size=30 id="simg" name="simg">
          * (例如：123456.jpg)</TD>
      </TR>
      <TR> 
        <TD align="right"> 所属美术家：</TD>
        <TD> <select name="artist_id" id="artist_id" class="display_dropdown">
            <option value="" selected>请选择……</option>
            <%
Set prs=Server.CreateObject("ADODB.RecordSet")
psql="select * from art_artist"
prs.Open psql,conn,1,1
do while not prs.eof
%>
            <option value=<%=prs("artist_id")%>><%=prs("allname")%></option>
            <%
prs.movenext
loop
prs.close
set prs=nothing
%>
          </select>
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 规格/参数：</TD>
        <TD> <INPUT maxLength=250 size=30 id="sproperties" name="sproperties">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> 市场价：</TD>
        <TD> <INPUT maxLength=8 size=30 id="price" name="price"> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 状态：</TD>
        <TD> <select name="sstate" id="sstate">
            <option value="已出售">已出售</option>
            <option value="未出售" selected>未出售</option>
            <option value="收藏">收藏</option>
          </select> </TD>
      </TR>
      <TR> 
        <TD align="right"> 简介：</TD>
        <TD> <textarea name="brief" id="brief" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> 备注：</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><INPUT type=submit id="option" name="option" value="添加">
        </TD>
      </TR>
    </TBODY>
</FORM>
  </TABLE>
<br>
<br>
</body>
</html>
