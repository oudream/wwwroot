<script language="JavaScript">

function gettarg(targ,selObj,restore){ 
  eval(targ+".location='subs_add.asp?sort_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function checkform()
{
	if (form_administrator.sort_id.value.length == 0) {
		alert("请选取这个商品的种类 ");
		form_administrator.sort_id.focus();
		return false;
		}
	if (form_administrator.brand_id.value.length == 0) {
		alert("请选取这个商品的品牌 ");
		form_administrator.brand_id.focus();
		return false;
		}
	if (form_administrator.code.value.length == 0) {
		alert("请输入型号");
		form_administrator.code.focus();
		return false;
		}
	if (form_administrator.sproperties.value.length == 0) {
		alert("请输入规格/参数");
		form_administrator.sproperties.focus();
		return false;
		}
	if (form_administrator.suit.value.length == 0) {
		alert("请输入单位 ");
		form_administrator.suit.focus();
		return false;
		}
	if(! isNumber(form_administrator.pricein.value)) {
		alert('请用数字填写');
		form_administrator.pricein.focus();
		return false;
		}
	if(! isNumber(form_administrator.priceout.value)) {
		alert('请用数字填写');
		form_administrator.priceout.focus();
		return false;
		}
	if(! isNumber(form_administrator.basenum.value)) {
		alert('请用数字填写');
		form_administrator.basenum.focus();
		return false;
		}
	if(! isNumber(form_administrator.baseown.value)) {
		alert('请用数字填写');
		form_administrator.baseown.focus();
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
sort_id=request("sort_id")
if sort_id="" then sort_id=0

if request("option")="add" then

sort_id=request("sort_id")
brand_id=request("brand_id")
code=trim(request("code"))
code=replace(code,"'","''")
sproperties=trim(request("sproperties"))
sproperties=replace(sproperties,"'","''")
suit=trim(request("suit"))
suit=replace(suit,"'","''")
pricein=trim(request("pricein"))
pricein=replace(pricein,"'","''")
priceout=trim(request("priceout"))
priceout=replace(priceout,"'","''")
basenum=trim(request("basenum"))
basenum=replace(basenum,"'","''")
baseown=trim(request("baseown"))
baseown=replace(baseown,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

if sort_id<>"" and brand_id<>"" and code<>"" then
usql="select * from subs where sort_id=" & sort_id & " and brand_id=" & brand_id & " and code='" & code & "'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此商品有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if pricein="" then pricein=0
if priceout="" then priceout=0
if basenum="" then basenum=0
if baseown="" then baseown=0
if memo="" then memo=" "
sql="insert into subs (sort_id,brand_id,code,sproperties,suit,pricein,priceout,basenum,baseown,memo,inserttime) values ("&sort_id&","&brand_id&",'"&code&"','"&sproperties&"','"&suit&"',"&pricein&","&priceout&","&basenum&","&baseown&",'"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>商品添加</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="subs_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">所属种类</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >
          <select name="sort_id" id="select" class="display_dropdown"  onChange="gettarg('self',this,0)">
          <%
Set prs=Server.CreateObject("ADODB.RecordSet")
if sort_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from sort where sort_id=" & sort_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("sort_id")%> selected><%=prs("sname")%></option>
          <%
end if
prs.close
end if
%>
          <%
psql="select * from sort where sort_id<>" & sort_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("sort_id")%>><%=prs("sname")%></option>
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
        <TD width=150 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">所属品牌</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >
                    <select name="brand_id" id="brand_id" class="display_dropdown">
          <%
if sort_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="SELECT * FROM brand where brand_id in ( select brand_id from sb where sort_id="&sort_id&" ) order by bname"
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("brand_id")%>><%=prs("bname")%></option>
          <%
prs.movenext
loop
end if
prs.close
end if
%>
        </select>
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">型号</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=50 size=30 id="code" name="code">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">规格/参数</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=250 size=30 id="sproperties" name="sproperties">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">单位</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="suit" name="suit">
          * </FONT></TD>
      </TR>
     <TR> 
        <TD width=150 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">进货价</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=8 size=30 id="pricein" name="pricein">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >零售价</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=8 size=30 id="priceout" name="priceout">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >库存数量</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=8 size=30 id="basenum" name="basenum">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=150 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >库存欠量</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=8 size=30 id="baseown" name="baseown">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width="150" height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备注</font></DIV></TD>
        <TD><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <textarea name="memo" id="memo" cols="60" rows="10"></textarea>
          </FONT></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2>
<INPUT type=submit id="option" name="option" value="add">
            </FONT></STRONG></P>
          </TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
<br>
<br>
</body>
</html>
