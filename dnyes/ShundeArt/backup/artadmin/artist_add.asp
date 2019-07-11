<script language="JavaScript">

function checkform()
{
	if (form_administrator.allname.value.length == 0) {
		alert("请输入全称");
		form_administrator.allname.focus();
		return false;
		}
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('请输入正确的邮件地址');
		form_administrator.email.focus();
		return false;
		}}
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

function isMail(name)
{
	if(! isEnglish(name))
		return false;
	i = name.indexOf("@");
	j = name.lastIndexOf("@");
	if(i == -1)
		return false;
	if(i != j)
		return false;
	if(i == name.length)
		return false;
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
if request("option")="add" then

allname=trim(request("allname"))
allname=replace(allname,"'","''")
simname=trim(request("simname"))
simname=replace(simname,"'","''")

sex=trim(request("sex"))
sex=replace(sex,"'","''")

nation=trim(request("nation"))
nation=replace(nation,"'","''")

native=trim(request("native"))
native=replace(native,"'","''")

born_time=trim(request("born_time"))
born_time=replace(born_time,"'","''")

graduation=trim(request("graduation"))
graduation=replace(graduation,"'","''")

special=trim(request("special"))
special=replace(special,"'","''")

job_now=trim(request("job_now"))
job_now=replace(job_now,"'","''")

job_old=trim(request("job_old"))
job_old=replace(job_old,"'","''")
            
job_rank=trim(request("job_rank"))
job_rank=replace(job_rank,"'","''")

expert=trim(request("expert"))
expert=replace(expert,"'","''")

awarded=trim(request("awarded"))
awarded=replace(awarded,"'","''")

standest=trim(request("standest"))
standest=replace(standest,"'","''")

interest=trim(request("interest"))
interest=replace(interest,"'","''")

addr=trim(request("addr"))
addr=replace(addr,"'","''")
tel=trim(request("tel"))
tel=replace(tel,"'","''")
fax=trim(request("fax"))
fax=replace(fax,"'","''")
zip=trim(request("zip"))
zip=replace(zip,"'","''")
email=trim(request("email"))
email=replace(email,"'","''")
contacter=trim(request("contacter"))
contacter=replace(contacter,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")
if allname<>"" and simname<>"" then
usql="select * from art_artist where allname='"&allname&"' or simname='" & simname&"'"
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('出错！此美术家有同名存在，请用其它名字输入');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if sex="" then sex=" "

if nation="" then nation=" "

if native="" then native=" "

if born_time="" then born_time=" "

if graduation="" then graduation=" "

if special="" then special=" "

if job_now="" then job_now=" "

if job_old="" then job_old=" "

if job_rank="" then job_rank=" "

if expert="" then expert=" "

if awarded="" then awarded=" "

if standest="" then standest=" "

if interest="" then interest=" "

if addr="" then addr=" "
if tel="" then tel=" "
if fax="" then fax=" "
if zip="" then zip=" "
if email="" then email=" "
if contacter="" then contacter=" "
if memo="" then memo=" "
sql="insert into art_artist (allname,simname,sex, nation, native, born_time, graduation, special, job_now, job_old, job_rank, expert, awarded, standest, interest, addr,tel,fax,zip,email,contacter,memo,inserttime) values ('"&allname&"','"&simname&"','"&sex&"','"&nation&"','"&native&"','"&born_time&"','"&graduation&"','"&special&"','"&job_now&"','"&job_old&"','"&job_rank&"','"&expert&"','"&awarded&"','"&standest&"','"&interest&"','"&addr&"','"&tel&"','"&fax&"','"&zip&"','"&email&"','"&contacter&"','"&memo&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('添加成功');</SCRIPT>")
end if
end if
end if
%>
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>美术家添加</FONT></H2>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="artist_add.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
  <TABLE height=307 cellSpacing=0 cellPadding=0 width=550 
                  border=0>
    <TBODY>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">姓名</font></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="allname" name="allname">
          * </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">雅号</font></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="simname" name="simname">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">性别</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <select name="sex" id="sex">
            <option value="男" selected>男</option>
            <option value="女">女</option>
          </select>
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">民族</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="nation" name="nation">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">原籍</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="native" name="native">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">出生年月</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="born_time" name="born_time">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">毕业院校</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="graduation" name="graduation">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">所学专业</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="special" name="special">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">现任职务</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="job_now" name="job_now">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">曾任职务</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="job_old" name="job_old">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">职称</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="job_rank" name="job_rank">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">擅长</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="expert" name="expert">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">曾获奖项</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="awarded" name="awarded">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">代表作品</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="standest" name="standest">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">兴趣爱好</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="interest" name="interest">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">地址</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="addr" name="addr">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">电话</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="tel" name="tel">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">传真</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="fax" name="fax">
          </FONT></TD>
      </TR>
     <TR> 
        <TD width=120 height=45> <DIV align=center><font color="#000000"  face="Verdana, Arial, Helvetica, sans-serif">邮编</font></DIV></TD>
        <TD width=396><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=100 size=30 id="zip" name="zip">
          </FONT></TD>
      </TR>
      <TR> 
        <TD width=120 height=45> <DIV align=center><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 >E-mail</FONT></DIV></TD>
        <TD width=307><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000000 > 
          <INPUT maxLength=40 size=30 id="email" name="email">
          </FONT></TD>
      </TR>
      <TR> 
        <TD height=32> <DIV align=center><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">备注</font></DIV></TD>
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
<INPUT maxLength=100 size=30 id="contacter" name="contacter" type="hidden">
</FORM>
<br>
<br>
</body>
</html>
