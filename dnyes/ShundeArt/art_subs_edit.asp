<script language="JavaScript">

function checkform()
{
	if (form_administrator.sname.value.length == 0) {
		alert("��������Ʒ���� ");
		form_administrator.sname.focus();
		return false;
		}
	if (form_administrator.artist_id.value.length == 0) {
		alert("��ѡȡ�����Ʒ������ ");
		form_administrator.artist_id.focus();
		return false;
		}
	if (form_administrator.sproperties.value.length == 0) {
		alert("��������/����");
		form_administrator.sproperties.focus();
		return false;
		}
	if(! isNumber(form_administrator.price.value)) {
		alert('����������д');
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
subs_id=request("subs_id")

if request("option")="�༭" then

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
brief=trim(request("brief"))
brief=replace(brief,"'","''")
memo=trim(request("memo"))
memo=replace(memo,"'","''")

usql="select * from art_subs where sname='" & sname & "' and subs_id<>" & subs_id
set urs=server.createobject("ADODB.Recordset")
urs.open usql,conn,1,1
if not ( urs.eof or urs.bof ) then
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('��������Ʒ��ͬ�����ڣ�����������������');</SCRIPT>")
urs.close
set urs=nothing
else
urs.close
set urs=nothing

if price="" then price=0
if simg="" then simg=" "
if brief="" then brief=" "
if memo="" then memo=" "
conn.Execute("update art_subs set sname='"&sname&"',simg='"&simg&"',artist_id="&artist_id&",sproperties='"&sproperties&"',price="&price&",sstate='"&sstate&"',brief='"&brief&"',memo='"&memo&"' where subs_id=" & subs_id)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert('�޸ĳɹ�');</SCRIPT>")
end if
end if
%>
<%
sql="select * from art_subs where subs_id=" & subs_id 
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.open sql,conn,1,1
%>
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<TABLE width=100% 
                  border=1 cellPadding=3 cellSpacing=0 bordercolor="#CCCCCC">
<FORM action="art_subs_edit.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
    <TBODY>
      <TR align="center"> 
        <TD height="30" colspan="2"><font color="#FF0000" size="+1"><strong>��Ʒ�޸�</strong></font></TD>
      </TR>
      <TR> 
        <TD width="46%" align="right"> ��Ʒ���ƣ�</TD>
        <TD> <INPUT maxLength=100 size=30 id="sname" name="sname" value="<%=rs("sname")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ƷͼƬ��</TD>
        <TD> <INPUT maxLength=100 size=30 id="simg" name="simg" value="<%=rs("simg")%>">
          * (���磺123456.jpg)</TD>
      </TR>
      <TR> 
        <TD align="right"> ���������ң�</TD>
        <TD> <select name="artist_id" id="artist_id" class="display_dropdown">
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
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> ���/������</TD>
        <TD> <INPUT maxLength=250 size=30 id="sproperties" name="sproperties" value="<%=rs("sproperties")%>">
          * </TD>
      </TR>
      <TR> 
        <TD align="right"> �г��ۣ�</TD>
        <TD> <INPUT maxLength=8 size=30 id="price" name="price" value=<%=rs("price")%>> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> ״̬��</TD>
        <TD> <select name="sstate" id="sstate">
            <option value="<%=rs("sstate")%>" selected><%=rs("sstate")%></option>
            <option value="�ѳ���">�ѳ���</option>
            <option value="δ����">δ����</option>
            <option value="�ղ�">�ղ�</option>
          </select> </TD>
      </TR>
      <TR> 
        <TD align="right"> ��飺</TD>
        <TD> <textarea name="brief" id="brief" cols="60" rows="10"><%=rs("brief")%></textarea> 
        </TD>
      </TR>
      <TR> 
        <TD align="right"> ��ע��</TD>
        <TD> <textarea name="memo" id="memo" cols="60" rows="10"><%=rs("memo")%></textarea> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2> <INPUT type=submit id="option" name="option" value="�༭"> 
        </TD>
      </TR>
    </TBODY>
<input type="hidden" id="subs_id" name="subs_id" value="<%=subs_id%>">
</FORM>
  </TABLE>
<%
set prs=nothing
rs.close
set rs=nothing
%>
<br>
<br>
</body>
</html>
