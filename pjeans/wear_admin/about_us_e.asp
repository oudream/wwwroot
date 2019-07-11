<SCRIPT LANGUAGE="JavaScript">
<!--

function checkform()
{
	if (zform.uname.memo.length == 0) {
		alert("Please Enter the meno of competition. ");
		zform.memo.focus();
		return false;
		}
	if(! isEnglish(zform.memo.value)) {
		alert(' Sorry ! the name must be made up by ENGLISH');
		zform.memo.focus();
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

function isMail(name)
{
	if(name.length==0)
		return true;
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

//-->

</SCRIPT>

<!--#include file="adminconn.asp" -->
<%
ztid=request("ztid")

if request("option")="edit" then

wear_name=trim(request("wear_name"))
wear_name=replace(wear_name,"'","''")
if wear_name="" then wear_name=" "

wear_tel=trim(request("wear_tel"))
wear_tel=replace(wear_tel,"'","''")
if wear_tel="" then wear_tel=" "

wear_fax=trim(request("wear_fax"))
wear_fax=replace(wear_fax,"'","''")
if wear_fax="" then wear_fax=" "

wear_addr=trim(request("wear_addr"))
wear_addr=replace(wear_addr,"'","''")
if wear_addr="" then wear_addr=" "

wear_email=trim(request("wear_email"))
wear_email=replace(wear_email,"'","''")
if wear_email="" then wear_email=" "

wear_memo=trim(request("wear_memo"))
wear_memo=replace(wear_memo,"'","''")
if wear_memo="" then wear_memo=" "

conn.Execute("update myself set wear_name='"&wear_name&"',wear_tel='"&wear_tel&"',wear_fax='"&wear_fax&"',wear_addr='"&wear_addr&"',wear_email='"&wear_email&"',wear_memo='"&wear_memo&"' where wid=1")
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Edit about us is complete .');</SCRIPT>")
end if
%>
<%
sql="select * from myself where wid=1"
Set rs=Server.CreateObject("ADODB.RecordSet") 
rs.Open sql,conn,1,1 
if rs.bof or rs.eof then response.Redirect("error.asp")
%>

<form id="zform" action="about_us_e.asp" onSubmit="return checkform();" method="post">
  <table width="100%" border="0" cellspacing="2" cellpadding="2">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td colspan="2">&nbsp;</td>
      <td width="4%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2"><H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000066>About US EDIT Wizard </FONT></H2></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>SITE Name :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="wear_name"  name="wear_name" maxlength=98 size=80 value="<%=rs("wear_name")%>">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Telphone :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="wear_tel"  name="wear_tel" maxlength=98 size=80 value="<%=rs("wear_tel")%>">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>FAX :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="wear_fax"  name="wear_fax" maxlength=98 size=80 value="<%=rs("wear_fax")%>">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>ADDRESS :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="wear_addr"  name="wear_addr" maxlength=98 size=80 value="<%=rs("wear_addr")%>">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Email :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <input id="wear_email"  name="wear_email" maxlength=98 size=80 value="<%=rs("wear_email")%>">
        </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td width="12%"><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>C</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>ontent :</FONT><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </FONT></td>
      <td width="83%"><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3> 
        <textarea name="wear_memo" id="wear_memo" cols="70" rows="30"><%=rs("wear_memo")%></textarea>
        </font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3><font color="#FF0000">* </font></font></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#ffffff size=3>&nbsp; </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td colspan="2"> <input type="submit" value="edit" name="option" id="option">
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td height="50" colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<%
rs.close
set rs=nothing
%>