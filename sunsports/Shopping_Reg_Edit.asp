<script language="JavaScript">

function checkform()
{
	if(form_administrator.email.value.length > 0){
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
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
<!--#include file="Top_sun.asp"-->
<table width="776" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="181" valign="top"><table width="100%" border="0" cellpadding="3" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td> 
            <!-- ========================================================================================================		  

													cup table ============star

 ======================================================================================================== -->
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="League.asp"><b><font color="#1F4A71">LEAGUE MATCHES</font></b></a><font color="#1F4A71">+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Cup.asp"><b><font color="#1F4A71">CUP MATCHES </font></b></a><font color="#1F4A71">&nbsp;+ 
                  +</font></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Friendlies.asp"><b><font color="#1F4A71">FRIENDLY 
                  MATCHES </font></b></a></td>
              </tr>
            </table>
            <table width="169" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td width="163"><a href="TeamProfile.asp"><b><font color="#1F4A71">TEAMPORFILE</font></b></a> 
                  <font color="#1F4A71">&nbsp;+ +</font> </td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="TeamProfile.asp"><b><font color="#1F4A71">FIELDS 
                  BOOKING</font></b></a></td>
              </tr>
            </table>
            <table width="170" border="0" align="center" cellpadding="0" cellspacing="3">
              <tr> 
                <td><a href="Shopping_sort_view.asp"><b><font color="#1F4A71">SUNSPORTS 
                  PRODUCTS </font></b></a> <font color="#1F4A71">+ +</font></td>
              </tr>
            </table>
            <!-- ========================================================================================================		  

													cup table ============stop

 ======================================================================================================== -->
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table></td>
    <td width="20">&nbsp; </td>
    <td valign="top"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="25">&nbsp;</td>
        </tr>
      </table>
<%
sql="select * from player where pid=" & session("user_pid")
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
%>
      <table width="90%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#6184A3">
        <FORM action="Shopping_Reg_Save.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform();">
          <tr align="center" bgcolor="#FFFFFF"> 
            <td height="30" colspan="2"><B>Manage Your Infomation</B></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td width="28%">User ID :</td>
            <td width="72%"><%=rs("uid")%>&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Full Name:</td>
            <td><input name="uname" type="text" id="uname" size="30" maxlength="30" class="form" value="<%=rs("name")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Telphone:</td>
            <td><input name="tel" type="text" id="tel" size="30" maxlength="20" class="form" value="<%=rs("tel")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Email:</td>
            <td><input name="email" type="text" id="email" size="30" maxlength="30" class="form" value="<%=rs("email")%>"></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>ZIP:</td>
            <td><input name="zip" type="text" id="zip" size="10" maxlength="10" class="form" value="<%=rs("zip")%>"></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>Address:</td>
            <td><input name="contact" type="text" id="contact" size="40" maxlength="100" class="form" value="<%=rs("contact")%>">
              <font color="#FF0000">*</font></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td>&nbsp;</td>
            <td> 
			<input type="hidden" name="option" id="option" value="edit"> 
			<input type="submit" name="SING UP" value="Edit" class="button"> 
            </td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"> # <font color="#FF0000">*</font>Required Fields<br>
              <br>
              <br>
              # The infomation you fill in just use for connection and delivering 
              products.Thanks!<br>
              <br>
            </td>
          </tr>
        </form>
      </table>
<%
rs.close
set rs=nothing
%>	  
	  </td>
    <td width="20" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="24">&nbsp;</td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td>&nbsp;</td>
          <td height="280" background="images/10x1_blue.gif">&nbsp;</td>
        </tr>
      </table></td>
    <td width="181" valign="top">
          <!--#include file="User_l.asp" -->
	 <br> <!--#include file="Shopping_Search.asp" -->
	 <br> 	
	</td>
  </tr>
</table>
<!--#include file="Copyright.asp"-->