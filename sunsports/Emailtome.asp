<script language="JavaScript">

function checkform()
{
	if (form_administrator.name.value.length == 0) {
		alert("Please enter a valid Name. ");
		form_administrator.uname.focus();
		return false;
		}
	if(! isLetter(form_administrator.name.value)) {
		alert("Please enter a valid Name. ");
		form_administrator.name.focus();
		return false;
		}
	if(! isMail(form_administrator.email.value)) {
		alert('Please enter a valid EMAIL address');
		form_administrator.email.focus();
		return false;
		}
	if (form_administrator.subject.value.length == 0) {
		alert("Please enter a valid title. ");
		form_administrator.subject.focus();
		return false;
		}
	if(! isLetter(form_administrator.subject.value)) {
		alert("Please enter a valid title. ");
		form_administrator.subject.focus();
		return false;
		}
	if (form_administrator.mailtext.value.length == 0) {
		alert("Please enter a valid content. ");
		form_administrator.mailtext.focus();
		return false;
		}
	if(! isLetter(form_administrator.mailtext.value)) {
		alert("Please enter a valid content. ");
		form_administrator.mailtext.focus();
		return false;
		}
	return true;
}

function isLetter(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charCodeAt(i) > 128)
			return false;
	}
	return true;
}

function isNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if(name.charAt(i) < "0" || name.charAt(i) > "9")
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

<!--#include file="conn/conn.asp" -->
<%
if request("option")="add" then
uname=trim(request("name"))
uname=replace(uname,"'","''")
email=trim(request("email"))
subject=trim(request("subject"))
subject=replace(subject,"'","''")
mailtext=trim(request("mailtext"))
mailtext=replace(mailtext,"'","''")
sql="insert into mail (name,email,subject,mailtext,inserttime) values ('"&uname&"','"&email&"','"&subject&"','"&mailtext&"','"&now()&"')"
conn.Execute(sql)
response.Write("<SCRIPT LANGUAGE=JavaScript> alert(' Your Email is send , we will reply you as quickly as possible . ');</SCRIPT>")
end if
%>
<FORM action="emailtome.asp" method=post name="form_administrator" id="form_administrator" onSubmit="return checkform();">
                        
  <TABLE width="498" border=0 align="center" cellPadding=5 
                        cellSpacing=1 bgColor=#CCCCCC>
    <TBODY>
      <TR bgColor=#ffffff> 
                                        
        <TD 
                              width="25%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Dear 
          £º</TD>
        <TD bgColor=#ffffff width="75%"> 
          <INPUT id="name" name="name" size=30 maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        
        <TD 
                              width="25%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Your 
          Email £º</TD>
        <TD bgColor=#ffffff width="75%"> 
          <INPUT id="email" name="email" size=30 maxlength="50"></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        
        <TD 
                              width="25%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5> 
          Title £º</TD>
        <TD bgColor=#ffffff width="75%"><INPUT name=subject id="subject" size=50 maxlength="100"></TD>
      </TR>
      <TR bgColor=#ffffff> 
                                        
        <TD 
                              width="25%" align="center" borderColor=#c9c9c9 bgColor=#f5f5f5>Content 
          £º</TD>
                                        <TD bgColor=#ffffff width="75%"><textarea cols=49 name=mailtext id="mailtext" rows=10></textarea> 
                                        </TD>
      </TR>
      <TR bgColor=#ffffff>
        <TD bgColor=#f5f5f5 borderColor=#c9c9c9>&nbsp;</TD>
        <TD bgColor=#ffffff>
		<input type=submit id="options" name="options" value="OK AND SENT &gt;&gt;">
		<input type="hidden" id="option" name="option" value="add">
		</TD>
      </TR>
    </TBODY>
  </TABLE>
</FORM>
