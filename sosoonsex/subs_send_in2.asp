<script language="JavaScript">

function getbrand(targ,selObj,restore,sort_id){ 
  eval(targ+".location='subs_send_in2.asp?option="+'取消选择'+"&sort_id="+sort_id+"&brand_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function checkform()
{
	if (form_administrator.subs_id.value.length == 0) {
		alert("Please enter a valid Datetime. ");
		form_administrator.subs_id.focus();
		return false;
		}
	return true;
}
</script>

<!--#include file="adminconn.asp" -->

<%
sort_id=52
brand_id=request("brand_id")
if brand_id="" then brand_id=0
subs_id=request("subs_id")
if subs_id="" then subs_id=0

optionstr2=""
optionstr3="disabled"
zoption=request("option")
if zoption="确认选择>>" then
session("subs_id_2")=subs_id
optionstr2="disabled"
optionstr3=""
end if
if zoption="取消选择" then
session("subs_id_2")=0
optionstr2=""
optionstr3="disabled"
end if
%>
<form name="form_administrator" method="post" action="subs_send_in2.asp" onSubmit="return checkform();">
  <table width="97%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#EBEBEB">
    <tr> 
      <td width="150" bgcolor="#FFFFFF"> 
	  <select name="brand_id" id="brand_id" class="display_dropdown"  onChange="getbrand('self',this,0,<%=sort_id%>)">
<%
Set prs=Server.CreateObject("ADODB.RecordSet")
if brand_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from brand where brand_id=" & brand_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("brand_id")%> selected><%=prs("bname")%></option>
          <%
end if
prs.close
end if
%>
          <%
psql="SELECT * FROM brand where brand_id in ( select brand_id from sb where sort_id="&sort_id&" ) and brand_id<>" & brand_id &" order by bname"
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
%>
        </select></td>
      <td width="250" bgcolor="#FFFFFF"> 
	  <select name="subs_id" id="subs_id" class="display_dropdown">
          <%
if subs_id=0 then
response.Write("<option value='' selected>请选择……</option>")
else
psql="select * from subs where subs_id=" & subs_id
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
%>
          <option value=<%=prs("subs_id")%> selected><%=prs("code")%></option>
          <%
end if
prs.close
end if
%>
          <%
psql="select * from subs where brand_id="&brand_id&" and sort_id=" & sort_id &" and subs_id<>" & subs_id & " order by code"
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("subs_id")%>><%=prs("code")%></option>
          <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
        </select></td>
      <td width="200" align="center" bgcolor="#FFFFFF">
	  <input type="submit" name="option" id="option" value="确认选择>>" <%=optionstr2%>>&nbsp;
	  <input type="submit" name="option" id="option" value="取消选择" <%=optionstr3%>></td>
    </tr>
  </table>
</form>
