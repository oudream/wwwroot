<script language="JavaScript">
function gettarg(targ,selObj,restore){ 
  eval(targ+".location='subs_select_in7.asp?option="+'ȡ��ѡ��'+"&sort_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
function getbrand(targ,selObj,restore,sort_id){ 
  eval(targ+".location='subs_select_in7.asp?option="+'ȡ��ѡ��'+"&sort_id="+sort_id+"&brand_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
function getsubs(targ,selObj,restore,sort_id,brand_id){ 
  eval(targ+".location='subs_select_in7.asp?option="+'ȡ��ѡ��'+"&sort_id="+sort_id+"&subs_id="+selObj.options[selObj.selectedIndex].value+"&brand_id="+brand_id+"'");
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
sort_id=request("sort_id")
if sort_id="" then sort_id=0
brand_id=request("brand_id")
if brand_id="" then brand_id=0
subs_id=request("subs_id")
if subs_id="" then subs_id=0

optionstr2=""
optionstr3="disabled"
zoption=request("option")
if zoption="ȷ��ѡ��>>" then
session("subs_id_7")=subs_id
optionstr2="disabled"
optionstr3=""
end if
if zoption="ȡ��ѡ��" then
session("subs_id_7")=0
optionstr2=""
optionstr3="disabled"
end if
%>
<form name="form_administrator" method="post" action="subs_select_in7.asp" onSubmit="return checkform();">
  <table width="97%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#EBEBEB">
    <tr> 
      <td width="100" bgcolor="#FFFFFF"> <select name="sort_id" id="sort_id" class="display_dropdown"  onChange="gettarg('self',this,0)">
          <%
Set prs=Server.CreateObject("ADODB.RecordSet")
if sort_id=0 then
response.Write("<option value='' selected>��ѡ�񡭡�</option>")
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
        </select> </td>
      <td width="100" bgcolor="#FFFFFF"> 
	  <select name="brand_id" id="brand_id" class="display_dropdown"  onChange="getbrand('self',this,0,<%=sort_id%>)">
          <%
if brand_id=0 then
response.Write("<option value='' selected>��ѡ�񡭡�</option>")
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
      <td width="150" bgcolor="#FFFFFF"> 
	  <select name="subs_id" id="subs_id" class="display_dropdown"  onChange="getsubs('self',this,0,<%=sort_id%>,<%=brand_id%>)">
          <%
if subs_id=0 then
response.Write("<option value='' selected>��ѡ�񡭡�</option>")
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
      <td width="240" align="center" bgcolor="#FFFFFF">
	  <input type="submit" name="option" id="option" value="ȷ��ѡ��>>" <%=optionstr2%>> &nbsp;&nbsp;&nbsp;
	  <input type="submit" name="option" id="option" value="ȡ��ѡ��" <%=optionstr3%>></td>
    </tr>
  </table>
</form>
