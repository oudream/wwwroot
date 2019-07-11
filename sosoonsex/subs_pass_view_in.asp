<script language="JavaScript">
function gettarg(targ,selObj,restore){ 
  eval(targ+".location='subs_pass_view_in.asp?sort_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
function getbrand(targ,selObj,restore,sort_id){ 
  eval(targ+".location='subs_pass_view_in.asp?sort_id="+sort_id+"&brand_id="+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
function getsubs(targ,selObj,restore,sort_id,brand_id){ 
  eval(targ+".location='subs_pass_view_in.asp?sort_id="+sort_id+"&subs_id="+selObj.options[selObj.selectedIndex].value+"&brand_id="+brand_id+"'");
  if (restore) selObj.selectedIndex=0;
}

function checkform()
{
	if (form_administrator.pmonth.value.length > 0) {
	if(! TFNumber(form_administrator.pmonth.value)) {
		alert("请用数字填写月份");
		form_administrator.pmonth.focus();
		return false;
		}
	if (form_administrator.pmonth.value > 12 || form_administrator.pmonth.value < 1) {
		alert("请输入正确的月份, 必须在 1 至 12 . ");
		form_administrator.pmonth.focus();
		return false;
		}
		}

	if (form_administrator.pday.value.length > 0) {
	if(! TFNumber(form_administrator.pday.value)) {
		alert("请用数字填写日期");
		form_administrator.pday.focus();
		return false;
		}
	if (form_administrator.pday.value > 31 || form_administrator.pday.value < 1) {
		alert("请输入正确的日期, 必须在 1 至 31 . ");
		form_administrator.pday.focus();
		return false;
		}
		}
	return true;
}

function TFNumber(name)
{
	for(i = 0; i < name.length; i++) {
		if (name.charAt(i) < "0" || name.charAt(i) > "9")
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
%>

<form name="form_administrator" id="form_administrator" method="post" action="subs_pass_view.asp" target="_parent" onSubmit="return checkform();">
        
  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">
<tr> 
      <td height="30" colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="100"><select name="sort_id" id="sort_id" class="display_dropdown"  onChange="gettarg('self',this,0)">
                <%
Set prs=Server.CreateObject("ADODB.RecordSet")
if sort_id=0 then
response.Write("<option value='' selected>种类-->选择</option>")
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
              </select></td>
            <td width="140"><select name="brand_id" id="brand_id" class="display_dropdown"  onChange="getbrand('self',this,0,<%=sort_id%>)">
                <%
if brand_id=0 then
response.Write("<option value='' selected>品牌-->选择</option>")
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
            <td width="160"><select name="subs_id" id="subs_id" class="display_dropdown"  onChange="getsubs('self',this,0,<%=sort_id%>,<%=brand_id%>)">
                <%
if subs_id=0 then
response.Write("<option value='' selected>型号/规格-->选择</option>")
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
          </tr>
        </table></td>
      <td width="150" height="30"><select name="gid" id="select9" class="display_dropdown">
          <option value=0 selected>客户-->选择</option>
          <%
gsql="select * from guest"
Set grs=Server.CreateObject("ADODB.RecordSet")
grs.Open gsql,conn,1,1
if not ( grs.eof or grs.bof ) then
do while not grs.eof
%>
          <option value=<%=grs("gid")%>><%=grs("simname")%></option>
          <%
grs.movenext
loop
end if
grs.close
set grs=nothing
%>
        </select> </td>
      <td width="50">&nbsp;</td>
    </tr>
    <tr> 
      <td width="150" height="30"> 
        <select name="subspass_type" id="select8" class="display_dropdown">
          <option value="" selected>流通类型-->选择</option>
          <option value="in">进货</option>
          <option value="out">出货</option>
        </select></td>
      <td width="200" height="30">日期 
        <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2">
        月 
        <input name="pday" id="pday" type="text" size="3" maxlength="2">
        日 </td>
      <td width="200" height="30">
        <select name="operationer_id" id="operationer_id" class="display_dropdown">
          <option value=0 selected>本公司经手人-->选择</option>
          <%
psql="select * from user"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("id")%>><%=prs("uname")%></option>
          <%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
        </select></td>
      <td width="50"><input type="submit" name="Submit" value="查询>>"> </td>
    </tr>
    
  </table>
      </form> 
