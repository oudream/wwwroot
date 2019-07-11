<script language="JavaScript">

function addnum(value_1,value_2,selObj){ 
selObj.value=value_1*value_2;
}

function checkform(a,b,c,d,e,f,g,h)
{
	if (form_administrator.gid.value.length == 0) {
		alert("请选取这个商品的种类 ");
		form_administrator.gid.focus();
		return false;
		}
	if (form_administrator.pyear.value.length == 0) {
		alert("请输入年份 ");
		form_administrator.pyear.focus();
		return false;
		}
	if(! TFNumber(form_administrator.pyear.value)) {
		alert("请用数字填写年份");
		form_administrator.pyear.focus();
		return false;
		}
	if (form_administrator.pyear.value > 2005 || form_administrator.pyear.value < 2003) {
		alert("请输入正确的年份, 必须在 2004 至 2005 . ");
		form_administrator.pyear.focus();
		return false;
		}
		
	if (form_administrator.pmonth.value.length == 0) {
		alert("请输入月份 ");
		form_administrator.pmonth.focus();
		return false;
		}
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

	if (form_administrator.pday.value.length == 0) {
		alert("请输入日期 ");
		form_administrator.pday.focus();
		return false;
		}
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

	if( a != 0 ) {
	if(! isNumber(form_administrator.mount_1.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_1.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_1.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_1.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_1.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_1.focus();
		return false;
		}
		}
		

	if( b != 0 ) {
	if(! isNumber(form_administrator.mount_2.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_2.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_2.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_2.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_2.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_2.focus();
		return false;
		}
		}
		

	if( c != 0 ) {
	if(! isNumber(form_administrator.mount_3.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_3.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_3.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_3.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_3.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_3.focus();
		return false;
		}
		}
		

	if( d != 0 ) {
	if(! isNumber(form_administrator.mount_4.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_4.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_4.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_4.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_4.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_4.focus();
		return false;
		}
		}
		

	if( e != 0 ) {
	if(! isNumber(form_administrator.mount_5.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_5.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_5.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_5.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_5.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_5.focus();
		return false;
		}
		}
		

	if( f != 0 ) {
	if(! isNumber(form_administrator.mount_6.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_6.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_6.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_6.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_6.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_6.focus();
		return false;
		}
		}
		

	if( g != 0 ) {
	if(! isNumber(form_administrator.mount_7.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_7.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_7.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_7.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_7.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_7.focus();
		return false;
		}
		}
		

	if( h != 0 ) {
	if(! isNumber(form_administrator.mount_8.value)) {
		alert("请输入正确的数量.");
		form_administrator.mount_8.focus();
		return false;
		}
	if(! isNumber(form_administrator.price_8.value)) {
		alert("请输入正确的单价.");
		form_administrator.price_8.focus();
		return false;
		}
	if(! isNumber(form_administrator.counter_8.value)) {
		alert("请输入正确的金额.");
		form_administrator.counter_8.focus();
		return false;
		}
		}
		
	if (form_administrator.operationer_id.value == 0) {
		alert("请选择本公司经手人 ");
		form_administrator.operationer_id.focus();
		return false;
		}
	return true;
}

function isNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (!((name.charAt(i) >= "0" && name.charAt(i) <= "9") || name.charAt(i) == "."))
			return false;
	}
	return true;
}

function TFNumber(name)
{
	if (name.length == 0) 
		return false;
	for(i = 0; i < name.length; i++) {
		if (name.charAt(i) < "0" || name.charAt(i) > "9")
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
if session("subs_pass_type")="in" then 
passstr="进"
passstr1="出"
end if
if session("subs_pass_type")="out" then 
passstr="出"
passstr1="进"
end if
%>
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                   ERROR MSGS
<!-------------------------------------------------------------------------------------------------------------------- -->
<FORM action="subs_pass_in.asp" method="post" name="form_administrator" id="form_administrator" onSubmit="return checkform(<%=session("subs_id_1")%>,<%=session("subs_id_2")%>,<%=session("subs_id_3")%>,<%=session("subs_id_4")%>,<%=session("subs_id_5")%>,<%=session("subs_id_6")%>,<%=session("subs_id_7")%>,<%=session("subs_id_8")%>);">
  <TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
    <TBODY>
      <TR align="center">
        <TD height="60" colSpan=2><br><font size="6" face="宋体">数 信 <%=passstr%> 货 单</font> 
        </TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td width="50%"><font color="red" size="3"><%=passstr1%></font><font size="3">货单位 :
			  <select name="gid" id="gid" class="display_dropdown">
		  <option value="" selected> 请选择……</option>
          <%
psql="select * from guest"
Set prs=Server.CreateObject("ADODB.RecordSet")
prs.Open psql,conn,1,1
if not ( prs.eof or prs.bof ) then
do while not prs.eof
%>
          <option value=<%=prs("gid")%>><%=prs("simname")%></option>
<%
prs.movenext
loop
end if
prs.close
set prs=nothing
%>
        </select>
</font></td>
              <td width="50%" align="right"><table width="300" border="0" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td align="right">
					  <input name="pyear" id="pyear" type="text" size="5" maxlength="4" value="<%=year(now)%>">
                      年 
					  <input name="pmonth" id="pmonth" type="text" size="3" maxlength="2" value="<%=month(now)%>">
                      月 
                      <input name="pday" id="pday" type="text" size="3" maxlength="2" value="<%=day(now)%>">
                      日
					</td>
                  </tr>
                </table></td>
            </tr>
          </table></TD>
      </TR>
      <TR align="center"> 
        <TD height="50" colSpan=2><table width="100%" border="1" cellspacing="0" cellpadding="0">
            <tr align="center"> 
              <td width="200" height="30">商品名称及型号/规格</td>
              <td width="30">单位</td>
              <td width="80">数量</td>
              <td width="80">单价</td>
              <td width="80">金额</td>
              <td width="130">备注</td>
            </tr>
<%
Set rs=Server.CreateObject("ADODB.RecordSet")
Set brs=Server.CreateObject("ADODB.RecordSet")
if session("subs_id_1")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_1")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_1" id="mount_1" type="text" size="8" maxlength="8"></td>
              <td><input name="price_1" id="price_1" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_1" id="counter_1" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_1.value,form_administrator.price_1.value,this);"></td>
              <td><input name="memo_1" id="memo_1" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>



<%
if session("subs_id_2")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_2")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_2" id="mount_2" type="text" size="8" maxlength="8"></td>
              <td><input name="price_2" id="price_2" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_2" id="counter_2" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_2.value,form_administrator.price_2.value,this);"></td>
              <td><input name="memo_2" id="memo_2" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_3")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_3")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_3" id="mount_3" type="text" size="8" maxlength="8"></td>
              <td><input name="price_3" id="price_3" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_3" id="counter_3" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_3.value,form_administrator.price_3.value,this);"></td>
              <td><input name="memo_3" id="memo_3" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_4")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_4")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_4" id="mount_4" type="text" size="8" maxlength="8"></td>
              <td><input name="price_4" id="price_4" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_4" id="counter_4" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_4.value,form_administrator.price_4.value,this);"></td>
              <td><input name="memo_4" id="memo_4" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_5")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_5")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_5" id="mount_5" type="text" size="8" maxlength="8"></td>
              <td><input name="price_5" id="price_5" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_5" id="counter_5" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_5.value,form_administrator.price_5.value,this);"></td>
              <td><input name="memo_5" id="memo_5" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_6")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_6")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_6" id="mount_6" type="text" size="8" maxlength="8"></td>
              <td><input name="price_6" id="price_6" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_6" id="counter_6" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_6.value,form_administrator.price_6.value,this);"></td>
              <td><input name="memo_6" id="memo_6" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_7")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_7")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_7" id="mount_7" type="text" size="8" maxlength="8"></td>
              <td><input name="price_7" id="price_7" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_7" id="counter_7" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_7.value,form_administrator.price_7.value,this);"></td>
              <td><input name="memo_7" id="memo_7" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>

<%
if session("subs_id_8")<>0 then
sql="select * from subs where subs_id=" & session("subs_id_8")
rs.open sql,conn,1,1
if not(rs.eof or rs.bof) then 
bsql="select * from brand where brand_id=" & rs("brand_id")
brs.open bsql,conn,1,1
%>
            <tr align="center"> 
              <td>
<%
if not(brs.eof or brs.bof) then response.Write(brs("bname"))
response.Write(" -- " & rs("code"))			 
brs.close
%>
			  </td>
              <td>
<%
response.Write(rs("suit"))			 
%>
			  </td>
              <td><input name="mount_8" id="mount_8" type="text" size="8" maxlength="8"></td>
              <td><input name="price_8" id="price_8" type="text" size="8" maxlength="8"></td>
              <td><input name="counter_8" id="counter_8" type="text" size="8" maxlength="8" onFocus="addnum(form_administrator.mount_8.value,form_administrator.price_8.value,this);"></td>
              <td><input name="memo_8" id="memo_8" type="text" size="15" maxlength="100"></td>
            </tr>
<%
end if
rs.close
end if
%>
          </table>
          <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="35%" height="30"><font color="red"><%=passstr1%></font>货单位经手人<input name="gcontacter" id="gcontacter" type="text" size="10" maxlength="10"></td>
              <td width="40%">本公司<%=passstr%>货经手人<select name="operationer_id" id="operationer_id" class="display_dropdown">
                  <option value=0 selected> 请选择……</option>
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
              <td width="25%">填表人<input name="tablewrite" id="tablewrite" type="text" size="10" maxlength="10" value="<%=session("user_name")%>" disabled></td>
            </tr>
          </table></TD>
      </TR>
      <TR> 
        <TD height="50" colSpan=2> <P><STRONG><FONT 
                        face="Verdana, Arial, Helvetica, sans-serif" 
                        color=#000066 size=2> 
            <INPUT type=submit id="option" name="option" value="add">
			&nbsp;&nbsp;&nbsp;<input type="reset" name="cancelit" id="cancelit" value="清空">
			&nbsp;&nbsp;&nbsp;<input type="button" name="cancelit" id="cancelit" value="重新选取商品" onClick="location='subs_select.asp?passtype=<%=session("subs_pass_type")%>'">
			&nbsp;&nbsp;&nbsp;<input type="button" name="cancelit" id="cancelit" value="取消所有操作" onClick="location='adminlogin.asp'">
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
