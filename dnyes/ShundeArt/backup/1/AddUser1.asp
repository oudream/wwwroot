<!--#include file="conn.asp"-->
<!--#include file="config.asp"-->

<%if reg="1" then%>
	<html>
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=jjgn%> - 会员注册</title>
	<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
	<link rel="stylesheet" type="text/css" href="site.css">
	<script LANGUAGE="javascript">
	<!--
	function FrmAddLink_onsubmit() {
	var i, n;
	if (document.FrmAddLink.username.value=="")
		{
		  alert("对不起，请输入您的用户名！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length < 2)
		{
		  alert("您的用户名能不能长一点！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.length > 30)
		{
		  alert("您的用户名太长了吧！")
		  document.FrmAddLink.username.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value.indexOf('`') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('~') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('!') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('@') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('#') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('$') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('%') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('^') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('&') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('*') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('(') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(')') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('+') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('{') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('}') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('|') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('[') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(']') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('\\') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(';') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(':') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('>') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('<') >= 0 ||
	          document.FrmAddLink.username.value.indexOf(',') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('?') >= 0 ||
	          document.FrmAddLink.username.value.indexOf('/') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('\'') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('"') >= 0 || 
	          document.FrmAddLink.username.value.indexOf(' ') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('=') >= 0 || 
	          document.FrmAddLink.username.value.indexOf('%') >= 0
	         )  
	          {
	            alert("用户名中包含无效字符，请重新选择用户名！"); 
	            document.FrmAddLink.username.focus();
	            return false;
	          }
	else if (document.FrmAddLink.passwd.value=="")
		{
		  alert("对不起，请您输入密码！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length < 4)
		{
		  alert("为了安全，您的密码应该长一点！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd.value.length > 16)
		{
		  alert("您的密码太长了吧！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.username.value==document.FrmAddLink.passwd.value)
		{
		  alert("为了安全，用户名与密码不应该相同！")
		  document.FrmAddLink.passwd.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value=="")
		{
		  alert("对不起，请您输入验证密码！")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.passwd2.value !== document.FrmAddLink.passwd.value)
		{
		  alert("对不起，您两次输入的密码不一致！")
		  document.FrmAddLink.passwd2.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value=="")
		{
		  alert("对不起，请您输入提示问题！")
		  document.FrmAddLink.question.focus()
		  return false
		 }
	else if (document.FrmAddLink.answer.value=="")
		{
		  alert("对不起，请您输入问题答案！")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.question.value==document.FrmAddLink.answer.value)
		{
		  alert("为了安全，提示问题与问题答案不应该相同！")
		  document.FrmAddLink.answer.focus()
		  return false
		 }
	else if (document.FrmAddLink.fullname.value=="")
		{
		  alert("对不起，请输入您的真实姓名！")
		  document.FrmAddLink.fullname.focus()
		  return false
		 }
	else if (document.FrmAddLink.depid.value=="")
		{
		  alert("对不起，请选择您的工作单位！")
		  document.FrmAddLink.depid.focus()
		  return false
		 }
	else if (document.FrmAddLink.sex.value=="")
		{
		  alert("对不起，请选择您的性别！")
		  document.FrmAddLink.sex.focus()
		  return false
		 }
	else if (document.FrmAddLink.tel.value=="")
		{
		  alert("对不起，请输入您的联系电话！")
		  document.FrmAddLink.tel.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value=="")
		{
		  alert("对不起，请输入您的电子邮件！")
		  document.FrmAddLink.email.focus()
		  return false
		 }
	else if (document.FrmAddLink.email.value.indexOf("@",0)== -1||document.FrmAddLink.email.value.indexOf(".",0)==-1)
		  {
		  alert("对不起，您输入的电子邮件有误！")
		  document.FrmAddLink.email.focus()
		  return false
		 }
	}
	
	//Function to open pop up window
	function openWin(theURL,winName,features) {
	  	window.open(theURL,winName,features);
	}
	//-->
	</script>
	
	<//生日选择日期处理开始>
<SCRIPT language=JavaScript src="inc/User_Info_Modify.js"></SCRIPT>
	<//生日选择日期处理结束>
	
	</head>
	<!--#include file=User_Top.asp-->
	<body topmargin="0">
    
	    
	<form align="center" method="post" name="FrmAddLink" LANGUAGE="javascript" onsubmit="return FrmAddLink_onsubmit()" action="saveuser.asp">
  <table align=center border="1" width="350" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="<%=border%>">
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle"><span class="smallFont">┊<strong> 会 员 注 册 </strong>┊</span></td>
    </tr>
    <tr> 
      <td height="20" colspan="2" align="center" valign="middle">为了使您能正常使用本系统，请详细填写每一项资料</td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right" valign="middle"> <div align="center"><span class="smallFont">用 
          户 名：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="username" size="26" class="smallInput" maxlength="30" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的用户名，注册成功后将用此用户名登录本系统。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">密&nbsp;&nbsp;&nbsp;码：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password"  title="请在这里填写您的登录密码，在登录时本系统将验证您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">验证密码：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="passwd2" size="26" class="smallInput" maxlength="15" style="font-family: 宋体; font-size: 9pt" type="password" title="请在这里填写您的验证密码，必须与上面的密码一致，主要是防止您的错误输入。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">提示问题：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="question" size="26" class="smallInput" maxlength="50" style="font-family: 宋体; font-size: 9pt" type="text" title="请在这里填写您的提示问题，如果您忘记了密码，可以利用此功能来找回您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">问题答案：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="answer" size="26" class="smallInput" maxlength="50" style="font-family: 宋体; font-size: 9pt" type="text" title="请在这里填写您的提示问题的答案，如果您忘记了密码，可以利用此功能来找回您的密码。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">真实姓名：</span></div></td>
      <td width="264" height="20"> <div align="left"> 
          <input name="fullname" size="26" class="smallInput" maxlength="10" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的真实姓名。">
        </div></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">工作单位：</span></div></td>
      <td width="264" height="20"> 
        <%set rstype=createobject("adodb.recordset")
					sql="select * from "& db_Dep_Table &" order by id"
					rstype.Open sql,conn,1,3%>
        <select name="depid" style="font-family: 宋体; font-size: 9pt" title="请在这里选择您的工作单位。">
          <option value="">请选择工作单位</option>
          <%do while not rstype.EOF%>
          <option value="<%=rstype("id")%>"><%=rstype("depname")%>==<%=rstype("deptype")%></option>
          <%rstype.MoveNext
					loop
					set rstype=nothing%>
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">性&nbsp;&nbsp;&nbsp; 
          别：</span></div></td>
      <td width="264" height="20"> <select size="1" name="sex" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的性别。">
          <option selected value="">请选择性别</option>
          <%if db_Sex_Select = "FeiTium" then%>
          <option value="先生">先生</option>
          <option value="女士">女士</option>
          <option value="保密">保密</option>
          <%else%>
          <%if db_Sex_Select = "Number" then%>
          <option value="1">先生</option>
          <option value="0">女士</option>
          <%end if%>
          <%end if%>
        </select> </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">生&nbsp;&nbsp;&nbsp;日：</span></div></td>
      <td width="264" height="40"> 
        <%if db_Birthday_Select = "FeiTium" then%>
        <div align="left"> 
          <select size="1" name="birthyear" style="font-family: 宋体; font-size: 9pt">
            <%for i=1950 to 2004%>
            <option value="<%=i%>"><%=i%></option>
            <%next%>
          </select>
          年 
          <select size="1" name="birthmonth" style="font-family: 宋体; font-size: 9pt">
            <%for i=1 to 12%>
            <option value="<%=i%>"><%=i%></option>
            <%next%>
          </select>
          月 
          <select size="1" name="birthday" style="font-family: 宋体; font-size: 9pt">
            <%for i=1 to 31%>
            <option value="<%=i%>"><%=i%></option>
            <%next%>
          </select>
          日 </div>
        <%else%>
        <%if db_Birthday_Select = "Text" then%>
        <div align="left"> 
          <INPUT  name=birthday onfocus="show_cele_date(birthday,'','',birthday)" value="<%=year(now())-18%>-<%=month(now())%>-<%=day(now())%>" >
        </div>
        <%end if%>
        <%end if%>
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">联系电话：</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="tel" size="26" class="smallInput" maxlength="100" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的联系电话，以便我们与您联系。"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">电子邮件：</span></div></td>
      <td width="264" height="20" valign="middle"> <input name="email" size="26" class="smallInput" maxlength="100" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的电子邮件地址。"></td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">个人照片：</span></div></td>
      <td width="264" height="20" valign="middle"> 
        <%if UserTableType = "Dvbbs"	then ''显示用户头像，加'bbs/'前缀路径,使图文系统直接显示定向到论坛头像%>
        bbs/ <input id="photo" name="photo" value="Images/userface/image1.gif" size="26" class="smallInput" maxlength="255" style="font-family: 宋体; font-size: 9pt" title="个人照片。您可以上传自己的照片，也可以直接填写您的网上照片的地址。"> 
        <p align=center> <img src="<%=BbsPath%>Images/userface/image1.gif" border="0"> 
        </P>
        <%else%>
        <input id="photo" name="photo" value="images/nopic.gif" size="26" class="smallInput" maxlength="255" style="font-family: 宋体; font-size: 9pt" title="个人照片。您可以上传自己的照片，也可以直接填写您的网上照片的地址。"> 
        <p align=center> <img src="images/nopic.gif" border="0"> </P>
        <%end if%>
      </td>
    </tr>
    <tr> 
      <td width="81" height="20" align="right"> <div align="center"><span class="smallFont">自我介绍：</span></div></td>
      <td width="264" height="32" valign="middle"> <textarea rows="5" name="content" cols="29" style="font-family: 宋体; font-size: 9pt" title="请在这里填写您的个人介绍。"></textarea> 
      </td>
    </tr>
  </table>
		<div align="center">
				<br>
					<input name="purview" type="hidden" value="1">
					<input name="oskey" type="hidden" value="selfreg">
					<input name="reglevel" type="hidden" value="1">
					<%if fabiao=1 then%>
						<input name="jingyong" type="hidden" value="0">
					<% else %>
					    <input name="jingyong" type="hidden" value="1">
					<%end if%>
					<input type="submit" value=" 确 定 " name="cmdOk" class="buttonface" style="font-family: 宋体; font-size: 9pt;">
					&nbsp; 
					<input type="reset" value=" 重 填 " name="cmdReset" class="buttonface" style="font-family: 宋体; font-size: 9pt;" >
		</div>
	</form>
	</body>
	</html>
	<!--#include file=User_Bottom.asp-->
<%else%>
	<script language=javascript>
	alert("对不起，用户注册功能已被管理员关闭！")
	</script>
	<body onload="javascript:window.close()"></body>
<%end if%>