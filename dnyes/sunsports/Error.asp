<%
errid=request("err")
if errid="" then err="001"
%>
<table width="100%" align="center">
  <tr align="center"> 
    <td align="center"> <br>
      === <b>operation error</b> ===<br> <br> <br> </td>
  </tr>
  <tr> 
    <td> <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td width="50" valign="top"><br> <br> <br> </td>
          <td valign="top"><p> <br>
              <%if errid="001" then%>
              Operation of you have the following problems,suggestion under please 
              read carefully: You are not the form form information referred from 
              this server! Your data stock is problematic,please contact website 
              builder! !</p>
            <p> Operate the right put forward and appeal further!</p>
            <p> It is to the letter,please if there is any doubt: ：<a href="mailto:<%=Application("sunsports_email")%>"><%=Application("sunsports_email")%></a></p>
            <p><br>
              <%end if%>
              <%if errid="002" then%>
              Sorry ! <font color=red>You have not loginde</font></p>
            <form action="login.asp" method="post">
              Userid : 
              <input type="text" name="uid2" size="16" maxlength="20" class="form">
              &nbsp; 
              <input type="submit" name="Submit" class=botton value="login">
              <input type="hidden" name="url" value="<%=request("return_url")%>">
              <br>
              Password : 
              <input type="password" name="pwd" size="16" maxlength="20" class="form">
              &nbsp; 
              <input type="submit" name="Submit2" class=botton value="cancel">
            </form>
            <p><br>
              <br>
              <br>
              It is to the letter,please if there is any doubt: ：<a href="mailto:<%=Application("sunsports_email")%>"><%=Application("sunsports_email")%></a></p>
            <p><br>
              <%end if%>
              <%if errid="003" then%>
              您不是从本服务器上提交的表单信息！！！<br>
              <br>
              你的数据库存有问题，请与网站建设者联系！！<br>
              <br>
              操作提出进一步上诉的权利！！<br>
              <br>
              <br>
              <br>
              It is to the letter,please if there is any doubt: ：<a href="mailto:<%=Application("sunsports_email")%>"><%=Application("sunsports_email")%></a><br>
              <%end if%>
              <%if errid="004" then%>
              您不是从本服务器上提交的表单信息！！！<br>
              <br>
              你的密码不正确,请返回从新输入！！<br>
              <br>
              操作提出进一步上诉的权利！！<br>
              <br>
              <br>
              <br>
              It is to the letter,please if there is any doubt: ：<a href="mailto:<%=Application("sunsports_email")%>"><%=Application("sunsports_email")%></a><br>
              <%end if%>
              <br>
            </p></td>
        </tr>
      </table></td>
  </tr>
</table>

