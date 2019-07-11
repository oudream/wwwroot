<%
errid=request("err")
if errid="" then err="001"
%>
<table width="100%" align="center">
  <tr align="center"> 
    <td align="center"> <br>
      -- <strong>you operation error</strong> --<br> <br> <br> </td>
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
            <p> It is to the letter,please if there is any doubt: £º<a href="mailto:suport@dnyes.com%>">suport@dnyes.com</a></p>
            <p><br>
              <%end if%>
              <br>
            </p></td>
        </tr>
      </table></td>
  </tr>
</table>

