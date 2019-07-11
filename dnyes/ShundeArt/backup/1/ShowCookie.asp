<%
  For each cookie in Request.Cookies
      if Request.cookies(cookie).HasKeys =false then
          Response.write cookie & "=" & Request.Cookies(cookie)
          Response.write ("<br>") 
      Else
          for each key in Request.Cookies(cookie)
              Response.write ("<font color=blue>")
              Response.write cookie & ".("&key&")" & "=" & Request.Cookies(cookie)(key)
              Response.write ("</font><br>")
          next
      end if
  next
%> 