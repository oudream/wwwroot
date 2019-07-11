<!--#include file="top.asp" -->
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>League Management </FONT></H2>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>To Edit / Delete a League or Cup, select 
  the competition from the list below.</FONT></P>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>Deleting a League / Cup would mean the 
  permanent removal of its corresponding records that are stored in the archives.</FONT></P>
<FORM name=form action=tournament_e.asp method=post>
  <P> <font color="#000000">
<select name="cid">
<%
sql="select * from tournament"
set rs=server.createobject("ADODB.Recordset")
rs.open sql,conn,1,1
while not rs.eof
%>
                                    <option value="<%=rs("cid")%>"><%=rs("name")%></option>
<%
rs.movenext
wend
rs.close
set rs=nothing
%>
                                  </select>
</font></P>
  <P> <font color="#000000">
    <INPUT type=submit value=Submit name=Submit>
    </font></P>
</FORM>
<font color="#000000">
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                    EDIT / DELETE FIELDS
<!-------------------------------------------------------------------------------------------------------------------- -->
</font>
