<!--#include file="Top.asp" -->
<H2><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000>Access Management</FONT></H2>
<P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
                  color=#000000 size=3>To create, update or delete an account, please select the type of account.</FONT></P>
<FORM name=form action=User_v.asp method=post>
  <P> <font color="#000000">
	<select name="adminlevel">
      <option value=1>Administrator</option>
      <option value=2>Managerial</option>
      <option value=3>Visitor</option>
    </select>
</font></P>
  <P> <font color="#000000">
    <INPUT type=submit value=view name=Submit>
    </font></P>
</FORM>
<font color="#000000">
<!----------------------------------------------------------------------------------------------------------------------
<!--                                                    EDIT / DELETE FIELDS
<!-------------------------------------------------------------------------------------------------------------------- -->
</font>
