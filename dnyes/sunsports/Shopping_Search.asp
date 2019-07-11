<SCRIPT LANGUAGE="JavaScript">
<!--
function checkform_search()
{
	if (form_search.hw_name.value.length == 0) {
		alert('Please enter a valid value for search .');
		form_search.hw_name.focus();
		return false;
		}
	return true;
}

//-->
</SCRIPT>
<table width="100%" border="0" cellpadding="0" cellspacing="7">
  <tr>
    <td>
          <form method="POST" action="Shopping_Search_Result.asp" id="form_search" name="form_search" onSubmit="return checkform_search();" target="_blank">
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
            <td><b><font color="#1F4A71">Product search :</font></b></td>
        </tr>
        <tr> 
          <td><input type="text" name="hw_name" size="15" class=input style="FONT-SIZE: 12px; WIDTH: 150px; background:white;border:1px solid #6184A3" onMouseOver="this.style.border='1px solid red'" onMouseOut="this.style.border='1px solid #6184A3'"></td>
        </tr>
        <tr> 
            <td height="10"><img src="1.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td><input type="submit" name="Submit" value="Search" class="button"></td>
        </tr>
      </table>
          </form>
	  </td>
  </tr>
</table>
