<%
if request("module")<> "" then
	if request("module")="personal" then
	response.Redirect("personalreg.asp")
	else
	response.Redirect("businessreg.asp")
	end if 
else
	response.Redirect("personalreg.asp")
end if
%>