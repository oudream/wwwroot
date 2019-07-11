<%server_v1=Cstr(Request.ServerVariables("SERVER_NAME")) & replace((Cstr(Request.ServerVariables("url"))),"testurL.asp","",1,-1,1)
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))

response.write server_v1
response.write "<br>"
response.write server_v2
response.end
%>