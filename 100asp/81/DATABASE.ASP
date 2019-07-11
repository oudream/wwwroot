<%
  dim conn
  dim connstr
  call conn_init()

  sub conn_init()
       connstr = "DBQ=" + server.mappath("message.mdb") + ";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"       
       set conn=server.createobject("ADODB.CONNECTION")
       if err.number<>0 then 
           err.clear
           set conn=nothing
		   response.write "数据库连接出错！"
		   Response.End
       else
		   conn.open connstr
           if err then 
              err.clear
              set conn=nothing
			   response.write "数据库连接出错！"
              Response.End 
           end if
       end if   
  end sub
	
  sub endConnection()
      set conn=nothing
  end sub  
%>