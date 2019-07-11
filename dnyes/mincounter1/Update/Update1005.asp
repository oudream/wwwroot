<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/Common.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "ALTER TABLE t_Client ADD COLUMN"
				+ " TimeZone varchar(4) NULL";
			oConn.Execute( sSqlString )
			Response.write("t_Client 更新成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经从升过级了。")
			Response.write(e.description)
		}
		Response.write("<br>\n")
	oConn.Close()
%>
