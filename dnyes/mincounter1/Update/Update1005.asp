<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/Common.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "ALTER TABLE t_Client ADD COLUMN"
				+ " TimeZone varchar(4) NULL";
			oConn.Execute( sSqlString )
			Response.write("t_Client ���³ɹ�!")
		}catch(e){
			Response.write("û������κβ������������Ѿ����������ˡ�")
			Response.write(e.description)
		}
		Response.write("<br>\n")
	oConn.Close()
%>
