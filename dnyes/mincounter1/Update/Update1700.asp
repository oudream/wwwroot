<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "CREATE VIEW QueryDuplicateData AS "
				+"SELECT * FROM t_count WHERE id not in( select min(id) from t_Count group by siteid,countdate );";
			oConn.Execute( sSqlString )
			Response.write("QueryDuplicateData �����ɹ�!")
		}catch(e){
			Response.write("û������κβ������������Ѿ�ִ�й����������");
			Response.write(e.description);
		}
		Response.write("<br>\n")

	oConn.Close()
%>


1700 ������ɡ�
