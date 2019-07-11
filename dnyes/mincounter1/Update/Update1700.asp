<%@LANGUAGE="JAVASCRIPT" CODEPAGE="936"%>
<!--#include file="../_inc/chkLogin.asp"-->
<%
	oConn.Open()
		try{
			sSqlString = "CREATE VIEW QueryDuplicateData AS "
				+"SELECT * FROM t_count WHERE id not in( select min(id) from t_Count group by siteid,countdate );";
			oConn.Execute( sSqlString )
			Response.write("QueryDuplicateData 创建成功!")
		}catch(e){
			Response.write("没有完成任何操作，可能你已经执行过此项操作。");
			Response.write(e.description);
		}
		Response.write("<br>\n")

	oConn.Close()
%>


1700 更新完成。
