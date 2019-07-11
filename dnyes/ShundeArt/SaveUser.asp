<!--#include file="conn.asp"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="md5.asp"-->
<!--#include file="ChkURL.asp"-->
<%
UserName=trim(request("UserName"))
if Instr(request("UserName"),"=")>0 or Instr(request("UserName"),"%")>0 or Instr(request("username"),chr(32))>0 or Instr(request("username"),"?")>0 or Instr(request("username"),"&")>0 or Instr(request("username"),";")>0 or Instr(request("username"),",")>0 or Instr(request("username"),"'")>0 or Instr(request("username"),",")>0 or Instr(request("username"),chr(34))>0 or Instr(request("username"),chr(9))>0 or Instr(request("username"),"")>0 or Instr(request("username"),"$")>0 then
	Response.Write "<script>alert('对不起，您输入的用户名中包含有非法字符。');history.back();</Script>"
   	Response.End 
end if
Set rs=Server.CreateObject("Adodb.RecordSet")
rs.Open "Select * from "& db_User_Table &" where  "& db_User_Name &"='"&username&"'",ConnUser
if not rs.EOF then
	Response.Write "<script>alert('对不起，您需要的用户名已经被别人捷足先登了！另选一个吧。');history.back();</Script>"
   	Response.End 
end if

dim purview
dim oskey
dim sql
dim rs
dim fullname
dim passwd
dim question
dim answer
dim username
dim email
dim sex
dim birthyear,birthmonth,birthday
dim ip
dim content
dim tel
dim depid
dim depname
dim deptype
dim jingyong
dim reglevel
dim photo

fullname=htmlencode(request.form("fullname"))
content=htmlencode(request.form("content"))
passwd=md5(trim(request("passwd")))
question=htmlencode(request.form("question"))
answer=md5(trim(request("answer")))
username=htmlencode(request.form("username"))
email=htmlencode(request.form("email"))
purview=request.form("purview")
oskey=request.form("oskey")
sex=request.form("sex")
birthyear=request.form("birthyear")
birthmonth=request.form("birthmonth")
birthday=request.form("birthday")
tel=htmlencode(request.form("tel"))
depid=request.form("depid")
ip=Request.serverVariables("REMOTE_ADDR")
jingyong=request.form("jingyong")
reglevel=request.form("reglevel")
photo=request.form("photo")
set rs1=server.createobject("adodb.recordset")
sql1="select * from "& db_Dep_Table &" where id="&depid
rs1.open sql1,conn,1,3
depname=rs1("depname")
deptype=rs1("deptype")
rs1.close
set rs1=nothing

set rs=server.createobject("adodb.recordset")
sql="select * from "& db_User_Table &" where ("&db_User_ID&" is null)" 

rs.open sql,ConnUser,1,3
rs.addnew
rs(db_User_Name)=username
rs("content")=content
rs(db_User_Password)=passwd
rs(db_User_Question)=question
rs(db_User_Answer)=answer
rs("fullname")=fullname
rs(db_User_Email)=email
rs("purview")=purview
rs("oskey")=oskey
rs(db_User_Sex)=sex
''rs("birthyear")=birthyear
''rs("birthmonth")=birthmonth
rs(db_User_birthday)=birthday
rs("depid")=depid
rs("depname")=depname
rs("deptype")=deptype
rs("tel")=tel
rs(db_User_LastLoginIP)=ip
rs(db_User_LoginTimes)=1
rs(db_User_LastLoginTime)=now()
rs("shenhe")=1
rs("jingyong")=jingyong
rs("reglevel")=reglevel
rs(db_User_Face)=photo
rs(db_User_RegDate)=now()
rs.update
id=rs(db_User_Id)
rs.close


IF UserTableType = "Dvbbs" then
	'以下往动网设置参数表添加 db_Form_UserNum 和 db_Form_lastUser 值[总用户数、新注册用户名]
	dim usernum
	set rs2=server.CreateObject("ADODB.RecordSet")
	rs2.open "select * from "& db_Set_Table &"",conn,3,2
	usernum=rs2(db_Forum_UserNum)
	rs2(db_Forum_UserNum)=usernum+1
	rs2(db_Forum_lastUser)=Username
	rs2.update
	rs2.close
	set rs2=nothing
END IF
%>
<!--#include file=User_Top.asp-->
<head><title>申请成功</title>
<body topmargin="0">
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<div align="center">
    <table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="760" id="AutoNumber1">
    <tr>
      <td width="100%" height="20"> 
        <p align="center"><font color=red><b>会员注册成功</b></font><br><br><br><%if fabiao="1" then%>现在您可以从首页的会员登录处登录发表您的文章！<br><br><%else%>请等待管理员为您开通帐号！<br><br><%end if%>
      </td>
    </tr>
    <tr valign="middle" > 
      <td width="100%"> 
        <div align="center"><br>
          <br>
          <a href="javascript:window.close()">关闭窗口</a><br>
          </div>
      </td>
    </tr>
    </table>
</div>
</body>
<!--#include file=User_Bottom.asp-->