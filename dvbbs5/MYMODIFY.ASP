<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/chkinput.asp"-->
<!--#include file="inc/usercp.asp"-->
<%
'=========================================================
' File: mymodify.asp
' Version:5.0
' Date: 2002-9-11
' Script Written by satan
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
response.buffer=true
stats="修改资料"
dim sex
dim showre
dim face,width,height
dim birthday
dim usercookies
call nav()

if Cint(GroupSetting(16))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本论坛修改资料的权限，请<a href=login.asp>登陆</a>或者同管理员联系。"
	founderr=true
end if

if not founduser then
	Errmsg=Errmsg+"<br>"+"<li>您还没有登陆，请登陆后进行修改。"
	founderr=true
end if

if founderr then
  	errmsg=errmsg+"<br>"+"<li>您没有<a href=login.asp target=_blank>登录</a>。"
	call head_var(1)
	call dvbbs_error()
else
	call cp_var()
	if request("action")="updat" then
		call update()
		if founderr then
			call dvbbs_error()
		else
			sucmsg="<li>更改资料成功！"
			call dvbbs_suc()
		end if
	else
		call userinfo()
		if founderr then call dvbbs_error()
	end if
end if

call activeonline()
call footer()

sub userinfo()
set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
	exit sub
else
%>

<form action="mymodify.asp?action=updat&username=<%=htmlencode(membername)%>" method=POST name="theForm">
<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
<tr> 
      <th colspan="2" width="100%">基本设置选项
      </th>
    </tr> 
<%if cint(Forum_Setting(6))=1 then%>
      <tr> 
        <td width="40%" class=tablebody1><B>头&nbsp;&nbsp;&nbsp;&nbsp;衔</B>：</td>   
        <td width="60%" class=tablebody1> <input type="text" name="title" value="<%if trim(rs("title"))<>"" then%><%=htmlencode(RS("title"))%><%end if%>" size=15 maxlength=13> </td>   
      </tr>   
<%end if%>
      <tr>    
        <td width="40%" class=tablebody1><B>性&nbsp;&nbsp;&nbsp;&nbsp;别</B>：<BR>请选择您的性别</td>   
        <td width="60%" class=tablebody1><input type="radio" value="1"<%if rs("sex")=1 then%> checked <%end if%>name="Sex">酷哥
      <input type="radio" name="Sex"<%if rs("sex")=0 then%> checked <%end if%> value="0">靓妹</td>
</td>   
      </tr>
      <tr>    
        <td width="40%" class=tablebody1><B>生&nbsp;&nbsp;&nbsp;&nbsp;日</B><BR>如不想填写，请全部留空</td>   
        <td width="60%" class=tablebody1>    
				  <select name=birthyear>
				  <option value=""></option>
				  <%for i=1901 to 2000%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(year(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    年 
				  <select name=birthmonth>
				  <option value=""></option>
				  <%for i=1 to 12%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(month(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    月 
				  <select name=birthday>
				  <option value=""></option>
				  <%for i=1 to 31%>
					<option value="<%=i%>" <%if not isnull(rs("birthday")) and isdate(rs("birthday")) then%><%if cint(day(rs("birthday")))=cint(i) then response.write "selected"%><%end if%>><%=i%></option>
				  <%next%>
				  </select>
                    日</font>
        </td>   
      </tr>
      <tr>    
        <td width="40%" valign=top class=tablebody1><B>头&nbsp;&nbsp;&nbsp;&nbsp;像</B>：<BR>
您可以选择一个个性图片，当你发表时将显示在您的名字下方。<BR>
如果你填写了下面的自定义头像部分，那么你的头像以自定义的为准。否则，请你留空自定义头像的所有栏目！<BR>
<B>关于自定义头像</B>：<BR>
你也可以在这里给出你自定义头像的 URL 地址，头像的高度和宽度(像素)。 如果不想要自定义头像，请将相应栏目栏目全部留空！
		</td>   
        <td width="60%" class=tablebody1>
 <table border="0" cellpadding="0" cellspacing="1" width="98%" >
            <tr>
              <td width="*"  class=tablebody1>论坛头像选择 
<select name=face style="font-family:arial" size=1 onChange="document.images['face'].src=options[selectedIndex].value; parent.document.forms[0].myface.value=options[selectedIndex].value"> 
<%
	dim userregface
	for i=1 to 60
	userregface=Forum_info(7) & "Image" & i & ".gif"
	if trim(lcase(rs("face"))) = trim(lcase(userregface)) then
%> 
   <option value='<%=Forum_info(7)%>Image<%=i%>.gif' selected>Image<%=i%></option> 
     <% else %> 
   <option value='<%=Forum_info(7)%>Image<%=i%>.gif'>Image<%=i%></option> 
<%
	end if 
	next
%> 
                </select><a href="allface.asp" target="_blank" title="头像预览列表">[<font color=red>头像预览</font>]</a></td>
              <td width="180"  rowspan="2"  valign=top align=center>
<img id=face src="<%=rs("face")%>" width=<%=rs("width")%> height=<%=rs("height")%>>
</td>
            </tr>
            <tr>
              <td width="*" >
<% if Forum_Setting(7) then %> 
          <iframe name="ad" frameborder=0 width=300 height=40 scrolling=no src=reg_upload.asp></iframe> <br>
		<% end if %>
		
		  自定义头像地址：<input type="TEXT" name="myface" size=30 maxlength=100 value=<%=rs("face")%>> <br>
	  宽&nbsp;&nbsp;&nbsp;&nbsp;度：<input type="TEXT" name="width" size=3 maxlength=3 value=<%=rs("width")%>>20---120的整数<br>
	  高&nbsp;&nbsp;&nbsp;&nbsp;度：<input type="TEXT" name="height" size=3 maxlength=3 value=<%=rs("height")%>>20---120的整数<br>
</td>
            </tr>
          </table>

        </td>   
      </tr>	  
      <tr>    
        <td width="40%" class=tablebody1><B>个人照片</B>：<BR>如果您有照片在网上，请输入网页地址。此项可选</td>   
        <td width="60%" class=tablebody1>    
          <input type="TEXT" name="userphoto" value="<%if trim(rs("userphoto"))<>"" then%><%=htmlencode(rs("userphoto"))%><%end if%>" size=30 maxlength=100>   
        </td>   
      </tr>   
<%if cint(Forum_Setting(32))=1 then%>
      <tr>    
        <td width="40%" class=tablebody1><B>门&nbsp;&nbsp;&nbsp;&nbsp;派</B>：<BR>如果您有兴趣，可加入以下门派</td>   
        <td width="60%" class=tablebody1>    
<select name=groupname>
<%
dim grs
set grs=conn.execute("select * from GroupName")
if grs.eof and grs.bof then
%>
<option value="无门无派">无门无派</option>
<%else%>
<%do while not grs.eof%>
<option value="<%=grs("Groupname")%>"><%=grs("GroupName")%></option>
<%
	grs.movenext
	loop
%>
<%end if%>
</select>
<%set grs=nothing%>
        </td>   
      </tr>
<%end if%>
      <tr>    
        <td valign=top width="40%" class=tablebody1><B>签&nbsp;&nbsp;&nbsp;&nbsp;名</B>：<BR>不能超过 300 个字符   
          <br>   
          文字将出现在您发表的文章的结尾处。</td>   
        <td width="60%" class=tablebody1>    
          <textarea name="Signature" rows=5 cols=60 wrap=PHYSICAL><%if rs("sign")<>"" then%>
<%
dim signtrue
signtrue=replace(rs("sign"),"<BR>",chr(13))
signtrue=replace(signtrue,"&nbsp;"," ")
%><%=signtrue%><%end if%></textarea>   
        </td>   
      </tr>
<tr> 
      <th colspan="2" width="100%">其他选项
      </th>
    </tr> 
 <tr>    
        <td width="40%" class=tablebody1><B>有回帖时是否提示您</B>：<BR>当您的帖子有人回复后有信息通知您</td>   
        <td width="60%" class=tablebody1>    
			<input type="radio" name="showRe" value="1" <% if rs("showRe") then response.write "checked" %> >
			提示我
			<input type="radio" name="showRe" value="0" <% if not rs("showRe") then response.write "checked" %> >
			不提示
        </td>   
      </tr>	  
      <tr>    
        <td width="40%" class=tablebody1><B>选择Cookie的保留时间</B>：</td>   
        <td width="60%" class=tablebody1>    
			
    <input type="radio" name="usercookies" value="1" <% if request.cookies("aspsky")("usercookies")=1 then response.write "checked" %> >
    <font color=red>1天</font>
    <input type="radio" name="usercookies" value="2" <% if request.cookies("aspsky")("usercookies")=2 then response.write "checked" %> >
    1个月
    <input type="radio" name="usercookies" value="3" <% if request.cookies("aspsky")("usercookies")=3 then response.write "checked" %>>
    1年
    <input type="radio" name="usercookies" value="0" <% if request.cookies("aspsky")("usercookies")=0 then response.write "checked" %>>
    不保留 </td>   
      </tr>
<%
dim usersetting,setuserinfo,setusertrue
if rs("usersetting")<>"" then
	usersetting=split(rs("usersetting"),"|||")
	if ubound(usersetting)=1 then
	if isnumeric(usersetting(0)) then setuserinfo=usersetting(0) else setuserinfo=1
	if isnumeric(usersetting(1)) then setusertrue=usersetting(1) else setusertrue=0
	else
	setuserinfo=1
	setusertrue=0
	end if
else
	setuserinfo=1
	setusertrue=0
end if
%>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的基本资料</B>：<BR>开放后别人可以看到您的性别、Email、QQ等信息</td>  <td width=60% class=tablebody1>    
			
              <input type=radio name=setuserinfo value=1  <%if setuserinfo=1 then%>checked<%end if%>>
              开放              <input type=radio name=setuserinfo value=0  <%if setuserinfo=0 then%>checked<%end if%>>
    不开放 </td>   
      </tr>
<tr>    
<td width=40% class=tablebody1><B>是否开放您的真实资料</B>：<BR>开放后别人可以看到您的真实姓名、联系方式等信息</font></td>  <td width=60% class=tablebody1>    
			
              <input type=radio name=setusertrue value=1 <%if setusertrue=1 then%>checked<%end if%>>
              开放              <input type=radio name=setusertrue value=0 <%if setusertrue=0 then%>checked<%end if%>>
    不开放 </td>   
      </tr>
<%
dim userinfo
dim realname,character,personal,country,province,city,shengxiao,blood,belief,occupation,marital, education,college,userphone,address
if rs("userinfo")<>"" then
	userinfo=split(rs("userinfo"),"|||")
	if ubound(userinfo)=14 then
		realname=userinfo(0)
		character=userinfo(1)
		personal=userinfo(2)
		country=userinfo(3)
		province=userinfo(4)
		city=userinfo(5)
		shengxiao=userinfo(6)
		blood=userinfo(7)
		belief=userinfo(8)
		occupation=userinfo(9)
		marital=userinfo(10)
		education=userinfo(11)
		college=userinfo(12)
		userphone=userinfo(13)
		address=userinfo(14)
	else
		realname=""
		character=""
		personal=""
		country=""
		province=""
		city=""
		shengxiao=""
		blood=""
		belief=""
		occupation=""
		marital=""
		education=""
		college=""
		userphone=""
		address=""
	end if
else
	realname=""
	character=""
	personal=""
	country=""
	province=""
	city=""
	shengxiao=""
	blood=""
	belief=""
	occupation=""
	marital=""
	education=""
	college=""
	userphone=""
	address=""
end if
%>
<tr>
<th height=25 align=left valign=middle colspan=2>&nbsp;个人真实信息（以下内容建议填写）</th>
</tr>
<tr>
<td valign=top width=40% class=tablebody1> 　<b>真实姓名：</b>
<input type=text name=realname size=18 value="<%=realname%>">
</td>
<td height=71 align=left valign=top  class=tablebody1 rowspan=14 width=60% >
<table width=100% border=0 cellspacing=0 cellpadding=5>
<tr>
<td class=tablebody1><b>性　格： &nbsp; </b>
<textarea name=character rows=6 cols=65><%=character%></textarea></td>
</tr>
<tr>
<td class=tablebody1><b>个人简介： &nbsp;
<textarea name=personal rows=6 cols=65><%=personal%></textarea>
</b></td>
</tr>
</table>
</td>
</tr>
<tr>
<td valign=top width=40%  class=tablebody1>　<b>国　　家：</b>
<b>
<input type=text name=country size=18 value="<%=country%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40%  class=tablebody1>　<b>联系电话：</b>
<b>
<input type=text name=userphone size=18 value="<%=userphone%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>通信地址：</b>
<b>
<input type=text name=address size=18 value="<%=address%>">
</b> </td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>省　　份：</b>
<input type=text name=province size=18 value="<%=province%>">
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>城　　市：
</b>
<input type=text name=city size=18 value="<%=city%>">
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>生　　肖：
</b>
<select size=1 name=shengxiao>
<option <%if shengxiao="" then%>selected<%end if%>></option>
<option value=鼠 <%if shengxiao="鼠" then%>selected<%end if%>>鼠</option>
<option value=牛 <%if shengxiao="牛" then%>selected<%end if%>>牛</option>
<option value=虎 <%if shengxiao="虎" then%>selected<%end if%>>虎</option>
<option value=兔 <%if shengxiao="兔" then%>selected<%end if%>>兔</option>
<option value=龙 <%if shengxiao="龙" then%>selected<%end if%>>龙</option>
<option value=蛇 <%if shengxiao="蛇" then%>selected<%end if%>>蛇</option>
<option value=马 <%if shengxiao="马" then%>selected<%end if%>>马</option>
<option value=羊 <%if shengxiao="羊" then%>selected<%end if%>>羊</option>
<option value=猴 <%if shengxiao="猴" then%>selected<%end if%>>猴</option>
<option value=鸡 <%if shengxiao="鸡" then%>selected<%end if%>>鸡</option>
<option value=狗 <%if shengxiao="狗" then%>selected<%end if%>>狗</option>
<option value=猪 <%if shengxiao="猪" then%>selected<%end if%>>猪</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>血　　型：</b>
<select size=1 name=blood>
<option <%if blood="" then%>selected<%end if%>></option>
<option value=A <%if blood="A" then%>selected<%end if%>>A</option>
<option value=B <%if blood="B" then%>selected<%end if%>>B</option>
<option value=AB <%if blood="AB" then%>selected<%end if%>>AB</option>
<option value=O <%if blood="O" then%>selected<%end if%>>O</option>
<option value=其他 <%if blood="其他" then%>selected<%end if%>>其他</option>
</select>
</td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>信　　仰：</b>
<select size=1 name=belief>
<option <%if belief="" then%>selected<%end if%>></option>
<option value=佛教 <%if belief="佛教" then%>selected<%end if%>>佛教</option>
<option value=道教 <%if belief="道教" then%>selected<%end if%>>道教</option>
<option value=基督教 <%if belief="基督教" then%>selected<%end if%>>基督教</option>
<option value=天主教 <%if belief="天主教" then%>selected<%end if%>>天主教</option>
<option value=回教 <%if belief="回教" then%>selected<%end if%>>回教</option>
<option value=无神论者 <%if belief="无神论者" then%>selected<%end if%>>无神论者</option>
<option value=共产主义者 <%if belief="共产主义者" then%>selected<%end if%>>共产主义者</option>
<option value=其他 <%if belief="其他" then%>selected<%end if%>>其他</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>职　　业： </b>
<select name=occupation>
<option <%if occupation="" then%>selected<%end if%>> </option>
<option value="财会/金融" <%if occupation="财会/金融" then%>selected<%end if%>>财会/金融</option>
<option value=工程师 <%if occupation="工程师" then%>selected<%end if%>>工程师</option>
<option value=顾问 <%if occupation="顾问" then%>selected<%end if%>>顾问</option>
<option value=计算机相关行业 <%if occupation="计算机相关行业" then%>selected<%end if%>>计算机相关行业</option>
<option value=家庭主妇 <%if occupation="家庭主妇" then%>selected<%end if%>>家庭主妇</option>
<option value="教育/培训" <%if occupation="教育/培训" then%>selected<%end if%>>教育/培训</option>
<option value="客户服务/支持" <%if occupation="客户服务/支持" then%>selected<%end if%>>客户服务/支持</option>
<option value="零售商/手工工人" <%if occupation="零售商/手工工人" then%>selected<%end if%>>零售商/手工工人</option>
<option value=退休 <%if occupation="退休" then%>selected<%end if%>>退休</option>
<option value=无职业 <%if occupation="无职业" then%>selected<%end if%>>无职业</option>
<option value="销售/市场/广告" <%if occupation="销售/市场/广告" then%>selected<%end if%>>销售/市场/广告</option>
<option value=学生 <%if occupation="学生" then%>selected<%end if%>>学生</option>
<option value=研究和开发 <%if occupation="研究和开发" then%>selected<%end if%>>研究和开发</option>
<option value="一般管理/监督" <%if occupation="一般管理/监督" then%>selected<%end if%>>一般管理/监督</option>
<option value="政府/军队" <%if occupation="政府/军队" then%>selected<%end if%>>政府/军队</option>
<option value="执行官/高级管理" <%if occupation="执行官/高级管理" then%>selected<%end if%>>执行官/高级管理</option>
<option value="制造/生产/操作" <%if occupation="制造/生产/操作" then%>selected<%end if%>>制造/生产/操作</option>
<option value=专业人员 <%if occupation="专业人员" then%>selected<%end if%>>专业人员</option>
<option value="自雇/业主" <%if occupation="自雇/业主" then%>selected<%end if%>>自雇/业主</option>
<option value=其他 <%if occupation="其他" then%>selected<%end if%>>其他</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>婚姻状况：</b>
<select size=1 name=marital>
<option <%if marital="" then%>selected<%end if%>></option>
<option value=未婚 <%if marital="未婚" then%>selected<%end if%>>未婚</option>
<option value=已婚 <%if marital="已婚" then%>selected<%end if%>>已婚</option>
<option value=离异 <%if marital="离异" then%>selected<%end if%>>离异</option>
<option value=丧偶 <%if marital="丧偶" then%>selected<%end if%>>丧偶</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>最高学历：</b>
<select size=1 name=education>
<option <%if education="" then%>selected<%end if%>></option>
<option value=小学 <%if education="小学" then%>selected<%end if%>>小学</option>
<option value=初中 <%if education="初中" then%>selected<%end if%>>初中</option>
<option value=高中 <%if education="高中" then%>selected<%end if%>>高中</option>
<option value=大学 <%if education="大学" then%>selected<%end if%>>大学</option>
<option value=硕士 <%if education="硕士" then%>selected<%end if%>>硕士</option>
<option value=博士 <%if education="博士" then%>selected<%end if%>>博士</option>
</select></td>
</tr>
<tr>
<td valign=top width=40% class=tablebody1>　<b>毕业院校：</b>
<input type=text name=college size=18 value="<%=college%>"></td>
</tr>

    <tr align="center"> 
      <td colspan="2" width="100%" class=tablebody2>
            <input type=Submit value="更 新" name="Submit"> &nbsp; <input type="reset" name="Submit2" value="清 除">
      </td>
    </tr>  
    </table></form>
   
</body>
</html> 
<%
end if
rs.close
set rs=nothing
end sub

sub update()
if request("sex")="" then
	errmsg=errmsg+"<br>"+"<li>请选择您的性别。"
	founderr=true
elseif request("sex")=0 or request("sex")=1 then
	sex=request("sex")
else
	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
	founderr=true
end if
		
if request("showRe")="" then
	errmsg=errmsg+"<br>"+"<li>请选择有回帖时是否提示您。"
	founderr=true
elseif request("showRe")=0 or request("showRe")=1 then
	showRe=request("showRe")
else
	errmsg=errmsg+"<br>"+"<li>您输入的字符非法。"
	founderr=true
end if
				
if request("myface")<>"" then
	if request("width")="" or request("height")="" then
		errmsg=errmsg+"<br>"+"<li>请输入图片的宽度和高度。"
		founderr=true
	elseif not isInteger(request("width")) or not isInteger(request("height")) then
		errmsg=errmsg+"<br>"+"<li>您输入的字符不合法。"
		founderr=true
	elseif request("width")<20 or request("width")>120 then
		errmsg=errmsg+"<br>"+"<li>您输入的图片宽度不符合标准。"
		founderr=true
	elseif request("height")<20 or request("height")>120 then
		errmsg=errmsg+"<br>"+"<li>您输入的图片高度不符合标准。"
		founderr=true
	else
		face=request("myface")
		width=request("width")
		height=request("height")
	end if
else
	if request("face")="" then
		errmsg=errmsg+"<br>"+"<li>请选择您的个性头像。"
		founderr=true
	elseif Instr(request("face"),Forum_info(7))>0 then
		face=request("face")
		width=32
		height=32
	else
		errmsg=errmsg+"<br>"+"<li>您选择了错误的头像。"
		founderr=true
	end if
end if
	
if request.Form("birthyear")="" or request.form("birthmonth")="" or request.form("birthday")="" then
	birthday=""
else
	birthday=trim(Request.Form("birthyear"))&"-"&trim(Request.Form("birthmonth"))&"-"&trim(Request.Form("birthday"))
	if not isdate(birthday) then birthday=""
end if

dim userinfo,usersetting
userinfo=checkreal(request.Form("realname")) & "|||" & checkreal(request.Form("character")) & "|||" & checkreal(request.Form("personal")) & "|||" & checkreal(request.Form("country")) & "|||" & checkreal(request.Form("province")) & "|||" & checkreal(request.Form("city")) & "|||" & request.Form("shengxiao") & "|||" & request.Form("blood") & "|||" & request.Form("belief") & "|||" & request.Form("occupation") & "|||" & request.Form("marital") & "|||" & request.Form("education") & "|||" & checkreal(request.Form("college")) & "|||" & checkreal(request.Form("userphone")) & "|||" & checkreal(request.Form("address"))
usersetting=request.Form("setuserinfo") & "|||" & request.Form("setusertrue")

set rs=server.createobject("adodb.recordset")
sql="Select * from [User] where userid="&userid
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>该用户名不存在。"
	founderr=true
else
	rs("face")=face
	rs("width")=width
	rs("height")=height
	rs("sex")=sex
	rs("sign")=trim(request("Signature"))
	rs("userphoto")=request("userphoto")
	rs("showRe")=showRe
	rs("UserGroup")=request("GroupName")
	if request("title")<>"" then
		rs("title")=htmlencode(request("title"))
	end if
	if birthday<>"" then
		rs("birthday")=birthday
	end if
	rs("userinfo")=userinfo
	rs("usersetting")=usersetting
	rs("bbstype")=request("bbstype")
	rs.Update
	usercookies=request("usercookies")
	select case usercookies
	case 0
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 1
		Response.Cookies("aspsky").Expires=Date+1
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 2
		Response.Cookies("aspsky").Expires=Date+31
		Response.Cookies("aspsky")("usercookies") = usercookies
	case 3
		Response.Cookies("aspsky").Expires=Date+365
		Response.Cookies("aspsky")("usercookies") = usercookies
	end select			
		
	Response.Cookies("aspsky")("tempid") = request("bbstype")
	Response.Cookies("aspsky").path=cookiepath
end if
rs.close
set rs=nothing
end sub

function checkreal(v)
dim w
if not isnull(v) then
	w=replace(v,"|||","§§§")
	checkreal=w
end if
end function
%>