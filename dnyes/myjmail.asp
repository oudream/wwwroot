<%	
sub Jmail2(email,topic,mailbody) ' support@dnyes.com postmaster@dnyes.com
dim JMail
Set JMail=Server.CreateObject("JMail.Message")
JMail.Logging=True
JMail.Charset="gb2312"
JMail.ContentType = "text/html"
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
JMail.MailServerUserName = "postmaster@dnyes.com"
JMail.MailServerPassword = "winer031203"
JMail.From = "support@dnyes.com" '发件人的信箱
JMail.FromName = "DNYES.COM"   '发件人的名字
JMail.Subject = topic  '邮件的主题
JMail.Body = mailbody   '邮件的内容
'==============================收件人的地址！
JMail.AddRecipient email         
'JMail.AddRecipientBcc "backup@dnyes.com" '收件人的地址
JMail.Priority=5   '邮件级别1-5数字越大级别越高---3为普通邮件
JMail.Send("218.85.132.56")   '红色变量是邮件服务器地址
JMail.Close 
Set JMail=nothing 
end sub 
%>


<%	
sub Jmail3(email,topic,mailbody) ' sales@dnyes.com
dim JMail
Set JMail=Server.CreateObject("JMail.Message")
JMail.Logging=True
JMail.Charset="gb2312"
JMail.ContentType = "text/html"
JMail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
JMail.MailServerUserName = "send@dnyes.com"
JMail.MailServerPassword = "winer031203"
JMail.From = "sales@dnyes.com" '发件人的信箱
JMail.FromName = "DNYES.COM"   '发件人的名字
JMail.Subject = topic  '邮件的主题
JMail.Body = mailbody   '邮件的内容
'==============================收件人的地址！
JMail.AddRecipient email         
JMail.AddRecipientBcc "backup@dnyes.com" '收件人的地址
JMail.Priority=5   '邮件级别1-5数字越大级别越高---3为普通邮件
JMail.Send("218.85.132.56")   '红色变量是邮件服务器地址
JMail.Close 
Set JMail=nothing 
end sub 
%>


<%
sub getemailbody1(emailbody,uid,pwd) '注册成功后的
emailbody="<br>"
emailbody=emailbody&"尊敬的客户:"
emailbody=emailbody&"<p> 您好！感谢您选择数信科技的服务。</p>"
emailbody=emailbody&"<p>请记住您的帐号： "&uid&"密码："&pwd&"</p>"
emailbody=emailbody&"<p>如果您有任何意见和建议，请和我们联系。</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  　　　　　　　　顺祝<br>"
emailbody=emailbody&"  商安<br>"
emailbody=emailbody&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  数信科技，以客为尊</p>"
emailbody=emailbody&"<p>================================================</p>"
emailbody=emailbody&"<p>数信科技有限公司<br>"
emailbody=emailbody&"  ShuXin KeJi Co.,Ltd. <br>"
emailbody=emailbody&"  地址：广东省顺德均安西堤路50号<br>"
emailbody=emailbody&"  邮编：528300</p>"
emailbody=emailbody&"<p>------------------------------------------------<br>"
emailbody=emailbody&"  电 话：0757-26823530，25571842<br>"
emailbody=emailbody&"  传 真：0757-25571842<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  公司网址： http://www.dnyes.com<br>"
emailbody=emailbody&"  E-mail: sales@dnyes.com <br>"
emailbody=emailbody&"  OICQ业务咨询：30013002 30013004<br>"
emailbody=emailbody&"  = = = = = = = = = = = = = = = = = = = =<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  信网客户分为4种级别,分别需要如下:</p>"
emailbody=emailbody&"<p>普通客户: 无需积分,注册后就是<br>"
emailbody=emailbody&"  铜牌客户: 预付款2000元 (或消费2000元)<br>"
emailbody=emailbody&"  银牌客户: 预付款10000元(或消费10000元)<br>"
emailbody=emailbody&"  金牌客户: 预付款20000元(或消费20000元)<br>"
emailbody=emailbody&"  不同会员级别享有不同的价格.</p>"
emailbody=emailbody&"<p>如何让我们更好地为你服务?<br>"
emailbody=emailbody&"  1.进入本网站 http://www.dnyes.com 时请先登录,本站的客户专区为您提供以下服务<br>"
emailbody=emailbody&"  --信息维护<br>"
emailbody=emailbody&"  1）更改密码：您可以更改您现在的会员的密码，改正后即可使用新的密码。<br>"
emailbody=emailbody&"  2）更改会员注册信息：对您填写的详细信息进行在线的修改，修改成功后，您以后注册和选购网站信息均不变更为您修改后的信息<br>"
emailbody=emailbody&"  --查询处理<br>"
emailbody=emailbody&"  1）未付费订单处理：查询您当前已经申请注册但还未付费的域名或网站的清单，您可以随时了解您的未付费订单的处理情况<br>"
emailbody=emailbody&"  2）查看已付费订单：查询您已经申请并且已经付费的域名或网站的订单情况<br>"
emailbody=emailbody&"  --相关知识<br>"
emailbody=emailbody&"  1）有关域名，邮箱，网站建设，虚拟主机等等的各种基本知识---<br>"
emailbody=emailbody&"  --齐全WHOIS<br>"
emailbody=emailbody&"  1) 有比较完全的域名查询(WHOIS)功能，查询后都有完整的相关说明<br>"
emailbody=emailbody&"  -----------------------------------------------<br>"
emailbody=emailbody&"  ...我们会不定期举办一些活动,请时刻关注网站通知</p>"
emailbody=emailbody&"<p>如何购买产品或服务?<br>"
emailbody=emailbody&"  您支付的所有款项,都将作为已消费款项入帐,达到一定金额时会自动转更高级别<br>"
emailbody=emailbody&"  不论是否有预付款,您都可以提交订单;<br>"
emailbody=emailbody&"  但只在存有预付款的情况下,您才可以用预付款确认购买;<br>"
emailbody=emailbody&"  进行付款确认的订单才会正式生效.</p>"
emailbody=emailbody&"<p>如果您以前是本站的用户,以前的业务可以积分吗?<br>"
emailbody=emailbody&"  凡是原信网任何代理商和用户,本站注册后,均可将原来的业务金额通知我们,我们查核后,将予以加入相应的积分消费.<br>"
emailbody=emailbody&"  -------------------------------------------<br>"
emailbody=emailbody&"  数信科技有限公司 http://www.dnyes.com</p>"
end sub
%>

<%
sub getemailbody2(emailbody,inBillNo,subid,subsname,price,amount,gross,nowtime,paymenttype,paymentmessage)'购买成功后的
emailbody="<br>"
emailbody=emailbody&"尊敬的客户:"
emailbody=emailbody&"<p>您提交订单如下：</p>"
emailbody=emailbody&"<p>订 单 号： "&inBillNo&"<br>"
emailbody=emailbody&"   产品编号： "&subid&"<br>"
emailbody=emailbody&"   产品名称： "&subsname&"<br>"
emailbody=emailbody&"   产品单价： "&price&"<br>"
emailbody=emailbody&"   订购数量： "&amount&"<br>"
emailbody=emailbody&"   价格总计： "&gross&"<br>"
emailbody=emailbody&"   订购时间： "&nowtime&"<br>"
if paymenttype="预付款支付" then 
emailbody=emailbody&"   所选的付款款方式： "&paymenttype&"<br><br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
emailbody=emailbody&"付款方式： "&paymenttype&"<br><br>"&paymentmessage&"<br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
else
emailbody=emailbody&"   所选的汇款方式： "&paymenttype&"<br><br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
emailbody=emailbody&"汇款方式： "&paymenttype&"<br><br>"&paymentmessage&"<br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
end if
emailbody=emailbody&"<a href=http://www.dnyes.com/banks.asp target=_blank >信网 数信科技——详细付款方式</a>"&"</p><br>"
emailbody=emailbody&"<p>非感谢您选择我们的产品，如果您有任何疑问请与我们联系，我们会及时替你解答</p><br>"
emailbody=emailbody&"<p>期待进一步的为您服务</p><br>"
emailbody=emailbody&"<p>如果您有任何意见和建议，请和我们联系。</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  　　　　　　　　致<br>"
emailbody=emailbody&"  礼！<br>"
emailbody=emailbody&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  数信科技，以客为尊</p>"
emailbody=emailbody&"<p>================================================<br><br>"
emailbody=emailbody&"  数信科技有限公司<br>"
emailbody=emailbody&"  ShuXin KeJi Co.,Ltd.<br>"
emailbody=emailbody&"  地址：广东省顺德均安西堤路50号<br>"
emailbody=emailbody&"  邮编：528300</p>"
emailbody=emailbody&"<p>------------------------------------------------<br>"
emailbody=emailbody&"  电 话：0757-26823530，25571842<br>"
emailbody=emailbody&"  传 真：0757-25571842<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  公司网址： http://www.dnyes.com<br>"
emailbody=emailbody&"  E-mail: sales@dnyes.com <br>"
emailbody=emailbody&"  OICQ业务咨询：30013002 30013004<br>"
emailbody=emailbody&"  = = = = = = = = = = = = = = = = = = = =</p>"
emailbody=emailbody&"<p>= = = = = = = = = = = = = = = = = = = =<br>"
emailbody=emailbody&"</p>"
end sub
%>

<%
sub getemailbody3(emailbody,uid,passlink)'取回密码
emailbody="<br><br><br>"
emailbody=emailbody&"<p>"&uid&"用户:您好!</p>"
emailbody=emailbody&"<p>您在信网 数信科技 - DNYES.COM的用户密码是:"&passlink&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  请保管好您的会员密码，谢谢！</p>"
emailbody=emailbody&"<p>-------------------------------------------------<br>"
emailbody=emailbody&"  信网 数信科技有限公司 http://www.dnyes.com </p>"
end sub
%>