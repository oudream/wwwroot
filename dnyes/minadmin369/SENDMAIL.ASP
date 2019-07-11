<!--#include file="MYJAMIL.ASP" -->
<%
sub getemailbody(emailbody)
emailbody="<br>"
emailbody=emailbody&"<br>"
emailbody=emailbody&"发件人：sales@dns99.cn 发送时间：2003-12-19 16:52:23 优先级：普通 <br>"
emailbody=emailbody&"收件人：oudream@163.com 抄　　送：(无) <br>"
emailbody=emailbody&"主　题：亿流网络订单提交通知 "
emailbody=emailbody&"<p>尊敬的用户：</p>"
emailbody=emailbody&"<p>您提交订单如下：</p>"
emailbody=emailbody&"<p>订 单 号： 5221<br>"
emailbody=emailbody&"  订购时间： 2003-12-19 16:52:23<br>"
emailbody=emailbody&"  产品编号： A003<br>"
emailbody=emailbody&"  产品名称： .cn国内域名<br>"
emailbody=emailbody&"  产品单价： 210.00<br>"
emailbody=emailbody&"  订购数量： 1<br>"
emailbody=emailbody&"  价格总计： 210 元</p>"
emailbody=emailbody&"<p>非常高兴您有意向购买我们的产品，如果您有任何问题请与我们联系，我们会及时替你解答<br>"
emailbody=emailbody&"  ，期待近一步的为您服务。</p>"
emailbody=emailbody&"<p>我们的付款方式：</p>"
emailbody=emailbody&"<p>1、邮局汇款（二天左右到帐） </p>"
emailbody=emailbody&"<p>地 址：江苏省南通市北郭东路26号 <br>"
emailbody=emailbody&"  邮 编：226001 <br>"
emailbody=emailbody&"  收款人：南通亿流网络有限公司 </p>"
emailbody=emailbody&"<p>汇款人附言：请注明用户名、汇款用途。</p>"
emailbody=emailbody&"<p>-----------------------------------------------------------------------</p>"
emailbody=emailbody&"<p>2、IPS网上支付（实时到帐，支持部分银行信用卡和借记卡） </p>"
emailbody=emailbody&"<p>详情请见http://www.dns99.cn/ipay/</p>"
emailbody=emailbody&"<p>-----------------------------------------------------------------------</p>"
emailbody=emailbody&"<p>3、中国银行电汇转帐（二天左右到帐，单位客户使用，个人用户请用其它方式） </p>"
emailbody=emailbody&"<p>开户银行：中国银行南通分行濠东路分理处 <br>"
emailbody=emailbody&"  公司全称：南通亿流网络有限公司 <br>"
emailbody=emailbody&"  银行帐号：844718831018091001 </p>"
emailbody=emailbody&"<p>-----------------------------------------------------------------------</p>"
emailbody=emailbody&"<p>4、个人用户快速汇款方式</p>"
emailbody=emailbody&"<p>个人用户推荐使用银行电子汇款，实时或当天到帐，便于我们更快处理您的订单</p>"
emailbody=emailbody&"<p>详情请见http://dns99.cn/i_pay.php</p>"
emailbody=emailbody&"<p>-----------------------------------------------------------------------</p>"
emailbody=emailbody&"<p>亿流网络（DNS99.CN）是具有因特网信息服务资质的专业ICP企业。拥有一支业内资深的技术<br>"
emailbody=emailbody&"  专家、优秀的运营和服务团队，公司已成功地创建并运营了拥有3万企业用户的系统平台，取<br>"
emailbody=emailbody&"  得有目共睹的成功。使企业可以在低成本、高效率、强保障的前提下建立自己的上网平台，从<br>"
emailbody=emailbody&"  而大大降低了企业信息化的门槛。</p> "
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  如果您有任何意见和建议，请和我们联系。</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  　　　　　　　　致<br>"
emailbody=emailbody&"  礼！<br>"
emailbody=emailbody&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  一流网络，一流服务--亿流网络</p>"
emailbody=emailbody&"<p>================================================</p>"
emailbody=emailbody&"<p>南通亿流网络有限公司<br>"
emailbody=emailbody&"  NanTong YiLiu NetWork Co.Ltd <br>"
emailbody=emailbody&" 地址：江苏省南通市北郭东路26号<br>"
emailbody=emailbody&"  邮编：226001</p>"
emailbody=emailbody&"<p>开户银行：中国银行南通分行濠东路分理处 <br>"
emailbody=emailbody&"  公司全称：南通亿流网络有限公司 <br>"
emailbody=emailbody&"  银行帐号：844718831018091001<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  电 话：0513-5280910，5284045 <br>"
emailbody=emailbody&"  传 真：0513-5292710<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  公司网址： http://www.dns99.cn<br>"
emailbody=emailbody&"  E-mail: sales@dns99.cn <br>"
emailbody=emailbody&"  OICQ业务咨询：666086 872200</p>"
emailbody=emailbody&"<p>= = = = = = = = = = = = = = = = = = = =<br>"
emailbody=emailbody&"</p>"
end sub
%>
<%
email=trim(request("email"))
emailsubject=trim(request("emailsubject"))
call getemailbody(emailmessage)
if email<>"" and emailmessage<>"" and emailsubject<>"" then
call Jmail(email,emailsubject,emailmessage)
end if
%>
<html>
<head>
<title>系统管理中心</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../inc/southdns.css" type="text/css">
</head>
<body bgcolor="#FFFFCC" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form action="sendmail.asp" method="post">
<table width="600" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="62">email to</td>
    <td width="612"><input maxlength=45 name="email" size=20 class="form"></td>
  </tr>
  <tr> 
    <td>subject</td>
    <td><input maxlength=45 name="emailsubject" size=50 class="form"></td>
  </tr>
  <tr> 
    <td>body</td>
    <td><textarea name="emailmessage" cols="80" class="form" rows="15"></textarea></td>
  </tr>
  <tr> 
    <td colspan="2"><input type="submit" name="Submit2" value="确 定">
      &nbsp;&nbsp; <input type="reset" name="Submit2" value="取 消"> &nbsp;&nbsp; 
      <input type="button" name="Submit2" value="后 退" onClick="history.go( -1 );return true;"></td>
  </tr>
</table>
  </form>
</body>
</html>
