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
JMail.From = "support@dnyes.com" '�����˵�����
JMail.FromName = "DNYES.COM"   '�����˵�����
JMail.Subject = topic  '�ʼ�������
JMail.Body = mailbody   '�ʼ�������
'==============================�ռ��˵ĵ�ַ��
JMail.AddRecipient email         
'JMail.AddRecipientBcc "backup@dnyes.com" '�ռ��˵ĵ�ַ
JMail.Priority=5   '�ʼ�����1-5����Խ�󼶱�Խ��---3Ϊ��ͨ�ʼ�
JMail.Send("218.85.132.56")   '��ɫ�������ʼ���������ַ
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
JMail.From = "sales@dnyes.com" '�����˵�����
JMail.FromName = "DNYES.COM"   '�����˵�����
JMail.Subject = topic  '�ʼ�������
JMail.Body = mailbody   '�ʼ�������
'==============================�ռ��˵ĵ�ַ��
JMail.AddRecipient email         
JMail.AddRecipientBcc "backup@dnyes.com" '�ռ��˵ĵ�ַ
JMail.Priority=5   '�ʼ�����1-5����Խ�󼶱�Խ��---3Ϊ��ͨ�ʼ�
JMail.Send("218.85.132.56")   '��ɫ�������ʼ���������ַ
JMail.Close 
Set JMail=nothing 
end sub 
%>


<%
sub getemailbody1(emailbody,uid,pwd) 'ע��ɹ����
emailbody="<br>"
emailbody=emailbody&"�𾴵Ŀͻ�:"
emailbody=emailbody&"<p> ���ã���л��ѡ�����ſƼ��ķ���</p>"
emailbody=emailbody&"<p>���ס�����ʺţ� "&uid&"�����룺"&pwd&"</p>"
emailbody=emailbody&"<p>��������κ�����ͽ��飬���������ϵ��</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  ����������������˳ף<br>"
emailbody=emailbody&"  �̰�<br>"
emailbody=emailbody&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  ���ſƼ����Կ�Ϊ��</p>"
emailbody=emailbody&"<p>================================================</p>"
emailbody=emailbody&"<p>���ſƼ����޹�˾<br>"
emailbody=emailbody&"  ShuXin KeJi Co.,Ltd. <br>"
emailbody=emailbody&"  ��ַ���㶫ʡ˳�¾�������·50��<br>"
emailbody=emailbody&"  �ʱࣺ528300</p>"
emailbody=emailbody&"<p>------------------------------------------------<br>"
emailbody=emailbody&"  �� ����0757-26823530��25571842<br>"
emailbody=emailbody&"  �� �棺0757-25571842<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  ��˾��ַ�� http://www.dnyes.com<br>"
emailbody=emailbody&"  E-mail: sales@dnyes.com <br>"
emailbody=emailbody&"  OICQҵ����ѯ��30013002 30013004<br>"
emailbody=emailbody&"  = = = = = = = = = = = = = = = = = = = =<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  �����ͻ���Ϊ4�ּ���,�ֱ���Ҫ����:</p>"
emailbody=emailbody&"<p>��ͨ�ͻ�: �������,ע������<br>"
emailbody=emailbody&"  ͭ�ƿͻ�: Ԥ����2000Ԫ (������2000Ԫ)<br>"
emailbody=emailbody&"  ���ƿͻ�: Ԥ����10000Ԫ(������10000Ԫ)<br>"
emailbody=emailbody&"  ���ƿͻ�: Ԥ����20000Ԫ(������20000Ԫ)<br>"
emailbody=emailbody&"  ��ͬ��Ա�������в�ͬ�ļ۸�.</p>"
emailbody=emailbody&"<p>��������Ǹ��õ�Ϊ�����?<br>"
emailbody=emailbody&"  1.���뱾��վ http://www.dnyes.com ʱ���ȵ�¼,��վ�Ŀͻ�ר��Ϊ���ṩ���·���<br>"
emailbody=emailbody&"  --��Ϣά��<br>"
emailbody=emailbody&"  1���������룺�����Ը��������ڵĻ�Ա�����룬�����󼴿�ʹ���µ����롣<br>"
emailbody=emailbody&"  2�����Ļ�Աע����Ϣ��������д����ϸ��Ϣ�������ߵ��޸ģ��޸ĳɹ������Ժ�ע���ѡ����վ��Ϣ�������Ϊ���޸ĺ����Ϣ<br>"
emailbody=emailbody&"  --��ѯ����<br>"
emailbody=emailbody&"  1��δ���Ѷ���������ѯ����ǰ�Ѿ�����ע�ᵫ��δ���ѵ���������վ���嵥����������ʱ�˽�����δ���Ѷ����Ĵ������<br>"
emailbody=emailbody&"  2���鿴�Ѹ��Ѷ�������ѯ���Ѿ����벢���Ѿ����ѵ���������վ�Ķ������<br>"
emailbody=emailbody&"  --���֪ʶ<br>"
emailbody=emailbody&"  1���й����������䣬��վ���裬���������ȵȵĸ��ֻ���֪ʶ---<br>"
emailbody=emailbody&"  --��ȫWHOIS<br>"
emailbody=emailbody&"  1) �бȽ���ȫ��������ѯ(WHOIS)���ܣ���ѯ�������������˵��<br>"
emailbody=emailbody&"  -----------------------------------------------<br>"
emailbody=emailbody&"  ...���ǻ᲻���ھٰ�һЩ�,��ʱ�̹�ע��վ֪ͨ</p>"
emailbody=emailbody&"<p>��ι����Ʒ�����?<br>"
emailbody=emailbody&"  ��֧�������п���,������Ϊ�����ѿ�������,�ﵽһ�����ʱ���Զ�ת���߼���<br>"
emailbody=emailbody&"  �����Ƿ���Ԥ����,���������ύ����;<br>"
emailbody=emailbody&"  ��ֻ�ڴ���Ԥ����������,���ſ�����Ԥ����ȷ�Ϲ���;<br>"
emailbody=emailbody&"  ���и���ȷ�ϵĶ����Ż���ʽ��Ч.</p>"
emailbody=emailbody&"<p>�������ǰ�Ǳ�վ���û�,��ǰ��ҵ����Ի�����?<br>"
emailbody=emailbody&"  ����ԭ�����κδ����̺��û�,��վע���,���ɽ�ԭ����ҵ����֪ͨ����,���ǲ�˺�,�����Լ�����Ӧ�Ļ�������.<br>"
emailbody=emailbody&"  -------------------------------------------<br>"
emailbody=emailbody&"  ���ſƼ����޹�˾ http://www.dnyes.com</p>"
end sub
%>

<%
sub getemailbody2(emailbody,inBillNo,subid,subsname,price,amount,gross,nowtime,paymenttype,paymentmessage)'����ɹ����
emailbody="<br>"
emailbody=emailbody&"�𾴵Ŀͻ�:"
emailbody=emailbody&"<p>���ύ�������£�</p>"
emailbody=emailbody&"<p>�� �� �ţ� "&inBillNo&"<br>"
emailbody=emailbody&"   ��Ʒ��ţ� "&subid&"<br>"
emailbody=emailbody&"   ��Ʒ���ƣ� "&subsname&"<br>"
emailbody=emailbody&"   ��Ʒ���ۣ� "&price&"<br>"
emailbody=emailbody&"   ���������� "&amount&"<br>"
emailbody=emailbody&"   �۸��ܼƣ� "&gross&"<br>"
emailbody=emailbody&"   ����ʱ�䣺 "&nowtime&"<br>"
if paymenttype="Ԥ����֧��" then 
emailbody=emailbody&"   ��ѡ�ĸ���ʽ�� "&paymenttype&"<br><br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
emailbody=emailbody&"���ʽ�� "&paymenttype&"<br><br>"&paymentmessage&"<br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
else
emailbody=emailbody&"   ��ѡ�Ļ�ʽ�� "&paymenttype&"<br><br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
emailbody=emailbody&"��ʽ�� "&paymenttype&"<br><br>"&paymentmessage&"<br>"
emailbody=emailbody&"  ------------------------------------------------<br><br>"
end if
emailbody=emailbody&"<a href=http://www.dnyes.com/banks.asp target=_blank >���� ���ſƼ�������ϸ���ʽ</a>"&"</p><br>"
emailbody=emailbody&"<p>�Ǹ�л��ѡ�����ǵĲ�Ʒ����������κ���������������ϵ�����ǻἰʱ������</p><br>"
emailbody=emailbody&"<p>�ڴ���һ����Ϊ������</p><br>"
emailbody=emailbody&"<p>��������κ�����ͽ��飬���������ϵ��</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  ������������������<br>"
emailbody=emailbody&"  ��<br>"
emailbody=emailbody&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  ���ſƼ����Կ�Ϊ��</p>"
emailbody=emailbody&"<p>================================================<br><br>"
emailbody=emailbody&"  ���ſƼ����޹�˾<br>"
emailbody=emailbody&"  ShuXin KeJi Co.,Ltd.<br>"
emailbody=emailbody&"  ��ַ���㶫ʡ˳�¾�������·50��<br>"
emailbody=emailbody&"  �ʱࣺ528300</p>"
emailbody=emailbody&"<p>------------------------------------------------<br>"
emailbody=emailbody&"  �� ����0757-26823530��25571842<br>"
emailbody=emailbody&"  �� �棺0757-25571842<br>"
emailbody=emailbody&"  ------------------------------------------------<br>"
emailbody=emailbody&"  ��˾��ַ�� http://www.dnyes.com<br>"
emailbody=emailbody&"  E-mail: sales@dnyes.com <br>"
emailbody=emailbody&"  OICQҵ����ѯ��30013002 30013004<br>"
emailbody=emailbody&"  = = = = = = = = = = = = = = = = = = = =</p>"
emailbody=emailbody&"<p>= = = = = = = = = = = = = = = = = = = =<br>"
emailbody=emailbody&"</p>"
end sub
%>

<%
sub getemailbody3(emailbody,uid,passlink)'ȡ������
emailbody="<br><br><br>"
emailbody=emailbody&"<p>"&uid&"�û�:����!</p>"
emailbody=emailbody&"<p>�������� ���ſƼ� - DNYES.COM���û�������:"&passlink&"</p>"
emailbody=emailbody&"<p><br>"
emailbody=emailbody&"  �뱣�ܺ����Ļ�Ա���룬лл��</p>"
emailbody=emailbody&"<p>-------------------------------------------------<br>"
emailbody=emailbody&"  ���� ���ſƼ����޹�˾ http://www.dnyes.com </p>"
end sub
%>