<!--#include file="Conn.ASP"-->
<!--#include file="ConnUser.asp"-->
<!--#include file="config.asp"-->
<!--#include file="char.inc"-->
<!--#include file="chkuser.asp" -->
<!--#include file="ChkURL.asp"-->
<HTML><HEAD><TITLE>������</TITLE>
<META http-equiv=content-type content="text/html; charset=gb2312">
<META content=0 name=expires>
<META content=no-cache name=Pragma>
<META content=no-cache name=Cache-Control>
<SCRIPT language=JScript>
<!--
/*****************************************************************************
                                   ����ƫ���趨
*****************************************************************************/

var conWeekend = 3;  // ��ĩ��ɫ��ʾ: 1=��ɫ, 2=��ɫ, 3=��ɫ, 4=������


/*****************************************************************************
                                   ��������
*****************************************************************************/

var lunarInfo=new Array(
0x4bd8,0x4ae0,0xa570,0x54d5,0xd260,0xd950,0x5554,0x56af,0x9ad0,0x55d2,
0x4ae0,0xa5b6,0xa4d0,0xd250,0xd255,0xb54f,0xd6a0,0xada2,0x95b0,0x4977,
0x497f,0xa4b0,0xb4b5,0x6a50,0x6d40,0xab54,0x2b6f,0x9570,0x52f2,0x4970,
0x6566,0xd4a0,0xea50,0x6a95,0x5adf,0x2b60,0x86e3,0x92ef,0xc8d7,0xc95f,
0xd4a0,0xd8a6,0xb55f,0x56a0,0xa5b4,0x25df,0x92d0,0xd2b2,0xa950,0xb557,
0x6ca0,0xb550,0x5355,0x4daf,0xa5b0,0x4573,0x52bf,0xa9a8,0xe950,0x6aa0,
0xaea6,0xab50,0x4b60,0xaae4,0xa570,0x5260,0xf263,0xd950,0x5b57,0x56a0,
0x96d0,0x4dd5,0x4ad0,0xa4d0,0xd4d4,0xd250,0xd558,0xb540,0xb6a0,0x95a6,
0x95bf,0x49b0,0xa974,0xa4b0,0xb27a,0x6a50,0x6d40,0xaf46,0xab60,0x9570,
0x4af5,0x4970,0x64b0,0x74a3,0xea50,0x6b58,0x5ac0,0xab60,0x96d5,0x92e0,
0xc960,0xd954,0xd4a0,0xda50,0x7552,0x56a0,0xabb7,0x25d0,0x92d0,0xcab5,
0xa950,0xb4a0,0xbaa4,0xad50,0x55d9,0x4ba0,0xa5b0,0x5176,0x52bf,0xa930,
0x7954,0x6aa0,0xad50,0x5b52,0x4b60,0xa6e6,0xa4e0,0xd260,0xea65,0xd530,
0x5aa0,0x76a3,0x96d0,0x4afb,0x4ad0,0xa4d0,0xd0b6,0xd25f,0xd520,0xdd45,
0xb5a0,0x56d0,0x55b2,0x49b0,0xa577,0xa4b0,0xaa50,0xb255,0x6d2f,0xada0,
0x4b63,0x937f,0x49f8,0x4970,0x64b0,0x68a6,0xea5f,0x6b20,0xa6c4,0xaaef,
0x92e0,0xd2e3,0xc960,0xd557,0xd4a0,0xda50,0x5d55,0x56a0,0xa6d0,0x55d4,
0x52d0,0xa9b8,0xa950,0xb4a0,0xb6a6,0xad50,0x55a0,0xaba4,0xa5b0,0x52b0,
0xb273,0x6930,0x7337,0x6aa0,0xad50,0x4b55,0x4b6f,0xa570,0x54e4,0xd260,
0xe968,0xd520,0xdaa0,0x6aa6,0x56df,0x4ae0,0xa9d4,0xa4d0,0xd150,0xf252,
0xd520);

var solarMonth=new Array(31,28,31,30,31,30,31,31,30,31,30,31);
var Gan=new Array("��","��","��","��","��","��","��","��","��","��");
var Zhi=new Array("��","��","��","î","��","��","��","δ","��","��","��","��");
var Animals=new Array("��","ţ","��","��","��","��","��","��","��","��","��","��");
var solarTerm = new Array("С��","��","����","��ˮ","����","����","����","����","����","С��","â��","����","С��","����","����","����","��¶","���","��¶","˪��","����","Сѩ","��ѩ","����");
var sTermInfo = new Array(0,21208,42467,63836,85337,107014,128867,150921,173149,195551,218072,240693,263343,285989,308563,331033,353350,375494,397447,419210,440795,462224,483532,504758);
var nStr1 = new Array('��','һ','��','��','��','��','��','��','��','��','ʮ');
var nStr2 = new Array('��','ʮ','إ','ئ','��');
var monthName = new Array("һ��","����","����","����","����","����","����","����","����","ʮ��","ʮһ��","ʮ����");

//�������� *��ʾ�ż���
var sFtv = new Array(
"0101*����Ԫ��",
"0202 ����ʪ����",
"0207 ������Ԯ�Ϸ���",
"0210 ���������",
"0214 ���˽�",
"0301 ���ʺ�����",
"0303 ȫ��������",
"0308 ���ʸ�Ů��",
"0312 ֲ���� ����ɽ����������",
"0314 ���ʾ�����",
"0315 ����������Ȩ����",
"0317 �й���ҽ�� ���ʺ�����",
"0321 ����ɭ���� �����������ӹ�����",
"0321 ���������",
"0322 ����ˮ��",
"0323 ����������",
"0324 ������ν�˲���",
"0325 ȫ����Сѧ����ȫ������",
"0330 ����˹̹������",
"0401 ���˽� ȫ�����������˶���(����) ˰��������(����)",
"0407 ����������",
"0422 ���������",
"0423 ����ͼ��Ͱ�Ȩ��",
"0424 �Ƿ����Ź�������",
"0501 �����Ͷ���",
"0504 �й����������",
"0505 ��ȱ����������",
"0508 �����ʮ����",
"0512 ���ʻ�ʿ��",
"0515 ���ʼ�ͥ��",
"0517 ���������",
"0518 ���ʲ������",
"0520 ȫ��ѧ��Ӫ����",
"0523 ����ţ����",
"0531 ����������", 
"0601 ���ʶ�ͯ��",
"0605 ���绷����",
"0606 ȫ��������",
"0617 ���λ�Į���͸ɺ���",
"0623 ���ʰ���ƥ����",
"0625 ȫ��������",
"0626 ���ʷ���Ʒ��",
"0701 �й������������� ���罨����",
"0702 �������������� ��Ʒ�ƽ�վ(http://www.21softs.com/)��ʽ���⿪�ż�����",
"0707 �й�������ս��������",
"0711 �����˿���",
"0730 ���޸�Ů��",
"0801 �й�������",
"0808 �й����ӽ�(�ְֽ�)",
"0815 �ձ���ʽ����������Ͷ����",
"0908 ����ɨä�� �������Ź�������",
"0910 ��ʦ��",
"0914 ������������ ÷���� ����^o^",
"0916 ���ʳ����㱣����",
"0918 �š�һ���±������",
"0920 ���ʰ�����",
"0927 ����������",
"1001*����� ���������� �������˽�",
"1001 ����������",
"1002 ���ʺ�ƽ���������ɶ�����",
"1004 ���綯����",
"1008 ȫ����Ѫѹ��",
"1008 �����Ӿ���",
"1009 ���������� ���������",
"1010 �������������� ���羫��������",
"1013 ���籣���� ���ʽ�ʦ��",
"1014 �����׼��",
"1015 ����ä�˽�(�����Ƚ�)",
"1016 ������ʳ��",
"1017 ��������ƶ����",
"1022 ���紫ͳҽҩ��",
"1024 ���Ϲ��� ���緢չ��Ϣ��",
"1031 �����ڼ���",
"1107 ʮ������������������",
"1108 �й�������",
"1109 ȫ��������ȫ����������",
"1110 ���������",
"1111 ���ʿ�ѧ���ƽ��(����������һ��)",
"1112 ����ɽ����������",
"1114 ����������",
"1117 ���ʴ�ѧ���� ����ѧ����",
"1121 �����ʺ��� ���������",
"1129 ������Ԯ����˹̹���������",
"1201 ���簬�̲���",
"1203 ����м�����",
"1205 ���ʾ��ú���ᷢչ־Ը��Ա��",
"1208 ���ʶ�ͯ������",
"1209 ����������",
"1210 ������Ȩ��",
"1212 �����±������",
"1213 �Ͼ�����ɱ(1937��)�����գ�����Ѫ��ʷ��",
"1221 ����������",
"1224 ƽ��ҹ",
"1225 ʥ����",
"1229 ���������������");

//ĳ�µĵڼ������ڼ��� 5,6,7,8 ��ʾ������ 1,2,3,4 �����ڼ�
var wFtv = new Array(
"0110 ������",
"0150 ���������", //һ�µ����һ�������գ��µ�����һ�������գ�
"0520 ����ĸ�׽�",
"0530 ȫ��������",
"0630 ���׽�",
"0911 �Ͷ���",
"0932 ���ʺ�ƽ��",
"0940 �������˽� �����ͯ��",
"0950 ���纣����",
"1011 ����ס����",
"1013 ���ʼ�����Ȼ�ֺ���(������)",
"1144 �ж���");

//ũ������
var lFtv = new Array(
"0101*����",
"0115 Ԫ����",
"0202 ��̧ͷ��",
"0323 �������� (����ʥĸ����)",
"0505 �����",
"0707 �����й����˽�",
"0815 �����",
"0909 ������",
"1208 ���˽�",
"1223 ���˽�",
"0100*��Ϧ");

//����ʱ������
var timeData = {
"Asia               ����": {   //----------------------------------------------
"Brunei             ����    ":["+0800","","˹��ͼ�����"],
"Burma              ���    ":["+0630","","����"],
"Cambodia           ����կ  ":["+0700","","���"],
"China              �й�    ":["+0800","","���������졢�Ϻ������"],
"China(HK,Macau)    �й�    ":["+0800","","��ۡ���������"],
"China(TaiWan)      �й�    ":["+0800","","̨��������"],
"China(Urumchi)     �й�    ":["+0700","","��³ľ��"],
"Indonesia          ӡ��    ":["+0700","","�żӴ�"],
"Japan              �ձ�    ":["+0900","","���������桢����"],
"Korea              ����    ":["+0900","","����"],
"Laos               ����    ":["+0700","","����"],
"Malaysia           ��������":["+0800","","��¡��"],
"Mongolia           �ɹ�    ":["+0800","03L03|09L03","�������С�����"],
"Philippines        ���ɱ�  ":["+0800","04F53|10F53","������"],
"Russia(Anadyr)     ����˹  ":["+1300","03L03|10L03","���ɵ¶���"],
"Russia(Kamchatka)  ����˹  ":["+1200","03L03|10L03","����Ӱ뵺"],
"Russia(Magadan)    ����˹  ":["+1100","03L03|10L03","���ӵ�"],
"Russia(Vladivostok)����˹  ":["+1000","03L03|10L03","��������˹�п�(������)"],
"Russia(Yakutsk)    ����˹  ":["+0900","03L03|10L03","�ſ�Ŀ�"],
"Singapore          �¼���  ":["+0800","","�¼���"],
"Thailand           ̩��    ":["+0700","","����"],
"Vietnam            Խ��    ":["+0700","","����"]
},
"ME, India pen.     �ж���ӡ�Ȱ뵺": {   //------------------------------------
"Afghanistan        ������  ":["+0430","","������"],
"Arab Emirates      ����������������":["+0400","","��������"],
"Bahrain            ����    ":["+0300","","������"],
"Bangladesh         �ϼ���  ":["+0600","","�￨"],
"Bhutan             ����    ":["+0600","","͢��"],
"Cyprus             ����·˹":["+0200","","�������"],
"Georgia            ������  ":["+0500","","�ڱ���˹"],
"India              ӡ��    ":["+0530","","�µ�����򡢼Ӷ�����"],
"Iran               ����    ":["+0330","04 13|10 13","�º���"],
"Iraq               ������  ":["+0300","04 13|10 13","�͸��"],
"Israel             ��ɫ�С�����˹̹":["+0200","04F53|09F53","Ү·����"],
"Jordan             Լ��    ":["+0200","","����"],
"Kuwait             ������  ":["+0300","","�����س�"],
"Lebanon            �����  ":["+0200","03L03|10L03","��³��"],
"Maldives           ��������":["+0500","","����"],
"Nepal              �Ჴ��  ":["+0545","","�ӵ�����"],
"Oman               ����    ":["+0400","","��˹����"],
"Pakistan           �ͻ�˹̹":["+0500","","�����桢��˹����"],
"Qatar              ������  ":["+0300","","���"],
"Saudi Arabia       ɳ�ذ�����":["+0300","","���ŵ�"],
"Sri Lanka          ˹������":["+0600","","������"],
"Syria              ������  ":["+0200","04 13|10 13","����ʿ��"],
"Tajikistan         ������˹̹":["+0500","","���б�"],
"Turkey             ������  ":["+0200","","��˹̹��"],
"Turkmenistan       ������˹̹":["+0500","","��ʲ���͵�"],
"Uzbekistan         ���ȱ��˹̹":["+0500","","��ʲ��"],
"Yemen              Ҳ��    ":["+0300","","����"]
},
"North Europe       ��ŷ": {   //----------------------------------------------
"Denmark            ����":["+0100","04F03|10L03","�籾����"],
"Finland            ����":["+0200","03L01|10L01","�ն�����"],
"Iceland            ����":["+0000","","�׿���δ��"],
"Norwegian          Ų��":["+0100","","��˹½"],
"Sweden             ���":["+0100","03L01|10L01","˹�¸��Ħ"]
},
"Eastern Europe     ��ŷ����ŷ": {   //----------------------------------------
"Armenia            ��������":["+0400","","������"],
"Austria            �µ���  ":["+0100","03L01|10L01","άҲ��"],
"Azerbaijan         �����ݽ�":["+0400","","�Ϳ�"],
"Czech              �ݿ�    ":["+0100","","������"],
"Estonia            ��ɳ����":["+0200","","����"],
"Germany            �¹�    ":["+0100","03L01|10L01","���֡�����"],
"Hungarian          ������  ":["+0100","","������˹"],
"Kazakhstan(Astana) ������˹̹":["+0600","","��˹���ɡ�����ľͼ"],
"Kazakhstan(Aqtobe) ������˹̹":["+0500","","�����б�"],
"Kazakhstan(Aqtau)  ������˹̹":["+0400","","����ͼ"],
"Kirghizia          ������˹":["+0500","","��˹����"],
"Latvia             ����ά��":["+0200","","���"],
"Lithuania          ������  ":["+0200","","ά��Ŧ˹"],
"Moldova            Ħ������":["+0200","","��ϣ����"],
"Poland             ����    ":["+0100","","��ɳ"],
"Rumania            ��������":["+0200","","������˹��"],
"Russia(Moscow)     ����˹  ":["+0300","03L03|10L03","Ī˹��"],
"Russia(Volgograd)  ����˹  ":["+0300","03L03|10L03","�����Ӹ���"],
"Slovakia           ˹�工��":["+0100","","������˹����"],
"Switzerland        ��ʿ    ":["+0100","","������"],
"Ukraine            �ڿ���  ":["+0200","","����"],
"Ukraine(Simferopol)�ڿ���  ":["+0300","","�����޲���"],
"Belarus            �׶���˹":["+0200","03L03|10L03","��˹��"]
},
"Western Europe     ��ŷ": {   //----------------------------------------------
"Belgium            ����ʱ ":["+0100","03L01|10L01","��³����"],
"France             ����   ":["+0100","03L01|10L01","����"],
"Ireland            ������ ":["+0000","03L01|10L01","������"],
"Monaco             Ħ�ɸ� ":["+0100","","Ħ�ɸ���"],
"Netherlands        ����   ":["+0100","03L01|10L01","��ķ˹�ص�"],
"Luxembourg         ¬ɭ�� ":["+0100","03L01|10L01","¬ɭ����"],
"United Kingdom     Ӣ��   ":["+0000","03L01|10L01","�׶ء�������"]
},
"South Europe       ��ŷ": { //------------------------------------------------
"Albania            ����������":["+0100","","������"],
"Bulgaria           ��������":["+0200","","������"],
"Greece             ϣ��    ":["+0200","03L01|10L01","�ŵ�"],
"Holy See           ������͢":["+0100","","��ٸ�"],
"Italy              �����  ":["+0100","03L01|10L01","����"],
"Malta              ������  ":["+0100","","������"],
"Portugal           ������  ":["+0000","03L01|10L01","��˹��"],
"San Marino         ʥ����ŵ":["+0100","","ʥ����ŵ"],
"Span               ������  ":["+0100","03L01|10L01","������"],
"Slovenia           ˹��������":["+0100","","¬��������"],
"Yugoslavia         ��˹����(����ά��)":["+0100","","����������"]
},
"North America      ������": {   //--------------------------------------------
"Canada(NST)        ���ô�":["-0330","04F02|10L02","Ŧ������ʥԼ������˹��"],
"Canada(AST)        ���ô�":["-0400","04F02|10L02","�����塢Pangnirtung"],
"Canada(EST)        ���ô�":["-0500","04F02|10L02","������"],
"Canada(CST)        ���ô�":["-0600","04F02|10L02","���ȼ{�����悡�Swift Current"],
"Canada(MST)        ���ô�":["-0700","04F02|10L02","ӡūά�ظ��塢�����ɶ١���ɭ��"],
"Canada(PST)        ���ô�":["-0800","04F02|10L02","�¸绪"],
"US(Eastern)        ����(����)":["-0500","04F02|10L02","��ʢ�١�ŦԼ"],
"US(Indiana)        ����      ":["-0500","","ӡ�ڰ���"],
"US(Central)        ����(�в�)":["-0600","04F02|10L02","֥�Ӹ�"],
"US(Mountain)       ����(ɽ��)":["-0700","04F02|10L02","����"],
"US(Arizona)        ����      ":["-0700","","����ɣ��"],
"US(Pacific)        ����(����)":["-0800","04F02|10L02","�ɽ�ɽ����ɼ��"],
"US(Alaska)         ����      ":["-0900","","����˹�ӡ���ŵ"]
},
"South America      ��������": {   //------------------------------------------
"Antigua & Barbuda  ����ϵ����Ͳ��ﵺ":["-0400","","ʥԼ��"],
"Argentina          ����͢  ":["-0300","","����ŵ˹����˹"],
"Bahamas            �͹���  ":["-0500","","��ɧ"],
"Barbados           �ͰͶ�˹��":["-0400","","�������(����)"],
"Belize             ����˹  ":["-0600","","����˹"],
"Bolivia            ����ά��":["-0400","","����˹"],
"Brazil(AST)        ����    ":["-0500","10F03|02L03","Porto Acre"],
"Brazil(EST)        ����    ":["-0300","10F03|02L03","�������ǡ���Լ����¬"],
"Brazil(FST)        ����    ":["-0200","10F03|02L03","ŵ����"],
"Brazil(WST)        ����    ":["-0400","10F03|02L03","���ǰ�"],
"Chilean            ����    ":["-0500","10F03|03F03","Hanga Roa"],
"Chilean            ����    ":["-0300","10F03|03F03","ʥ���Ǹ�"],
"Colombia           ���ױ���":["-0500","","�����"],
"Costa Rica         ��˹�����":["-0600","","ʥ����"],
"Cuba               �Ű�    ":["-0500","04 13|10L03","������"],
"Dominican          �������":["-0400","","ʥ������������"],
"Ecuador            ��϶��":["-0500","","����"],
"El Salvador        �����߶�":["-0600","","ʥ�����߶�"],
"Falklands          ������Ⱥ��":["-0300","09F03|04F03","ʷ����"],
"Guatemala          Σ������":["-0600","","Σ��������"],
"Haiti              ����    ":["-0500","","̫�Ӹ�"],
"Honduras           �鶼��˹":["-0600","","�ع����Ӷ���"],
"Jamaica            �����  ":["-0500","","��˹��"],
"Mexico(Mazatlan)   ī����  ":["-0700","","��������"],
"Mexico(�׶�)       ī����  ":["-0600","","ī�����"],
"Mexico(�ٻ���)     ī����  ":["-0800","","�ٻ���"],
"Nicaragua          �������":["-0500","","���ǹ�"],
"Panama             ������  ":["-0500","","��������"],
"Paraguay           ������  ":["-0400","10F03|02L03","����ɭ"],
"Peru               ��³    ":["-0500","","����"],
"Saint Kitts & Nevis ʥ���ĺ���ά˹":["-0400","","��˹�ض�(Basseterre)"],
"St. Lucia          ʥ¬����":["-0400","","��˹����"],
"St. Vincent & Grenadines ʥ��ɭ�غ͸����ɶ�˹":["-0400","","��˹��"],
"Suriname           ������":["-0300","","�������ﲩ(Paramaribo)"],
"Trinidad & Tobago  �������Ͷ�͸�":["-0400","","��������"],
"Uruguay            ������  ":["-0300","","�ɵ�ά����"],
"Venezuela          ί������":["-0400","","������˹"]
},
"Africa             ����": {   //----------------------------------------------
"Algeria            ����������":["+0100","","��������"],
"Angola             ������  ":["+0100","","�ް���"],
"Benin              ����    ":["+0100","","�¸�"],
"Botswana           ��������":["+0200","","��������"],
"Burundi            ��¡��  ":["+0200","","��������"],
"Cameroon           ����¡  ":["+0100","","���µ�"],
"Cape Verde         ��½�  ":["-0100","","������"],
"Central African    �зǹ��͹�":["+0100","","�༪"],
"Chad               է��    ":["+0100","","����÷����"],
"Congo              �չ�(��)":["+0100","","������ά��"],
"Djibouti           ������  ":["+0300","","������"],
"Egypt              ����    ":["+0200","04L53|09L43","����"],
"Equatorial Guinea  ���������":["+0100","","������"],
"Ethiopia           ����������":["+0300","","�ǵ�˹�Ǳ���"],
"Gabon              ����    ":["+0100","","����ά��"],
"Gambia             �Ա���  ":["+0000","","�����"],
"Ghana              ����    ":["+0000","","������"],
"Guinea             ������  ":["+0000","","���ɿ���"],
"Ivory Coast        ��������":["+0000","","�����á���������"],
"Kenya              ������  ":["+0300","","���ޱ�"],
"Lesotho            ������  ":["+0200","","����¬"],
"Liberia            ��������":["+0000","","����ά��"],
"Madagascar         �����˹��":["+0300","","����������"],
"Malawi             ����ά  ":["+0200","","��¡��"],
"Mali               ����    ":["+0000","","������"],
"Mauritania         ë��������":["+0000","","Ŭ�߿�Ф��"],
"Mauritius          ë����˹":["+0400","","·�׸�"],
"Morocco            Ħ���  ":["+0000","","����������"],
"Mozambique         Īɣ�ȿ�":["+0200","","������"],
"Namibia            ���ױ���":["+0200","09F03|04F03","�µúͿ�"],
"Niger              ���ն�  ":["+0100","","������"],
"Nigeria            ��������":["+0100","","������"],
"Rwanda             ¬����  ":["+0200","","������"],
"Sao Tome           ʥ����  ":["+0000","","ʥ����"],
"Senegal            ���ڼӶ�":["+0000","","�￨��"],
"Sierra Leone       ʨ��ɽ��":["+0000","","���ɳ�"],
"Somalia            ������  ":["+0300","","Ħ�ӵ�ɳ"],
"South Africa       �Ϸ�    ":["+0200","","���նء�����������"],
"Sudan              �յ�    ":["+0200","","������"],
"Tanzania           ̹ɣ����":["+0300","","����˹����ķ"],
"Togo               ���    ":["+0000","","����¡"],
"Tunisia            ͻ��˹  ":["+0100","","ͻ��˹��"],
"Uganda             �ڸɴ�  ":["+0300","","������"],
"Zaire              ������(�չ���)  ":["+0100","","��ɳ��"],
"Zambia             �ޱ���  ":["+0200","","¬����"],
"Zimbabwe           ��Ͳ�Τ":["+0200","","������"]
},
"Oceania            ������": { //----------------------------------------------
"American Samoa(US) ������Ħ��(��)":["-1100","","����������"],
"Aus.(Adelaide)     �Ĵ�����  ":["+0930","10F03|03F03","�����׵�"],
"Aus.(Brisbane)     �Ĵ�����  ":["+1000","10F03|03F03","����˹��"],
"Aus.(Darwin)       �Ĵ�����  ":["+0930","10F03|03F03","�����"],
"Aus.(Hobart)       �Ĵ�����  ":["+1000","10F03|03F03","�ɲ���"],
"Aus.(Perth)        �Ĵ�����  ":["+0800","10F03|03F03","��˼"],
"Aus.(Sydney)       �Ĵ�����  ":["+1000","10F03|03F03","Ϥ��"],
"Cook Islands(NZ)   ���Ⱥ��(������)  ":["-1000","","����³��"],
"Eniwetok           �������п˵�":["-1200","","�������п˵�"],
"Fiji               쳼�      ":["+1200","11F03|02L03","����"],
"Guam               �ص�      ":["+1000","","��������"],
"Hawaii(US)         ������(��)":["-1000","","̴��ɽ"],
"Kiribati           �����˹  ":["+1100","","������"],
//"Mariana Islands    ���ൺ    ":["","","���ൺ"],
"Marshall Is.       ���ܶ�Ⱥ��":["+1200","","������"],
"Micronesia         �ܿ�������������":["+1000","","��������(Palikir)"],
"Midway Is.(US)     ��;��(��)":["-1100","","��;��"],
"Nauru Rep.         �³���͹�":["+1200","","����"],
"New Calednia(FR)   �¿��������(��)":["+1100","","Ŭ����"],
"New Zealand        ������    ":["+1200","10F03|04F63","�¿���"],
"New Zealand(CHADT) ������    ":["+1245","10F03|04F63","�����"],
"Niue(NZ)           Ŧ��(��)  ":["-1100","","�����(Alofi)"],
"Nor. Mariana Is.   ����������Ⱥ��(��)":["+1000","","���ൺ"],
"Palau              ����Ⱥ��(����Ⱥ��)      ":["+0900","","���޶�"],
"Papua New Guinea   �Ͳ����¼�����":["+1000","","Ī��˹�ȸ�"],
"Pitcairn Is.(UK)   Ƥ�ؿ˶�Ⱥ��(Ӣ)":["-0830","","�ǵ�˹��"],
"Polynesia(FR)      ����������(��)":["-1000","","�ͱȵ١���ϣ��"],
"Solomon Is.        ������Ⱥ��":["+1100","","��������"],
"Tahiti             ��ϣ��  ":["-1000","","������"],
"Tokelau(NZ)        �п���Ⱥ��(��)":["-1100","","Ŭ��ŵŬ����������������"],
"Tonga              ����    ":["+1300","10F63|04F63","Ŭ�Ⱒ�巨"],
"Tuvalu             ͼ��¬  ":["+1200","","���ɸ���"],
"Vanuatu            ��Ŭ��ͼ(�ºղ����Ⱥ��)":["+1100","","ά����"],
"Western Samoa      ����Ħ��":["-1100","","��Ƥ��"],
"���ʻ�����                   ":["-1200","","���ʻ�����"]
}
};



/*****************************************************************************
                                      ���ڼ���
*****************************************************************************/

//====================================== ����ũ�� y���������
function lYearDays(y) {
 var i, sum = 348;
 for(i=0x8000; i>0x8; i>>=1) sum += (lunarInfo[y-1900] & i)? 1: 0;
 return(sum+leapDays(y));
}

//====================================== ����ũ�� y�����µ�����
function leapDays(y) {
 if(leapMonth(y)) return( (lunarInfo[y-1899]&0xf)==0xf? 30: 29);
 else return(0);
}

//====================================== ����ũ�� y�����ĸ��� 1-12 , û�򷵻� 0
function leapMonth(y) {
 var lm = lunarInfo[y-1900] & 0xf;
 return(lm==0xf?0:lm);
}

//====================================== ����ũ�� y��m�µ�������
function monthDays(y,m) {
 return( (lunarInfo[y-1900] & (0x10000>>m))? 30: 29 );
}


//====================================== ���ũ��, �������ڿؼ�, ����ũ�����ڿؼ�
//                                       �ÿؼ������� .year .month .day .isLeap
function Lunar(objDate) {

   var i, leap=0, temp=0;
   var offset   = (Date.UTC(objDate.getFullYear(),objDate.getMonth(),objDate.getDate()) - Date.UTC(1900,0,31))/86400000;

   for(i=1900; i<2100 && offset>0; i++) { temp=lYearDays(i); offset-=temp; }

   if(offset<0) { offset+=temp; i--; }

   this.year = i;

   leap = leapMonth(i); //���ĸ���
   this.isLeap = false;

   for(i=1; i<13 && offset>0; i++) {
      //����
      if(leap>0 && i==(leap+1) && this.isLeap==false)
         { --i; this.isLeap = true; temp = leapDays(this.year); }
      else
         { temp = monthDays(this.year, i); }

      //�������
      if(this.isLeap==true && i==(leap+1)) this.isLeap = false;

      offset -= temp;
   }

   if(offset==0 && leap>0 && i==leap+1)
      if(this.isLeap)
         { this.isLeap = false; }
      else
         { this.isLeap = true; --i; }

   if(offset<0){ offset += temp; --i; }

   this.month = i;
   this.day = offset + 1;
}

//==============================���ع��� y��ĳm+1�µ�����
function solarDays(y,m) {
   if(m==1)
      return(((y%4 == 0) && (y%100 != 0) || (y%400 == 0))? 29: 28);
   else
      return(solarMonth[m]);
}
//============================== ���� offset ���ظ�֧, 0=����
function cyclical(num) {
   return(Gan[num%10]+Zhi[num%12]);
}

//============================== ��������
function calElement(sYear,sMonth,sDay,week,lYear,lMonth,lDay,isLeap,cYear,cMonth,cDay) {

      this.isToday    = false;
      //���
      this.sYear      = sYear;   //��Ԫ��4λ����
      this.sMonth     = sMonth;  //��Ԫ������
      this.sDay       = sDay;    //��Ԫ������
      this.week       = week;    //����, 1������
      //ũ��
      this.lYear      = lYear;   //��Ԫ��4λ����
      this.lMonth     = lMonth;  //ũ��������
      this.lDay       = lDay;    //ũ��������
      this.isLeap     = isLeap;  //�Ƿ�Ϊũ������?
      //����
      this.cYear      = cYear;   //����, 2������
      this.cMonth     = cMonth;  //����, 2������
      this.cDay       = cDay;    //����, 2������

      this.color      = '';

      this.lunarFestival = ''; //ũ������
      this.solarFestival = ''; //��������
      this.solarTerms    = ''; //����
}

//===== ĳ��ĵ�n������Ϊ����(��0С������)
function sTerm(y,n) {
   var offDate = new Date( ( 31556925974.7*(y-1900) + sTermInfo[n]*60000  ) + Date.UTC(1900,0,6,2,5) );
   return(offDate.getUTCDate());
}




//============================== ���������ؼ� (y��,m+1��)
/*
����˵��: ���������µ��������Ͽؼ�

ʹ�÷�ʽ: OBJ = new calendar(��,��������);

  OBJ.length      ���ص��������
  OBJ.firstWeek   ���ص���һ������

  �� OBJ[����].�������� ����ȡ�ø���ֵ

  OBJ[����].isToday  �����Ƿ�Ϊ���� true �� false

  ���� OBJ[����] ���Բμ� calElement() �е�ע��
*/
function calendar(y,m) {

   var sDObj, lDObj, lY, lM, lD=1, lL, lX=0, tmp1, tmp2, tmp3;
   var cY, cM, cD; //����,����,����
   var lDPOS = new Array(3);
   var n = 0;
   var firstLM = 0;

   sDObj = new Date(y,m,1,0,0,0,0);    //����һ������

   this.length    = solarDays(y,m);    //������������
   this.firstWeek = sDObj.getDay();    //��������1�����ڼ�

   ////////���� 1900��������Ϊ������(60����36)
   if(m<2) cY=cyclical(y-1900+36-1);
   else cY=cyclical(y-1900+36);
   var term2=sTerm(y,2); //��������

   ////////���� 1900��1��С����ǰΪ ������(60����12)
   var firstNode = sTerm(y,m*2) //���ص��¡��ڡ�Ϊ���տ�ʼ
   cM = cyclical((y-1900)*12+m+12);

   //����һ���� 1900/1/1 �������
   //1900/1/1�� 1970/1/1 ���25567��, 1900/1/1 ����Ϊ������(60����10)
   var dayCyclical = Date.UTC(y,m,1,0,0,0,0)/86400000+25567+10;

   for(var i=0;i<this.length;i++) {

      if(lD>lX) {
         sDObj = new Date(y,m,i+1);    //����һ������
         lDObj = new Lunar(sDObj);     //ũ��
         lY    = lDObj.year;           //ũ����
         lM    = lDObj.month;          //ũ����
         lD    = lDObj.day;            //ũ����
         lL    = lDObj.isLeap;         //ũ���Ƿ�����
         lX    = lL? leapDays(lY): monthDays(lY,lM); //ũ���������һ��

         if(n==0) firstLM = lM;
         lDPOS[n++] = i-lD+1;
      }

      //�������������·ֵ�����, ������Ϊ��
      if(m==1 && (i+1)==term2) cY=cyclical(y-1900+36);
      //����������, �ԡ��ڡ�Ϊ��
      if((i+1)==firstNode) cM = cyclical((y-1900)*12+m+13);
      //����
      cD = cyclical(dayCyclical+i);

      //sYear,sMonth,sDay,week,
      //lYear,lMonth,lDay,isLeap,
      //cYear,cMonth,cDay
      this[i] = new calElement(y, m+1, i+1, nStr1[(i+this.firstWeek)%7],
                               lY, lM, lD++, lL,
                               cY ,cM, cD );
   }

   //����
   tmp1=sTerm(y,m*2  )-1;
   tmp2=sTerm(y,m*2+1)-1;
   this[tmp1].solarTerms = solarTerm[m*2];
   this[tmp2].solarTerms = solarTerm[m*2+1];
   if(m==3) this[tmp1].color = 'red'; //������ɫ

   //��������
   for(i in sFtv)
      if(sFtv[i].match(/^(\d{2})(\d{2})([\s\*])(.+)$/))
         if(Number(RegExp.$1)==(m+1)) {
            this[Number(RegExp.$2)-1].solarFestival += RegExp.$4 + ' ';
            if(RegExp.$3=='*') this[Number(RegExp.$2)-1].color = 'red';
         }

   //���ܽ���
   for(i in wFtv)
      if(wFtv[i].match(/^(\d{2})(\d)(\d)([\s\*])(.+)$/))
         if(Number(RegExp.$1)==(m+1)) {
            tmp1=Number(RegExp.$2);
            tmp2=Number(RegExp.$3);
            if(tmp1<5)
               this[((this.firstWeek>tmp2)?7:0) + 7*(tmp1-1) + tmp2 - this.firstWeek].solarFestival += RegExp.$5 + ' ';
            else {
               tmp1 -= 5;
               tmp3 = (this.firstWeek+this.length-1)%7; //�������һ������?
               this[this.length - tmp3 - 7*tmp1 + tmp2 - (tmp2>tmp3?7:0) - 1 ].solarFestival += RegExp.$5 + ' ';
            }
         }

   //ũ������
   for(i in lFtv)
      if(lFtv[i].match(/^(\d{2})(.{2})([\s\*])(.+)$/)) {
         tmp1=Number(RegExp.$1)-firstLM;
         if(tmp1==-11) tmp1=1;
         if(tmp1 >=0 && tmp1<n) {
            tmp2 = lDPOS[tmp1] + Number(RegExp.$2) -1;
            if( tmp2 >= 0 && tmp2<this.length && this[tmp2].isLeap!=true) {
               this[tmp2].lunarFestival += RegExp.$4 + ' ';
               if(RegExp.$3=='*') this[tmp2].color = 'red';
            }
         }
      }


   //�����ֻ������3��4��
   if(m==2 || m==3) {
      var estDay = new easter(y);
      if(m == estDay.m)
         this[estDay.d-1].solarFestival = this[estDay.d-1].solarFestival+' ����� Easter Sunday';
   }


   if(m==2) this[20].solarFestival = this[20].solarFestival+unescape('%20%u6D35%u8CE2%u751F%u65E5');

   //��ɫ������
   if((this.firstWeek+12)%7==5)
      this[12].solarFestival += '��ɫ������';

   //����
   if(y==tY && m==tM) this[tD-1].isToday = true;
}

//======================================= ���ظ���ĸ����(���ֺ��һ�������ܺ�ĵ�һ����)
function easter(y) {

   var term2=sTerm(y,5); //ȡ�ô�������
   var dayTerm2 = new Date(Date.UTC(y,2,term2,0,0,0,0)); //ȡ�ô��ֵĹ������ڿؼ�(����һ��������3��)
   var lDayTerm2 = new Lunar(dayTerm2); //ȡ��ȡ�ô���ũ��

   if(lDayTerm2.day<15) //ȡ���¸���Բ���������
      var lMlen= 15-lDayTerm2.day;
   else
      var lMlen= (lDayTerm2.isLeap? leapDays(y): monthDays(y,lDayTerm2.month)) - lDayTerm2.day + 15;

   //һ����� 1000*60*60*24 = 86400000 ����
   var l15 = new Date(dayTerm2.getTime() + 86400000*lMlen ); //�����һ����ԲΪ��������
   var dayEaster = new Date(l15.getTime() + 86400000*( 7-l15.getUTCDay() ) ); //����¸�����

   this.m = dayEaster.getUTCMonth();
   this.d = dayEaster.getUTCDate();

}

//====================== ��������
function cDay(d){
   var s;

   switch (d) {
      case 10:
         s = '��ʮ'; break;
      case 20:
         s = '��ʮ'; break;
         break;
      case 30:
         s = '��ʮ'; break;
         break;
      default :
         s = nStr2[Math.floor(d/10)];
         s += nStr1[d%10];
   }
   return(s);
}

///////////////////////////////////////////////////////////////////////////////

var cld;

function drawCld(SY,SM) {
   var i,sD,s,size;
   cld = new calendar(SY,SM);

   if(SY>1874 && SY<1909) yDisplay = '����' + (((SY-1874)==1)?'Ԫ':SY-1874);
   if(SY>1908 && SY<1912) yDisplay = '��ͳ' + (((SY-1908)==1)?'Ԫ':SY-1908);
   if(SY>1911 && SY<1950) yDisplay = '���' + (((SY-1911)==1)?'Ԫ':SY-1911);
   if(SY>1948) yDisplay = '����' + (((SY-1949)==1)?'Ԫ':SY-1949);

   GZ.innerHTML = yDisplay +'�� ũ�� ' + cyclical(SY-1900+36) + '�� ��'+Animals[(SY-4)%12]+'�꡿';

   YMBG.innerHTML = "&nbsp;" + SY + "<BR>&nbsp;" + monthName[SM];

   for(i=0;i<42;i++) {

      sObj=eval('SD'+ i);
      lObj=eval('LD'+ i);

      sObj.className = '';

      sD = i - cld.firstWeek;

      if(sD>-1 && sD<cld.length) { //������
         sObj.innerHTML = sD+1;

         if(cld[sD].isToday) sObj.className = 'todyaColor'; //������ɫ

         sObj.style.color = cld[sD].color; //����������ɫ

         if(cld[sD].lDay==1) //��ʾũ����
            lObj.innerHTML = '<b>'+(cld[sD].isLeap?'��':'') + cld[sD].lMonth + '��' + (monthDays(cld[sD].lYear,cld[sD].lMonth)==29?'С':'��')+'</b>';
         else //��ʾũ����
            lObj.innerHTML = cDay(cld[sD].lDay);

         s=cld[sD].lunarFestival;
         if(s.length>0) { //ũ������
            if(s.length>6) s = s.substr(0, 4)+'...';
            s = s.fontcolor('red');
         }
         else { //��������
            s=cld[sD].solarFestival;
            if(s.length>0) {
               size = (s.charCodeAt(0)>0 && s.charCodeAt(0)<128)?8:4;
               if(s.length>size+2) s = s.substr(0, size)+'...';
               s=(s=='��ɫ������')?s.fontcolor('black'):s.fontcolor('blue');
            }
            else { //إ�Ľ���
               s=cld[sD].solarTerms;
               if(s.length>0) s = s.fontcolor('limegreen');
            }
         }

         if(cld[sD].solarTerms=='����') s = '������'.fontcolor('red');
         if(cld[sD].solarTerms=='â��') s = 'â��'.fontcolor('red');
         if(cld[sD].solarTerms=='����') s = '����'.fontcolor('red');
         if(cld[sD].solarTerms=='����') s = '����'.fontcolor('red');



         if(s.length>0) lObj.innerHTML = s;

      }
      else { //������
         sObj.innerHTML = '';
         lObj.innerHTML = '';
      }
   }
}


function changeCld() {
   var y,m;
   y=CLD.SY.selectedIndex+1900;
   m=CLD.SM.selectedIndex;
   drawCld(y,m);
}

function pushBtm(K) {
 switch (K){
    case 'YU' :
       if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
       break;
    case 'YD' :
       if(CLD.SY.selectedIndex<200) CLD.SY.selectedIndex++;
       break;
    case 'MU' :
       if(CLD.SM.selectedIndex>0) {
          CLD.SM.selectedIndex--;
       }
       else {
          CLD.SM.selectedIndex=11;
          if(CLD.SY.selectedIndex>0) CLD.SY.selectedIndex--;
       }
       break;
    case 'MD' :
       if(CLD.SM.selectedIndex<11) {
          CLD.SM.selectedIndex++;
       }
       else {
          CLD.SM.selectedIndex=0;
          if(CLD.SY.selectedIndex<200) CLD.SY.selectedIndex++;
       }
       break;
    default :
       CLD.SY.selectedIndex=tY-1900;
       CLD.SM.selectedIndex=tM;
 }
 changeCld();
}

var Today = new Date();
var tY = Today.getFullYear();
var tM = Today.getMonth();
var tD = Today.getDate();
//////////////////////////////////////////////////////////////////////////////

var width = "130";
var offsetx = 2;
var offsety = 8;

var x = 0;
var y = 0;
var snow = 0;
var sw = 0;
var cnt = 0;

var dStyle;
document.onmousemove = mEvn;

//��ʾ��ϸ��������
function mOvr(v) {
   var s,festival;
   var sObj=eval('SD'+ v);
   var d=sObj.innerHTML-1;

      //sYear,sMonth,sDay,week,
      //lYear,lMonth,lDay,isLeap,
      //cYear,cMonth,cDay

   if(sObj.innerHTML!='') {

      sObj.style.cursor = 's-resize';

      if(cld[d].solarTerms == '' && cld[d].solarFestival == '' && cld[d].lunarFestival == '')
         festival = '';
      else
         festival = '<TABLE WIDTH=100% BORDER=0 CELLPADDING=2 CELLSPACING=0 BGCOLOR="#CCFFCC"><TR><TD>'+
         '<FONT COLOR="#000000" STYLE="font-size:9pt;">'+cld[d].solarTerms + ' ' + cld[d].solarFestival + ' ' + cld[d].lunarFestival+'</FONT></TD>'+
         '</TR></TABLE>';

      s= '<TABLE WIDTH="130" BORDER=0 CELLPADDING="2" CELLSPACING=0 BGCOLOR="#000066" style="filter:Alpha(opacity=80)"><TR><TD>' +
         '<TABLE WIDTH=100% BORDER=0 CELLPADDING=0 CELLSPACING=0><TR><TD ALIGN="right"><FONT COLOR="#ffffff" STYLE="font-size:9pt;">'+
         cld[d].sYear+' �� '+cld[d].sMonth+' �� '+cld[d].sDay+' ��<br>����'+cld[d].week+'<br>'+
         '<font color="violet">ũ��'+(cld[d].isLeap?'�� ':' ')+cld[d].lMonth+' �� '+cld[d].lDay+' ��</font><br>'+
         '<font color="yellow">'+cld[d].cYear+'�� '+cld[d].cMonth+'�� '+cld[d].cDay + '��</font>'+
         '</FONT></TD></TR></TABLE>'+ festival +'</TD></TR></TABLE>';

      document.all["detail"].innerHTML = s;

      if (snow == 0) {
         dStyle.left = x+offsetx-(width/2);
         dStyle.top = y+offsety;
         dStyle.visibility = "visible";
         snow = 1;
      }
   }
}

//�����ϸ��������
function mOut() {
   if ( cnt >= 1 ) { sw = 0; }
   if ( sw == 0 ) { snow = 0; dStyle.visibility = "hidden";}
   else cnt++;
}

//ȡ��λ��
function mEvn() {
   x=event.x;
   y=event.y;
   if (document.body.scrollLeft)
      {x=event.x+document.body.scrollLeft; y=event.y+document.body.scrollTop;}
   if (snow){
      dStyle.left = x+offsetx-(width/2);
      dStyle.top = y+offsety;
   }
}

/*****************************************************************************
                                  ����ʱ�����
*****************************************************************************/
var OneHour = 60*60*1000;
var OneDay = OneHour*24;
var TimezoneOffset = Today.getTimezoneOffset()*60*1000;

function showUTC(objD) {
   var dn,s;
   var hh = objD.getUTCHours();
   var mm = objD.getUTCMinutes();
   var ss = objD.getUTCSeconds();
   s = objD.getUTCFullYear() + "��" + (objD.getUTCMonth() + 1) + "��" + objD.getUTCDate() +"�� ("+ nStr1[objD.getUTCDay()] +")";

   if(hh>12) { hh = hh-12; dn = '����'; }
   else dn = '����';

   if(hh<10) hh = '0' + hh;
   if(mm<10) mm = '0' + mm;
   if(ss<10) ss = '0' + ss;

   s += " " + dn + ' ' + hh + ":" + mm + ":" + ss;
   return(s);
}

function showLocale(objD) {
   var dn,s;
   var hh = objD.getHours();
   var mm = objD.getMinutes();
   var ss = objD.getSeconds();
   s = objD.getFullYear() + "��" + (objD.getMonth() + 1) + "��" + objD.getDate() +"�� ("+ nStr1[objD.getDay()] +")";

   if(hh>12) { hh = hh-12; dn = '����'; }
   else dn = '����';

   if(hh<10) hh = '0' + hh;
   if(mm<10) mm = '0' + mm;
   if(ss<10) ss = '0' + ss;

   s += " " + dn + ' ' + hh + ":" + mm + ":" + ss;
   return(s);
}

//����ʱ���ִ�, ����ƫ��֮��������
function parseOffset(s) {
   var sign,hh,mm,v;
   sign = s.substr(0,1)=='-'?-1:1;
   hh = Math.floor(s.substr(1,2));
   mm = Math.floor(s.substr(3,2));
   v = sign*(hh*60+mm)*60*1000;
   return(v);
}

//����UTC���ڿؼ� (��,��-1,�ڼ������ڼ�,����)
function getWeekDay(y,m,nd,w,h){
   var d,d2,w1;
   if(nd>0){
      d = new Date(Date.UTC(y, m, 1));
      w1 = d.getUTCDay();
      d2 = new Date( d.getTime() + ((w<w1? w+7-w1 : w-w1 )+(nd-1)*7   )*OneDay + h*OneHour);
   }
   else {
      nd = Math.abs(nd);
      d = new Date( Date.UTC(y, m+1, 1)  - OneDay );
      w1 = d.getUTCDay();
      d2 = new Date( d.getTime() + (  (w>w1? w-7-w1 : w-w1 )-(nd-1)*7   )*OneDay + h*OneHour);
   }
   return(d2);
}

//����ĳʱ��ֵ, �չ��Լ�ִ� ���� true �� false
function isDaylightSaving(d,strDS) {

   if(strDS == '') return(false);

   var m1,n1,w1,t1;
   var m2,n2,w2,t2;
   with (Math){
      m1 = floor(strDS.substr(0,2))-1;
      w1 = floor(strDS.substr(3,1));
      t1 = floor(strDS.substr(4,1));
      m2 = floor(strDS.substr(6,2))-1;
      w2 = floor(strDS.substr(9,1));
      t2 = floor(strDS.substr(10,1));
   }

   switch(strDS.substr(2,1)){
      case 'F': n1=1; break;
      case 'L': n1=-1; break;
      default : n1=0; break;
   }

   switch(strDS.substr(8,1)){
      case 'F': n2=1; break;
      case 'L': n2=-1; break;
      default : n2=0; break;
   }


   var d1, d2, re;

   if(n1==0)
      d1 = new Date(Date.UTC(d.getUTCFullYear(), m1, Math.floor(strDS.substr(2,2)),t1));
   else
      d1 = getWeekDay(d.getUTCFullYear(),m1,n1,w1,t1);

   if(n2==0)
      d2 = new Date(Date.UTC(d.getUTCFullYear(), m2, Math.floor(strDS.substr(8,2)),t2));
   else
      d2 = getWeekDay(d.getUTCFullYear(),m2,n2,w2,t2);

   if(d2>d1)
      re = (d>d1 && d<d2)? true: false;
   else
      re = (d>d1 || d<d2)? true: false;

   return(re);
}

var isDS = false;

//����ȫ��ʱ��
function getGlobeTime() {
   var d,s;
   d = new Date();

   d.setTime(d.getTime()+parseOffset(objTimeZone[0]));

   isDS=isDaylightSaving(d,objTimeZone[1]);
   if(isDS) d.setTime(d.getTime()+OneHour);
   return(showUTC(d));
}

var objTimeZone;
var objContinentMenu;
var objCountryMenu;

function tick() {
   var today;
   today = new Date();
   LocalTime.innerHTML = showLocale(today);
   GlobeTime.innerHTML = getGlobeTime();
   window.setTimeout("tick()", 1000);
}

//ָ���Զ�����ʱ��
function setTZ(a,c){
   objContinentMenu.options[a].selected=true;
   chContinent();
   objCountryMenu.options[c].selected=true;
   chCountry();
}

//�������
function chContinent() {
   var key,i;
   continent = objContinentMenu.options[objContinentMenu.selectedIndex].text;
   for (var i = objCountryMenu.options.length-1; i >= 0; i--)
      objCountryMenu[0]=null;
   for (key in timeData[continent])
      objCountryMenu.options[objCountryMenu.options.length]=new Option(key);
   objCountryMenu.options[0].selected=true;
   chCountry();
}

//�������
function chCountry() {
   var txtContinent = objContinentMenu.options[objContinentMenu.selectedIndex].text;
   var txtCountry = objCountryMenu.options[objCountryMenu.selectedIndex].text;

   objTimeZone = timeData[txtContinent][txtCountry];

   getGlobeTime();

   //��ͼλ��
   City.innerHTML = (isDS==true?"<SPAN STYLE='font-size:12pt;font-family:Wingdings; color:Red;'>R</span> ":'') + objTimeZone[2]; //�׶�
   var pos = Math.floor(objTimeZone[0].substr(0,3));
   if(pos<0) pos+=24;
   pos*=-10;
   world.style.left = pos;

}

function setCookie(name,value) {
   var today = new Date();
   var expires = new Date();
   expires.setTime(today.getTime() + 1000*60*60*24*365);
   document.cookie = name + "=" + escape(value) + "; expires=" + expires.toGMTString();
}

function getCookie(Name) {
   var search = Name + "=";
   if(document.cookie.length > 0) {
      offset = document.cookie.indexOf(search);
      if(offset != -1) {
         offset += search.length;
         end = document.cookie.indexOf(";", offset);
         if(end == -1) end = document.cookie.length;
         return unescape(document.cookie.substring(offset, end));
      }
      else return('');
   }
   else return('');
}

///////////////////////////////////////////////////////////////////////////

function initialize() {
   var key;

   //ʱ��
   map.filters.Light.Clear();
   map.filters.Light.addAmbient(255,255,255,60);
   map.filters.Light.addCone(120, 60, 80, 120, 60, 255,255,255,120,60);

   objContinentMenu=document.WorldClock.continentMenu;
   objCountryMenu=document.WorldClock.countryMenu;

   for (key in timeData)
      objContinentMenu[objContinentMenu.length]=new Option(key);


   var TZ1 = getCookie('TZ1');
   var TZ2 = getCookie('TZ2');


   if(TZ1=='') {TZ1=0; TZ2=3;}
   setTZ(TZ1,TZ2);

   tick();


   //����
   dStyle = detail.style;
   CLD.SY.selectedIndex=tY-1900;
   CLD.SM.selectedIndex=tM;
   drawCld(tY,tM);

}

function terminate() {
   setCookie("TZ1",objContinentMenu.selectedIndex);
   setCookie("TZ2",objCountryMenu.selectedIndex);
}
//-->
</SCRIPT>

<STYLE>.todyaColor {
	BACKGROUND-COLOR: aqua
}
</STYLE>

<META content="MSHTML 6.00.2800.1170" name=GENERATOR></HEAD>
<BODY onload=initialize() onunload=terminate() topmargin="0">
<!--#include file="Admin_Top.asp"-->      
<SCRIPT language=JavaScript><!--
   if(navigator.appName == "Netscape" || parseInt(navigator.appVersion) < 4)
   document.write("<h1>���������޷�ִ�д˳���</h1>�˳������� IE4 �Ժ�İ汾����ִ��!!")
//--></SCRIPT>

<DIV id=detail style="Z-INDEX: 3; FILTER: shadow(color=#333333,direction=135); WIDTH: 140px; POSITION: absolute; HEIGHT: 120px"></DIV>
<CENTER>
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" id="AutoNumber1">
<tr>
<td class="TDtop">
<TABLE border=0>
  <TBODY>
  <TR>
    <FORM name=WorldClock>
    <TD vAlign=top align=middle width=240><FONT style="FONT-SIZE: 9pt" 
      size=2>����ʱ��</FONT><BR><SPAN id=LocalTime 
      style="FONT-SIZE: 11pt; COLOR: #000080; FONT-FAMILY: Arial">0000��0��0��(��)�� 
      00:00:00</SPAN> 
      <P><SPAN id=City 
      style="FONT-SIZE: 9pt; WIDTH: 150px; FONT-FAMILY: '��ϸ����'">�й�</SPAN> 
      <BR><SPAN id=GlobeTime 
      style="FONT-SIZE: 11pt; COLOR: #000080; FONT-FAMILY: Arial">0000��0��0��(��)�� 
      00:00:00</SPAN><BR>
      <TABLE style="FONT-SIZE: 10pt; FONT-FAMILY: Wingdings">
        <TBODY>
        <TR>
          <TD align=middle>&Uacute; 
            <DIV id=map 
            style="FILTER: Light; OVERFLOW: hidden; WIDTH: 240px; HEIGHT: 120px; BACKGROUND-COLOR: mediumblue"><FONT 
            id=world 
            style="FONT-SIZE: 185px; LEFT: 0px; COLOR: green; FONT-FAMILY: Webdings; POSITION: relative; TOP: -26px">��</FONT> 
            </DIV>&Ugrave;</TD></TR></TBODY></TABLE><BR><SELECT 
      style="FONT: 9pt 'ϸ����'; WIDTH: 240px; BACKGROUND-COLOR: #e0e0ff" 
      onchange=chContinent() name=continentMenu></SELECT><BR><SELECT 
      style="FONT: 9pt 'ϸ����'; WIDTH: 240px; BACKGROUND-COLOR: #e0e0ff" 
      onchange=chCountry() name=countryMenu></SELECT></P></TD></FORM>
    <FORM name=CLD>
    <TD align=middle>
      <DIV style="Z-INDEX: -1; POSITION: absolute; TOP: 30px"><FONT id=YMBG 
      style="FONT-SIZE: 100pt; COLOR: #f0f0f0; FONT-FAMILY: 'Arial Black'">&nbsp;0000<BR>&nbsp;JUN</FONT> 
      </DIV>
      <TABLE border=0>
        <TBODY>
        <TR>
          <TD bgColor=#000080 colSpan=7><FONT style="FONT-SIZE: 9pt" 
            color=#ffffff size=2>��Ԫ<SELECT style="FONT-SIZE: 9pt" 
            onchange=changeCld() name=SY> 
              <SCRIPT language=JavaScript><!--
          for(i=1900;i<2101;i++) document.write('<option>'+i)
            //--></SCRIPT>
            </SELECT>��<SELECT style="FONT-SIZE: 9pt" onchange=changeCld() 
            name=SM> 
              <SCRIPT language=JavaScript><!--
            for(i=1;i<13;i++) document.write('<option>'+i)
            //--></SCRIPT>
            </SELECT>��</FONT> <FONT id=GZ face=�꿬�� color=#ffffff 
            size=4></FONT><BR></TD></TR>
        <TR align=middle bgColor=#e0e0e0>
          <TD width=54>��</TD>
          <TD width=54>һ</TD>
          <TD width=54>��</TD>
          <TD width=50>��</TD>
          <TD width=54>��</TD>
          <TD width=54>��</TD>
          <TD width=54>��</TD></TR>
        <SCRIPT language=JavaScript><!--
            var gNum, color1, color2;

            // ��������ɫ
            switch (conWeekend) {
            case 1:
               color1 = 'black';
               color2 = color1;
               break;
            case 2:
               color1 = 'green';
               color2 = color1;
               break;
            case 3:
               color1 = 'red';
               color2 = color1;
               break;
            default :
               color1 = 'green';
               color2 = 'red';
            }

            for(i=0;i<6;i++) {
               document.write('<tr align=center>')
               for(j=0;j<7;j++) {
                  gNum = i*7+j
                  document.write('<td id="GD' + gNum +'" onMouseOver="mOvr(' + gNum +')" onMouseOut="mOut()"><font id="SD' + gNum +'" size=5 face="Arial Black"')
                  if(j == 0) document.write(' color=red')
                  if(j == 6) {
                     if(i%2==1) document.write(' color='+color2)
                        else document.write(' color='+color1)
                  }
                  document.write(' TITLE=""> </font><br><font id="LD' + gNum + '" size=2 style="font-size:9pt"> </font></td>')
               }
               document.write('</tr>')
            }
            //--></SCRIPT>
        </TBODY></TABLE></TD>
    <TD vAlign=top align=middle width=40><BR><BR><BR>��<BR><BUTTON 
      style="FONT-SIZE: 9pt" 
      onclick="pushBtm('YD')"><B>��</B></BUTTON><BR><BUTTON 
      style="FONT-SIZE: 9pt" onclick="pushBtm('YU')"><B>��</B></BUTTON> 
      <P>��<BR><BUTTON style="FONT-SIZE: 9pt" 
      onclick="pushBtm('MD')"><B>��</B></BUTTON><BR><BUTTON 
      style="FONT-SIZE: 9pt" onclick="pushBtm('MU')"><B>��</B></BUTTON> 
      <P><BUTTON style="FONT-SIZE: 9pt" onclick="pushBtm('')">��<BR>��</BUTTON> 
      <P></P></TD></FORM></TR></TBODY></TABLE><FONT style="FONT-SIZE: 9pt" 
color=#ffffff>
<SCRIPT language=JavaScript>
	<!--
	document.write("����������: " + document.lastModified);
	//-->
	</SCRIPT>
</FONT>
</td>
</tr>
</table>
</BODY>
</HTML>
<!--#include file=Admin_Bottom.asp-->