<%
'����Ϊ�ɸ���������
'======================================================================
'����3AS���˳�Ե����ϵͳ���ݱ���
Const db_Attach_Table = "[attach]"		'ճ��[�ϴ�]�ļ���Ϣ��X
Const db_BigClass_Table = "[BigClass]"		'�����X
Const db_Board_Table = "[board]"		'�������Ϣ��X
Const db_Dep_Table = "[dep]"			'������Ϣ��X
Const db_Link_Table = "[link]"			'����������Ϣ��X
Const db_News_Table = "[News]"			'ͼ��ϵͳ��Ϣ��X
Const db_NewsFile_Table = "[newsfile]"		'���޵���-----------------
Const db_Review_Table = "[Review]"		'������Ϣ��X
Const db_SmallClass_Table = "[smallclass]"	'С���X
Const db_Special_Table = "[Special]"		'ר����Ϣ��X
Const db_System_Table = "[System]"		'ϵͳ������Ϣ��X
Const db_Type_Table = "[type]"			'�ļ�������Ϣ��X
Const db_UploadPic_Table = "[uploadpic]"	'�����ļ�������Ϣ��
Const db_Vote_Table = "[vote]"			'ͶƱ��Ϣ��
Const db_Manage_Table = "[Admin]"		'��̨���ù����û���
'======================================================================
'����3AS���˳�Ե����ϵͳ�뱻����ϵͳ���õ��û����ݱ�
Const db_User_Table = "[Admin]"		'������ϵͳ���õ��û���
						'����6.1 Ϊ [User]
						'����7.0 Ϊ [Dv_User]
						'��������ֵΪ [ADMIN]����[FT_User]
'======================================================================
'����3AS���˳�Ե����ϵͳ�붯����̳���õ��û��ֶ���
Const db_User_ID = "ID"				'ID		�û�ID��
Const db_User_Name = "UserName"			'UserName	�û���¼��
Const db_User_Sex = "sex"			'sex		�û��Ա�[����=1 Ůʿ=0]
Const db_User_Email = "email"			'email		�û���������
Const db_User_Password = "Passwd"		'PassWD	�û���¼����
Const db_User_Question = "question"		'question	�û�����
Const db_User_Answer = "answer"			'answer	�û��ش�
Const db_User_Face = "photo"			'photo		�û���Ƭ�ļ�·��
Const db_User_RegDate = "dateandtime"		'dateandtime	�û�ע��ʱ��
Const db_User_LoginTimes = "logins"		'logins	�û���¼����
Const db_User_LastLoginTime = "lastlogin"	'lastlogin	�û����һ�ε�¼ʱ��
Const db_User_LastLoginIP = "IP"		'IP		�û������¼IP
Const db_User_birthday = "birthday"		'birthday	�û����գ�����������
Const db_User_FaceWidth = "UserWidth"		'�û�ͷ����
Const db_User_FaceHeight = "UserHeight"		'�û�ͷ��߶�

'======================================================================
'����չ������ϵͳ�����û��ֶ�����һ�㲻Ҫ�޸ģ�
Const db_ManageUser_ID = "ID"				'ID		�����û�ID��
Const db_ManageUser_Name = "UserName"			'UserName	�����û���¼��
Const db_ManageUser_Sex = "sex"			'sex		�����û��Ա�[����=1 Ůʿ=0]
Const db_ManageUser_Email = "email"			'email		�����û���������
Const db_ManageUser_Password = "Passwd"		'PassWD	�����û���¼����
Const db_ManageUser_Question = "question"		'question	�����û�����
Const db_ManageUser_Answer = "answer"			'answer	�����û��ش�
Const db_ManageUser_Face = "photo"			'photo		�����û���Ƭ�ļ�·��
Const db_ManageUser_RegDate = "dateandtime"		'dateandtime	�����û�ע��ʱ��
Const db_ManageUser_LoginTimes = "logins"		'logins	�����û���¼����
Const db_ManageUser_LastLoginTime = "lastlogin"	'lastlogin	�����û����һ�ε�¼ʱ��
Const db_ManageUser_LastLoginIP = "IP"		'IP		�����û������¼IP
Const db_ManageUser_birthday = "birthday"		'birthday	�����û����գ�����������
'----------------------------------------------------------------------
'����3AS���˳�Ե����ϵͳ�����ֶ�
Const db_User_birthyear = "birthyear"		'ԭ������
Const db_User_birthmonth = "birthmonth"		'ԭ������
Const ReadNews_CopyRight_Logo = "images/ft_logo.gif"	'�����Ķ���Ȩ��ʶ
'======================================================================
'�ۺ�������

Const Forcast_SN="Forcast2004-0001"	'��װ���к�

dim UserTableType,FileUploadPath,BbsPath
UserTableType = "FeiTium"	   		'���϶�����̳��ֵΪ"Dvbbs"
						'��������ֵΪ"FeiTium"
						'---------------------------------
FileUploadPath = "uploadfile/"			'ͼ��ϵͳ�ļ��ϴ�·��[��������]
						'---------------------------------
BbsPath = "bbs/"				'BBS����ڱ�ϵͳ��Ŀ¼,������"/"
ScrollClick = "double"				'�������Զ���������¼���ʽ,"double"Ϊ˫���Զ�������ʽ,
						'����ֵ��Ϊ��������϶���ʽ��
mouse_wheel_zoom="on"				'����ͼƬ����������,"on"Ϊ��������Ч��,"off"Ϊ������
TransShow="off"					'ҳ���л�Ч����"on"Ϊ�����л�Ч��,"off"Ϊ������
eWebEditor="0"					'�Ƿ�����eWebEditor���������ԭ�༭����ֵΪ"1"ʱ���ã�����ֵΪ�����á�
IsDebug="0"							'�Ƿ�Ϊ����ģʽ��ֵΪ"1"ʱ����,����ֵΪ������.
'=======================================================================
'����Ϊ�ɸ���������

'=======================================================================
'������Ϣһ��������������
'=======================================================================
dim db_Sex_select , db_Birthday_Select , FaceUploadPath , db_Set_Table , db_Forum_UserNum , db_Form_lastUser

if UserTableType = "FeiTium" then
	'------------����ѡ��------------		========================================
	db_Sex_Select = "FeiTium"			'�Ա��ֶ�����ѡ���ҪΪ����ԭͼ��ϵͳ������ϵͳ
							'��������ֵΪ "FeiTium",Ϊ�ı���
							'����3AS���˳�Ե����ϵͳԭ�û��Ա��ֶ�����Ϊ�ı���
							'---------------------------------------
	db_Birthday_Select = "FeiTium"			'�����ֶμ���ѡ��
							'��������ֵΪ "FeiTium",Ϊһ�����������ڵ��ı����ֶ�,
							'����birthyear��birthmonth�ֶ���ϣ��γ�������������
							'=======================================
	FaceUploadPath = "uploadfile/face/"		'ͷ���ϴ�Ŀ¼,������"/"
							'��������ֵΪ "uploadfile/face/"
							'����3AS���˳�Ե�����ͷ��·��
							'---------------------------------
else
	if UserTableType = "Dvbbs" then
		'����3AS���˳�Ե����ϵͳ V0.45Finish���붯����̳�������ݿ��BBS���ò�����
		db_Set_Table = "[Dv_Setup]"
							'------------DVBBS���ò�������-------------
		db_Forum_UserNum = "Forum_UserNum"	'��̳���û���
		db_Forum_lastUser = "Forum_lastUser"	'��̳���ñ�����ע���û�

		'------------����ѡ��------------	========================================
		db_Sex_Select = "Number"		'�Ա��ֶ�����ѡ���ҪΪ����ԭͼ��ϵͳ������ϵͳ
							'����ϵͳ��ֵΪ "Number",Ϊ������
							'�����û��Ա��ֶ�����Ϊ��ֵ��
							'---------------------------------------
		db_Birthday_Select = "Text"	'�����ֶμ���ѡ��
							'����ϵͳ��ֵΪ "Text",Ϊ���������յ�һ���ı����ֶ�
							'=======================================
		FaceUploadPath = "UploadFace/"		'ͷ���ϴ�Ŀ¼,������"/"
							'���϶�����Ϊ"UploadFace/",�����BBS��װĿ¼
							'---------------------------------
		FaceDefault = "Images/userface/image1.gif"
							'�û�ע�ᣨ����ͼ�ĲࣩĬ��ͷ���ļ�������� BbsPath ·����
	end if
end if

''============================================
''����Ϊԭ����3AS���˳�Ե����ϵͳ V0.45������
''============================================
Set rs9 = Server.CreateObject("ADODB.Recordset")
sql9 ="SELECT * From "& db_System_Table &""
RS9.open sql9,Conn,3,3
xpurl = rs9("xpurl")
email = rs9("email")
copyright = rs9("copyright")
version = rs9("version")
ver = rs9("ver")
logo = rs9("logo")
logourl = rs9("logourl")
banner = rs9("banner")
bannerurl = rs9("bannerurl")
bgcolor = rs9("bgcolor")
bgimg = rs9("bgimg")
zs = rs9("zs")
zs1 = rs9("zs1")
zs2 = rs9("zs2")
reg = rs9("reg")
inputmodpwd = rs9("inputmodpwd")
upload = rs9("upload")
fabiao = rs9("fabiao")
fabiaomod = rs9("fabiaomod")
fabiaocheck = rs9("fabiaocheck")
gd1 = rs9("gd1")
gd2 = rs9("gd2")
checkmod = rs9("checkmod")
checkdel = rs9("checkdel")
smallgl = rs9("smallgl")
specialgl = rs9("specialgl")
system = rs9("system")
css = rs9("css")
init = rs9("init")
delreview = rs9("delreview")
shsmallgl = rs9("shsmallgl")
shspecialgl = rs9("shspecialgl")
shdelreview = rs9("shdelreview")
reviewable = rs9("reviewable")
modnewsable = rs9("modnewsable")
checkshow = rs9("checkshow")
showuser = rs9("showuser")
search = rs9("search")
showsearch = rs9("showsearch")
showspecial = rs9("showspecial")
showdata = rs9("showdata")
showauthor = rs9("showauthor")
showclub = rs9("showclub")
showlink = rs9("showlink")
showlinkmap = rs9("showlinkmap")
showvote = rs9("showvote")
showcount = rs9("showcount")
linkreg = rs9("linkreg")
linkpass = rs9("linkpass")
linkmana = rs9("linkmana")
votemana = rs9("votemana")
powermana = rs9("powermana")
showip = rs9("showip")
reviewcheck = rs9("reviewcheck")
reviewcheckshow = rs9("reviewcheckshow")
showyear = rs9("showyear")
showtime = rs9("showtime")
showclick = rs9("showclick")
shownum = rs9("shownum")
freetime = rs9("freetime")
uselevel = rs9("uselevel")
uploadtype = rs9("uploadtype")

top_news = rs9("top_news")
bigclassshownum = rs9("bigclassshownum")
top_sp = rs9("top_sp")
top_txt = rs9("top_txt")
top_img = rs9("top_img")
topuser = rs9("topuser")
linkshownum = rs9("linkshownum")
reviewnum = rs9("reviewnum")

topbg = rs9("topbg")
topfont = rs9("topfont")
t_bg = rs9("t_bg")
b_bg = rs9("b_bg")
T_m_bg = rs9("T_m_bg")
l_BG = rs9("l_BG")
l_top = rs9("l_top")
l_main = rs9("l_main")
m_BG = rs9("m_BG")
m_main = rs9("m_main")
m_top = rs9("m_top")
r_top = rs9("r_top")
r_BG = rs9("R_BG")
r_main = rs9("r_main")
border = rs9("border")

gg = rs9("gg")
gg1 = rs9("gg1")
jjgn = rs9("name")
PicNum = rs9("picnum")			'����ͼƬJS���ŵ�����
newsNum = rs9("newsnum")		'����һ��JS���ŵ�����
rs9.close
set rs9 = nothing
%>
