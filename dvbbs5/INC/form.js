<script language="javascript">
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){if(event.ctrlKey && window.event.keyCode==13){this.document.frmAnnounce.submit();}}
}

function DoTitle(addTitle) { 
var revisedTitle; 
var currentTitle = document.frmAnnounce.subject.value; 
revisedTitle = currentTitle+addTitle; 
document.frmAnnounce.subject.value=revisedTitle; 
document.frmAnnounce.subject.focus(); 
return; }

function isempty(str)
{
	var temp="";
	var ll=str.length;
	var index=true;
	for(var i=0;i<ll;i++)
	{
	temp=str.substring(i,i+1);
	if (temp!=" ")
	index=false
	}
	return index;
}
function check(frmAnnounce){
	if (isempty(frmAnnounce.username.value)){
	alert("��û����д�û�����");
	frmAnnounce.username.focus();
	return false;
	}
	if (isempty(frmAnnounce.passwd.value)){
	alert("��û����д���롣");
	frmAnnounce.passwd.focus();
	return false;
	}
	if (isempty(frmAnnounce.subject.value)){
	alert("��û����д�������⡣");
	frmAnnounce.subject.focus();
	return false;
	}
	return true;
}	
</script>