<script>
function change()
{
var val = prompt('ÇëÊäÈëÄÚÈİ','');
if(val){
document.form1.name.value = val;
}
}
</script>
<form name="form1">
  <input name="name" type="text" id="name">
</form>
<p>
  <input onClick="change()" type="button" name="Button" value="Button">