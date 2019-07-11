<SCRIPT language=JavaScript>

imgPath = new Array;
SiClickGoTo = new Array;
if (document.images)
	{
	i0 = new Image;
	i0.src = '../img/nc_works_001.gif';
	SiClickGoTo[0] = "#";
	imgPath[0] = i0.src;
	i1 = new Image;
	i1.src = '../img/nc_works_002.gif';
	SiClickGoTo[1] = "#";
	imgPath[1] = i1.src;
	i2 = new Image;
	i2.src = '../img/nc_works_003.gif';
	SiClickGoTo[2] = "#";
	imgPath[2] = i2.src;
	i3 = new Image;
	i3.src = '../img/nc_works_004.gif';
	SiClickGoTo[3] = "#"
	}
a = 0;
function ejs_img_fx(img)
	{
	if(img && img.filters && img.filters[0])
		{
		img.filters[0].apply();
		img.filters[0].play();
		}
	}

function StartAnim()
	{
	if (document.images)
		{
		document.write('<A HREF="#" onClick="ImgDest();return(false)"><IMG SRC="../img/nc_works_001.gif" BORDER=0 ALT=优良品质-精心挑选 NAME=defil style="filter:progid:DXImageTransform.Microsoft.Pixelate(MaxSquare=100,Duration=1)"></A>');
		defilimg()
		}
	else
		{
		document.write('<A HREF="#"><IMG SRC="../img/nc_works_001.gif" BORDER=0></A>')
		}
	}
function ImgDest()
	{
	document.location.href = SiClickGoTo[a-1];
	}
function defilimg()
	{
	if (a == 3)
		{
		a = 0;
		}
	if (document.images)
		{
		ejs_img_fx(document.defil)
		document.defil.src = imgPath[a];
		tempo3 = setTimeout("defilimg()",8000);
		a++;
		}
	}
</SCRIPT>
<body leftmargin="0" topmargin="0">
<table width="75%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><SCRIPT language=JavaScript>StartAnim()</SCRIPT></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
