<%
function astro(birth)
	on error resume next
	dim birthday,birthmonth
	birthday=day(birth)
	birthmonth=month(birth)
	select case birthmonth
	case 1
		if birthday>=21 then
			astro="<img src=star/z11.gif alt=Ë®Æ¿×ù"&birth&">"
		else
			astro="<img src=star/z10.gif alt=Ä§ôÉ×ù"&birth&">"
		end if
	case 2
		if birthday>=20 then
			astro="<img src=star/z12.gif alt=Ë«Óã×ù"&birth&">"
		else
			astro="<img src=star/z11.gif alt=Ë®Æ¿×ù"&birth&">"
		end if
	case 3
		if birthday>=21 then
			astro="<img src=star/z1.gif alt=°×Ñò×ù"&birth&">"
		else
			astro="<img src=star/z12.gif alt=Ë«Óã×ù"&birth&">"
		end if
	case 4
		if birthday>=21 then
			astro="<img src=star/z2.gif alt=½ğÅ£×ù"&birth&">"
		else
			astro="<img src=star/z1.gif alt=°×Ñò×ù"&birth&">"
		end if
	case 5
		if birthday>=22 then
			astro="<img src=star/z3.gif alt=Ë«×Ó×ù"&birth&">"
		else
			astro="<img src=star/z2.gif alt=½ğÅ£×ù"&birth&">"
		end if
	case 6
		if birthday>=22 then
			astro="<img src=star/z4.gif alt=¾ŞĞ·×ù"&birth&">"
		else
			astro="<img src=star/z3.gif alt=Ë«×Ó×ù"&birth&">"
		end if
	case 7
		if birthday>=23 then
			astro="<img src=star/z5.gif alt=Ê¨×Ó×ù"&birth&">"
		else
			astro="<img src=star/z4.gif alt=¾ŞĞ·×ù"&birth&">"
		end if
	case 8
		if birthday>=24 then
			astro="<img src=star/z6.gif alt=´¦Å®×ù"&birth&">"
		else
			astro="<img src=star/z5.gif alt=Ê¨×Ó×ù"&birth&">"
		end if
	case 9
		if birthday>=24 then
			astro="<img src=star/z7.gif alt=Ìì³Ó×ù"&birth&">"
		else
			astro="<img src=star/z6.gif alt=´¦Å®×ù"&birth&">"
		end if
	case 10
		if birthday>=24 then
			astro="<img src=star/z8.gif alt=ÌìĞ«×ù"&birth&">"
		else
			astro="<img src=star/z7.gif alt=Ìì³Ó×ù"&birth&">"
		end if
	case 11
		if birthday>=23 then
			astro="<img src=star/z9.gif alt=ÉäÊÖ×ù"&birth&">"
		else
			astro="<img src=star/z8.gif alt=ÌìĞ«×ù"&birth&">"
		end if
	case 12
		if birthday>=22 then
			astro="<img src=star/z10.gif alt=Ä§ôÉ×ù"&birth&">"
		else
			astro="<img src=star/z9.gif alt=ÉäÊÖ×ù"&birth&">"
		end if
	case else
		astro=""
	end select
end function
%>