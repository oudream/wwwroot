<%
function astro(birth)
	on error resume next
	dim birthday,birthmonth
	birthday=day(birth)
	birthmonth=month(birth)
	select case birthmonth
	case 1
		if birthday>=21 then
			astro="<img src=star/z11.gif alt=ˮƿ��"&birth&">"
		else
			astro="<img src=star/z10.gif alt=ħ����"&birth&">"
		end if
	case 2
		if birthday>=20 then
			astro="<img src=star/z12.gif alt=˫����"&birth&">"
		else
			astro="<img src=star/z11.gif alt=ˮƿ��"&birth&">"
		end if
	case 3
		if birthday>=21 then
			astro="<img src=star/z1.gif alt=������"&birth&">"
		else
			astro="<img src=star/z12.gif alt=˫����"&birth&">"
		end if
	case 4
		if birthday>=21 then
			astro="<img src=star/z2.gif alt=��ţ��"&birth&">"
		else
			astro="<img src=star/z1.gif alt=������"&birth&">"
		end if
	case 5
		if birthday>=22 then
			astro="<img src=star/z3.gif alt=˫����"&birth&">"
		else
			astro="<img src=star/z2.gif alt=��ţ��"&birth&">"
		end if
	case 6
		if birthday>=22 then
			astro="<img src=star/z4.gif alt=��з��"&birth&">"
		else
			astro="<img src=star/z3.gif alt=˫����"&birth&">"
		end if
	case 7
		if birthday>=23 then
			astro="<img src=star/z5.gif alt=ʨ����"&birth&">"
		else
			astro="<img src=star/z4.gif alt=��з��"&birth&">"
		end if
	case 8
		if birthday>=24 then
			astro="<img src=star/z6.gif alt=��Ů��"&birth&">"
		else
			astro="<img src=star/z5.gif alt=ʨ����"&birth&">"
		end if
	case 9
		if birthday>=24 then
			astro="<img src=star/z7.gif alt=�����"&birth&">"
		else
			astro="<img src=star/z6.gif alt=��Ů��"&birth&">"
		end if
	case 10
		if birthday>=24 then
			astro="<img src=star/z8.gif alt=��Ы��"&birth&">"
		else
			astro="<img src=star/z7.gif alt=�����"&birth&">"
		end if
	case 11
		if birthday>=23 then
			astro="<img src=star/z9.gif alt=������"&birth&">"
		else
			astro="<img src=star/z8.gif alt=��Ы��"&birth&">"
		end if
	case 12
		if birthday>=22 then
			astro="<img src=star/z10.gif alt=ħ����"&birth&">"
		else
			astro="<img src=star/z9.gif alt=������"&birth&">"
		end if
	case else
		astro=""
	end select
end function
%>