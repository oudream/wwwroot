<!--#include file="../conn/conn.asp" -->

<%
delfield=request("selfield")
delmonth=request("selmonth")
delday=request("selday")
delyear=request("selyear")
delquantum=request("selquantum")

if delfield="" then delfield=1
if delmonth="" then delmonth=1
if delday="" then delday=1
if delyear="" then delyear=2003
if delquantum="" then delquantum=1

dim RSEVENTS
Set RSEVENTS = Server.CreateObject("ADODB.RecordSet")
RSEVENTS.Open "MAIN", Conn, 2, 2
RSEVENTS.addnew

RSEVENTS("field") = cstr(request("selfield"))
RSEVENTS("quantum") = cstr(request("selquantum"))
RSEVENTS("year") = cstr(request("selyear"))
RSEVENTS("month")= cstr(request("selmonth"))
RSEVENTS("day")= cstr(request("selday"))

eventDate=cdate(cstr(delyear)+"-"+cstr(delmonth)+"-"+cstr(delday))
RSEVENTS("date") = eventDate

RSEVENTS("inserttime")=date
RSEVENTS.update

RSEVENTS.close
set RSEVENTS = nothing
%>

