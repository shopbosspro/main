<%
if request("excel") = "yes" then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=audit.xls"
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/conn.asp -->
<head>
<meta content="en-us" http-equiv="Content-Language" />
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Untitled 1</title>
<style type="text/css">
body{
	font-family:Arial, Helvetica, sans-serif
}
.auto-style1 {
	background-color: #C0C0C0;
}
</style>
</head>

<body>
<input name="Button1" type="button" onclick="location.href='activitydataro.asp?excel=yes&roid=<%=request("roid")%>'" value="Export to Excel" />
<b><br><%=request("category")%></b>
<strong>LOGS:
</strong>
<table style="width: 100%" cellpadding="3" cellspacing="0">
	<tr>
		<td class="auto-style1"><strong>Event Date &amp; Time</strong></td>
		<td class="auto-style1"><strong>Event</strong></td>
		<td class="auto-style1"><strong>User</strong></td>
	</tr>
	<%
	stmt = "select * from audit where shopid = '" & request.cookies("shopid") & "' and event like '%RO#" & request("roid") & "%' order by eventdatetime"
	set rs = con.execute(stmt)
	if not rs.eof then
		do until rs.eof
	%>
	<tr>
		<td><%=rs("eventdatetime")%>&nbsp;</td>
		<td><%=rs("event")%>&nbsp;</td>
		<td><%=rs("useraccount")%>&nbsp;</td>
	</tr>
	<%
			rs.movenext
		loop
	else
		response.write "No results generally means that there is no data in the log.  Audit log began on 4/1/2015"
	end if
	%>
</table>
<strong><br/>COMMUNICATIONS
<br/></strong>
<table cellpadding="3" cellspacing="0" style="width: 100%">
	<tr>
		<td class="auto-style1" style="width: 15%"><strong>Date / Time</strong></td>
		<td class="auto-style1" style="width: 15%"><strong>Logged By</strong></td>
		<td class="auto-style1" style="width: 70%"><strong>Communication</strong></td>
	</tr>
	<%
	stmt = "select * from repairordercommhistory where roid = " & request("roid") & " and shopid = '" & request.cookies("shopid") & "' order by datetime desc"
	set rs = con.execute(stmt)
	if not rs.eof then
		do until rs.eof
	%>
	<tr>
		<td style="width: 15%" valign="top"><%=rs("datetime")%>&nbsp;</td>
		<td style="width: 15%" valign="top"><%=rs("by")%>&nbsp;</td>
		<td style="width: 70%" valign="top"><%=rs("comm")%>&nbsp;</td>
	</tr>
	<%
			rs.movenext
		loop
	else
	%>
	<tr>
		<td colspan="3">No logged communication for this RO&nbsp;</td>
	</tr>
	<%
	end if
	%>
</table>
</body>

</html>
