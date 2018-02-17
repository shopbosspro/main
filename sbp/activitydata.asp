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
<input name="Button1" type="button" onclick="location.href='activitydata.asp?excel=yes&sd=<%=request("sd")%>&ed=<%=request("ed")%>&cat=<%=request("cat")%>'" value="Export to Excel" />
<b><%=request("category")%></b>
<table style="width: 100%" cellpadding="3" cellspacing="0">
	<tr>
		<td class="auto-style1"><strong>Event Date &amp; Time</strong></td>
		<td class="auto-style1"><strong>Event</strong></td>
		<td class="auto-style1"><strong>User</strong></td>
	</tr>
	<%
	stmt = "select * from audit where category = '" & request("cat") & "' and shopid = '" & request.cookies("shopid") & "' and DATE_FORMAT(eventdatetime,'%Y-%m-%d') >= '" & dconv(request("sd")) & "' and DATE_FORMAT(eventdatetime,'%Y-%m-%d') <= '" & dconv(request("ed")) & "' order by eventdatetime"

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
	end if
	%>
</table>

</body>

</html>
