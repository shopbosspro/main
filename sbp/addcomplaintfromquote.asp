<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=aspscripts/auth.asp -->
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Search&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
Last Nam</title>
<style type="text/css">
.style7 {
	text-align: center;
	color:white;
	font-family: Arial, Helvetica, sans-serif;
	font-size: large;
}
.style3 {
	background-image:url('newimages/wipheader.jpg');
	color:white
}

.style8 {
	font-family: Arial, Helvetica, sans-serif;
}
.style9 {
	font-family: Arial, Helvetica, sans-serif;
	color: #FFFFFF;
	background-image: url('newimages/pageheader.jpg');
}

</style>
</head>

<body>
<table width="100%" class="style3" style="height: 60px">
<tr>
<td class="style7" style="width: 100%"><strong>Complete The Quote Conversion</strong></td>
</tr>
</table>
<p class="style8"><strong>The following Labor and Parts are on Quote #<%=request.cookies("quoteid")%>.  Please assign them to a Vehicle Issue 
and Assign a Tech</strong></p>
<form action="completequoteconversion.asp" method="post">
<input type="hidden" name="roid" value="<%=request("roid")%>" />
<table cellpadding="5" cellspacing="0" style="width: 100%">
	<tr>
		<td class="style9"><strong>Vehicle Issue</strong></td>
		<td class="style9"><strong>Tech</strong></td>
		<td class="style9"><strong>Part/Labor</strong></td>
		<td class="style9"><strong>Description</strong></td>

	</tr>
	<%
	stmt = "select * from quotelabor where shopid = '" & request.cookies("shopid") & "' and quoteid = " & request.cookies("quoteid")
	set rs = con.execute(stmt)
	if not rs.eof then
		do until rs.eof
	%>
	<tr>
		<td class="style8">
		<select name="l*<%=rs("id")%>*issue">
		<%
		stmt = "select * from complaints where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
		set trs = con.execute(stmt)
		if not trs.eof then
			do until trs.eof
		%>
		<option value="<%=trs("complaintid")%>"><%=trs("complaint")%></option>
		<%
				trs.movenext
			loop
		end if
		%>
		</select></td>
		<td class="style8">
		<select size="1" name="l*<%=rs("id")%>*tech">
              
      <%
		strsql = "Select EmployeeLast, EmployeeFirst from employees where showtechlist = 'yes' and shopid = '" & request.cookies("shopid") & "' and Active = 'yes'"
		set ars = con.execute(strsql)
		do until ars.eof
	   		response.write "<option value='" & ars("EmployeeLast") & ", " & ars("EmployeeFirst")
	  		response.write "'>" & ars("EmployeeLast") & ", " & ars("EmployeeFirst") & "</option>"
	  		ars.movenext
	  	loop

      %>
        </select>		
        </td>
		<td class="style8">Labor&nbsp;</td>
		<td class="style8"><%=rs("labor")%>&nbsp;</td>
	</tr>
	<%
			rs.movenext
		loop
	end if
	stmt = "select * from quoteparts where shopid = '" & request.cookies("shopid") & "' and quoteid = " & request.cookies("quoteid")
	set rs = con.execute(stmt)
	if not rs.eof then
		do until rs.eof
	%>
	<tr>
		<td class="style8">
		<select name="p*<%=rs("id")%>*issue">
		<%
		stmt = "select * from complaints where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
		set trs = con.execute(stmt)
		if not trs.eof then
			do until trs.eof
		%>
		<option value="<%=trs("complaintid")%>"><%=trs("complaint")%></option>
		<%
				trs.movenext
			loop
		end if
		%>
		</select></td>
		<td class="style8">
		&nbsp;</td>
		<td class="style8">Part&nbsp;</td>
		<td class="style8"><%=rs("partnumber") & " - " & rs("part")%>&nbsp;</td>
	</tr>
	<%
			rs.movenext
		loop
	end if
	%>

</table>
<br />
<input name="Button1" type="submit" value="Save Changes and Continue" />
</form>
</body>

</html>
