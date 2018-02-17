<!-- #include file=conn.asp -->
<%
if len(request("c")) > 0 then
	stmt = "select distinct state from zip where country = '" & request("c") & "' order by state"
	set rs = con.execute(stmt)
	response.write "<select onchange='getStates(""this.value"")' name='state' style='height: 30px; width: 103px;font-size:18px'>"
	do until rs.eof
		response.write "<option value='" & rs("state") & "'>" & rs("state") & "</option>" & chr(10)
		rs.movenext
	loop
	response.write "</select>"
end if
%>
