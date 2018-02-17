<!-- #include file=aspscripts/conn.asp -->
<%
stmt = "select city,state from zip where zip = '" & request("zip") & "'"
set rs = con.execute(stmt)
if not rs.eof then
	rs.movefirst
	response.write rs("city") & "*" & rs("state")
else
	response.write "none*none"
end if
%>