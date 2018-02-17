<!-- #include file=aspscripts/conn.asp -->
<%
'****** change employee status, active to inactive or vice-versa **********
if ucase(request("curstat")) = "NO" then
	newstat = "YES"
else
	newstat = "NO"
end if
stmt = "update employees set Active = '" & newstat & "' where shopid = '" _
& request.cookies("shopid") & "' and id = '" & request("empid") & "'"
response.write stmt
con.execute(stmt)

response.redirect "employees.asp"
%>
