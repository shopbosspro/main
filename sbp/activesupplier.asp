<!-- #include file=aspscripts/conn.asp -->
<%

'*********  change status of supplier from active to inactive and vice versa ***********
if ucase(request("curstat")) = "NO" then
	newstat = "YES"
else
	newstat = "NO"
end if
stmt = "update supplier set Active = '" & newstat & "' where shopid = '" & request.cookies("shopid") & "' and id = " & request("suppid")

con.execute(stmt)
response.redirect "suppliers.asp"
%>
