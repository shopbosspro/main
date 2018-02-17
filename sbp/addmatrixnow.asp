<!-- #include file=aspscripts/conn.asp -->
<%
redir = "addmatrix.asp?err=yes&"
for each thing in request.form
	redir = redir & thing & "=" & request.form(thing) & "&"
	if len(request.form(thing)) = 0 then
		'errfound ="yes"
	end if
next
if errfound = "yes" then 
	response.redirect redir
else
	stmt = "insert into category (shopid,category,factor,start,end,displayorder,minfactor) values ('" & request.cookies("shopid") & "', '" & ucase(request("Category")) & "'"
	stmt = stmt & ", " & request("Factor") & ", " & request("Start") & ", " & request("End") & ",0,0)"
	response.write stmt
	con.execute(stmt)
	response.redirect "partpricematrix.asp"
end if
%>
