<!-- #include file=aspscripts/conn.asp -->
<%
on error resume next
if len(request("totalrec")) = 0 then
	totalrec = 0
else
	if isnumeric(request("totalrec")) then
		totalrec = request("totalrec")
	else
		totalrec = 0
	end if
end if
stmt = "insert into recommend (shopid,roid,`desc`,totalrec) values ('" & request.cookies("shopid") & "', " & request("roid") _
& ", '" & replace(request("desc"),"'","''") & "', " & totalrec & ")"
con.execute(stmt)

if err.number > 0 then
	response.write "There was an error processing your request"
else
	response.write "success"
end if
%>
