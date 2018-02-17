<!-- #include file=aspscripts/connwoshopid.asp --> 
<%
' get joined shops if exists
shopid = request("shopid")

'test for linked shops
stmt = "select shopid,joinedshopid from joinedshops where shopid = '" & shopid & "' or joinedshopid = '" & shopid & "'"
set rs = con.execute(stmt)
if rs.eof then
	joinedshops = false
	jslist = "none"
else
	joinedshops = true
	do until rs.eof
		jslist = jslist & "'" & rs("shopid") & "',"
		rs.movenext
	loop

end if

if instr(jslist,",") then
	jslist = left(jslist,len(jslist)-1)
end if

if len(request("fn")) > 0 then
	fn = replace(request("ln"),"'","''")
else 
	fn = ""
end if
if len(request("ln")) > 0 then
	ln = replace(request("fn"),"'","''")
else
	ln = ""
end if
if len(request("a")) > 0 then
	a = replace(request("a"),"'","''")
else
	a = ""
end if

if joinedshops = false then
	
	stmt = "select * from customer where shopid = '" & shopid & "' and lastname = '" & ln & "' and firstname = '" & fn & "' and address = '" & a & "'"
	'response.write stmt
	set rs = con.execute(stmt)
	if rs.eof then
		response.write "nodup"
	else
		response.write "dup"
	end if
	
else

	stmt = "select * from customer where lastname = '" & ln & "' and firstname = '" & fn & "' and address = '" & a & "' and shopid in (" & jslist & ")"
	'response.write stmt
	set rs = con.execute(stmt)
	if rs.eof then
		response.write "nodup"
	else
		response.write "dup"
	end if

end if



%>