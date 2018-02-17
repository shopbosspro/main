<!-- #include file=aspscripts/conn.asp -->
<%
for each x in request.form
	if len(request(x)) = 0 and x <> "cid" and x <> "vid" then
		'response.write x & "<BR>"
		response.write "<script>history.go(-1)</script>"
	end if
next
ed = request("d1") & "/" & request("d2") & "/" & request("d3")
stmt = "insert into schedule (sameday,oneday,email,sendreminder,customerid,FirstName, LastName, Year, Make, Model, Reason, SchDate, SchTime, "
stmt = stmt & "homephone, workphone, cellphone, phonetype,vehid, Hours,shopid,colorcode,origshopid) values ('"
stmt = stmt & request("sameday") & "','" & request("oneday") & "','" & request("email") & "','" & request("reminder") & "','" & request("cid") & "', '" & replace(request("firstname"),"'","''") & "', '"
stmt = stmt & replace(request("lastname"),"'","''") & "', '"
stmt = stmt & request("year") & "', '"
stmt = stmt & replace(request("make"),"'","''") & "', '"
stmt = stmt & replace(request("model"),"'","''") & "', '"
stmt = stmt & replace(request("problem"),"'","''") & "', '"
stmt = stmt & dconv(request("d1") & "/" & request("d2") & "/" & request("d3")) & "', '"
stmt = stmt & formatdatetime(request("time"),vbshorttime) & "', '" 
stmt = stmt & request("ac") & request("pf") & request("last") & "', '"
stmt = stmt & request("wac") & request("wpf") & request("wlast") & "', '"
stmt = stmt & request("cac") & request("cpf") & request("clast") & "', '"
stmt = stmt & request("phonetype") & "', '"
stmt = stmt & request("vid") & "', "
stmt = stmt & request("hours") & ", '" & request.cookies("shopid") & "', '" & request("category") & "','" & request("origshopid") & "')"

con.execute(stmt)
if request.cookies("daypilot") = "yes" then
	response.redirect "daypilot/calendar.asp?type=Resources&sd=" & request("d1") & "/" & request("d2") & "/" & request("d3")
end if

if request.cookies("cdv") = "yes" then
	redir = "calendar-dayview.asp?sd=" & ed
else
	if request("ret") = "daily" then
		redir = "schedule/daily.asp?sd=" & request("d1") & "/" & request("d2") & "/" & request("d3")
	else
		redir = "calendar.asp?ed=" & ed
	end if
end if
response.cookies("cdv").expires = now
response.redirect redir
%>