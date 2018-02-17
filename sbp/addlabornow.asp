<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
	'verify all information
	'connect to database and update information

	set rs = nothing
	mytech = request("tech")
	'response.write "my tech = " & mytech
	if len(request("comid")) > 0 then
		comid = request("comid")
	else
		comid = 0
	end if

	'get employee hourly rate
	'create array of first and last name
	marray = split(request("tech"), ",")
	rstmt = "select hourlyrate  from employees where shopid = '" & request.cookies("shopid") & "' and employeelast = '" & trim(marray(0)) & "' and employeefirst = '" & trim(marray(1)) & "'"
	set rrs = con.execute(rstmt)
	if not rrs.eof then
		rrs.movefirst
		if isnumeric(rrs("hourlyrate")) then
			ehr = rrs("hourlyrate")
		else
			ehr = 0
		end if
	end if
	stmt = "insert into labor (shopid,roid, hourlyrate, laborhours, labor, tech, linetotal,discount,discountpercent, complaintid, techrate) values ('" & request.cookies("shopid") & "', "
	'response.write stmt
	stmt=stmt& request("roid")
	if isnumeric(request("rate")) then
		hrate = request("rate")
	else
		hrate = 0
	end if
	if isnumeric(request("hours")) then
		hhours = request("hours")
	else
		hhours = 0
	end if
	stmt=stmt&", " & hrate
	stmt=stmt&", " & hhours
	stmt=stmt&", '" & replace(ucase(request("Labor")),"'","''")
	stmt=stmt&"', '" & ucase(mytech)
	' *******  calculate the discount
	linettl = request("calcrate")
	stmt=stmt&"', " & linettl
	stmt=stmt&"," & request("discountdollars")
	stmt=stmt&"," & request("discount")
	stmt=stmt&", " & comid
	
'set fso = server.createobject("scripting.filesystemobject")
'set mfile = fso.opentextfile(server.mappath("stmtlog.txt"),8)
'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt

	stmt=stmt&", " & ehr * request("hours") & ")"
	mysplit = "no"

	'response.write "<BR>" & stmt
	con.execute stmt
	
	recordAudit "Add Labor", "Added Labor " & ucase(replace(ucase(request("Labor")),"'","''")) & " (" & hhours & " hours) to RO#" & request("roid")
	
	if request("addanother") = "Yes" then
		response.redirect "addlabor.asp?comid=" & request("comid") & "&roid=" & request("roid")
	else
		response.redirect "ro.asp?roid=" & request("roid")
	end if

%>