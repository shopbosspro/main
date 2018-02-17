<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
'verify all information
redirstmt = "addcust.asp?err=yes"
if len(request("LastName")) = 0 then
	for each thing in request.form
		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
	next
	response.redirect redirstmt
end if
'if len(request("FirstName")) = 0 then
'	for each thing in request.form
'		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
'	next
'	response.redirect redirstmt
'end if
if len(request("Address")) = 0 then
	for each thing in request.form
		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
	next
	response.redirect redirstmt
end if
if len(request("City")) = 0 then
	for each thing in request.form
		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
	next
	response.redirect redirstmt
end if
if len(request("State")) = 0 then
	for each thing in request.form
		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
	next
	response.redirect redirstmt
end if
if len(request("Zip")) = 0 then
	for each thing in request.form
		redirstmt = redirstmt & "&" & thing & "=" & request.form(thing)
	next
	response.redirect redirstmt
end if
'connect to database and update information
srostmt = "select CustomerID from customer where shopid = '" & request.cookies("shopid") & "' order by CustomerID DESC"
set rs = con.execute(srostmt)
if not rs.eof then
	rs.movefirst
	newcid = cdbl(rs("CustomerID")) + 1
else
	newcid = 1
end if
stmt = "insert into customer (shopid, CustomerID, LastName, FirstName, Address, cellprovider, City"
stmt=stmt&", State, Zip, HomePhone, WorkPhone, CellPhone, Fax, contact, EMail,spousename,spousecell,spousework,extension,customertype,userdefined1,userdefined2,userdefined3) values ('" & request.cookies("shopid") & "', "
stmt=stmt& newcid
stmt=stmt& ", '" & ucase(replace(request("LastName"),"'","''"))
stmt=stmt& "', '" & ucase(replace(request("FirstName"),"'","''"))
stmt=stmt& "', '" & ucase(replace(request("Address"),"'","''"))
stmt=stmt& "', '" & replace(request("cellprovider"),"'","''")
stmt=stmt& "', '" & replace(ucase(request("City")),"'","''")
stmt=stmt& "', '" & replace(ucase(request("State")),"'","''")
stmt=stmt& "', '" & replace(request("Zip"),"'","''")
stmt=stmt& "', '" & replace(request("HomePhone"),"'","''") & replace(request("HomePhone1"),"'","''") & replace(request("HomePhone2"),"'","''")
stmt=stmt& "', '" & replace(request("WorkPhone"),"'","''") & replace(request("WorkPhone1"),"'","''") & replace(request("WorkPhone2"),"'","''")
stmt=stmt& "', '" & replace(request("CellPhone"),"'","''") & replace(request("CellPhone1"),"'","''") & replace(request("CellPhone2"),"'","''")
stmt=stmt& "', '" & replace(request("Fax"),"'","''") & replace(request("Fax1"),"'","''") & replace(request("Fax2"),"'","''")
stmt=stmt& "', '" & replace(request("contact"),"'","''")
stmt=stmt& "', '" & replace(ucase(request("EMail")),"'","''")
stmt=stmt& "', '" & replace(ucase(request("spousename")),"'","''")
stmt=stmt& "', '" & replace(request("spousecell"),"'","''")
stmt=stmt& "', '" & replace(request("spousework"),"'","''") & "', '" & replace(request("workphone3"),"'","''") & "','" & replace(ucase(request("customerpaytype")),"'","''")
stmt=stmt& "', '" & replace(request("userdefined1"),"'","''")
stmt=stmt& "', '" & replace(request("userdefined2"),"'","''")
stmt=stmt& "', '" & replace(request("userdefined3"),"'","''") & "')"


'response.write stmt
con.execute(stmt)
response.redirect "addveh.asp?vin=" & request("vin") & "&d=" & request("d") & "&appt=" & request("appt") & "&year=" & request("vyear") & "&make=" & request("vmake") & "&model=" & request("vmodel") & "&reason=" & server.urlencode(request("reason")) & "&custid=" & newcid & "&sp=" & request("sp")
%>
