<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
if len(request("origshopid")) > 0 then
	shopid = request("origshopid")
else
	shopid = request.cookies("shopid")
end if

srostmt = "select VehID from vehicles order by VehID DESC limit 1"
set rs = con.execute(srostmt)
if not rs.eof then
	rs.movefirst
	newid = cdbl(rs("VehID")) + 1
else
	newid = 1
end if
set rs = nothing
stmt = "insert into vehicles (shopid,vehid,customerid,year,make,model,miles,licnumber,licstate,vin,engine,drivetype,transmission,cyl,fleetno,color,currentmileage,custom1,custom2,custom3,custom4,custom5,custom6,custom7,custom8) values ('" & shopid & "', " & newid
stmt=stmt&", " & request("custid")
stmt=stmt&", '" & request("year")
stmt=stmt&"', '" & replace(ucase(request("Make")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("Model")),"'","''")
stmt=stmt&"', '', '" & replace(ucase(request("License")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("State")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("Vin")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("Engine")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("DriveType")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("Trans")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("Cyl")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("fleetno")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("color")),"'","''")
stmt=stmt&"', '" & replace(request("currentmileage"),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom1")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom2")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom3")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom4")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom5")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom6")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom7")),"'","''")
stmt=stmt&"', '" & replace(ucase(request("custom8")),"'","''")
stmt=stmt&"')"
response.write stmt



con.execute stmt
if len(request("vin")) > 0 then
	stmt = "delete from vinscan where shopid = '" & shopid & "' and vin = '" & request("vin") & "'"
	'response.write stmt
	con.execute(stmt)
end if


if request("sp") = "prero" then
	' add the vid to the prero
	stmt = "update prero set vehid = " & newid & " where shopid = '" & request.cookies("shopid") & "' and id = " & request("id")
	con.execute(stmt)
	' redirect to prero-create-js.asp
	response.redirect "prero-create-js.asp?cid=" & request("custid") & "&vid=" & newid & "&id=" & request("id")
		
end if
if request("appt") = "y" then
	response.redirect "customer2appt.asp?customerid=" & request("custid") & "&d=" & request("d")
else
	if len(request("origshopid")) > 0 then
		response.redirect "customer2.asp?origshopid=" & request("origshopid") & "&appt=" & request("appt") & "&customerid=" & request("custid")
	else	
		response.redirect "customer2.asp?appt=" & request("appt") & "&customerid=" & request("custid")
	end if
end if
%>
