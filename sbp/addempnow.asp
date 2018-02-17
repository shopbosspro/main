<!-- #include file=aspscripts/conn.asp -->
<%

' create employee id from lastname and firstname without vowels
tempid = request("employeeid")

vowels = "A,E,I,O,U"
tar = split(vowels,",")
tempid = replace(ucase(request("EmployeeLast")),"'","") & replace(ucase(request("EmployeeFirst")),"'","")
for j = lbound(tar) to ubound(tar)
	if instr(tempid,tar(j)) then
		tempid = replace(tempid,tar(j),"")
	end if
next
tempid = replace(tempid," ","")


if len(request("pin")) = 0 then
	pin = "null"
else
	pin = request("pin")
end if

fldlist = "`shopid`,`EmployeeID`,`EmployeeLast`,`EmployeeFirst`,`JobDesc`,`DateHired`,`DefaultWriter`," 
fldlist = fldlist & "`hourlyrate`,`pin`,`tooltip`,`CompanyAccess`,`EmployeeAccess`,`ReportAccess`,`CreateRO`,`CreateCT`,`EditSupplier`,`InventoryLookup`,`EditInventory`,"
fldlist = fldlist & "`ReOpenRO`,`ChangeUserSecurity`,`ChangePartMatrix`,`ChangePartCodes`,`ChangeJobDescription`,`ChangeSources`,`ChangeRepairOrderTypes`,`logintosbp`,`accounting`,`mode`,"
fldlist = fldlist & "`password`,`sendupdates`,`deletepaymentsreceived`,`mechanicnumber`,`paytype`,`showtechlist`,`candelete`,`changeshopnotice`,`deletecustomer`"

stmt = "Insert into employees (" & fldlist & ") values ('"
stmt = stmt & request.cookies("shopid") & "','" 
stmt = stmt & tempid
stmt = stmt & "', '" &  replace(ucase(request("EmployeeLast")),"'","")
stmt = stmt & "', '" & replace(ucase(request("EmployeeFirst")),"'","")
stmt = stmt & "', '" & ucase(request("Position"))
stmt = stmt & "', '" & dconv(request("sd"))
stmt = stmt & "', '" & ucase(request("DefaultWriter"))
stmt = stmt & "', " & request("hourlyrate")
stmt = stmt & ", '" & pin
stmt = stmt & "', '" & request("tooltip")
stmt = stmt & "', '" & request("companyaccess")
stmt = stmt & "', '" & request("employeeaccess")
stmt = stmt & "', '" & request("reportaccess")
stmt = stmt & "', '" & request("createro")
stmt = stmt & "', '" & request("CreateCT")
stmt = stmt & "', '" & request("editsupplier")
stmt = stmt & "', '" & request("inventorylookup")
stmt = stmt & "', '" & request("editinventory")
stmt = stmt & "', '" & request("ReOpenRO")
stmt = stmt & "', '" & request("ChangeUserSecurity")
stmt = stmt & "', '" & request("ChangePartMatrix")
stmt = stmt & "', '" & request("ChangePartCodes")
stmt = stmt & "', '" & request("ChangeJobDescription")
stmt = stmt & "', '" & request("ChangeSources")
stmt = stmt & "', '" & request("ChangeRepairOrderTypes")
stmt = stmt & "', '" & request("loginaccess")
stmt = stmt & "', '" & request("accounting")
stmt = stmt & "', '" & request("mode")
stmt = stmt & "', '" & request("password")
stmt = stmt & "', '" & request("sendupdates")
stmt = stmt & "', '" & request("deletepaymentsreceived")
stmt = stmt & "', '" & request("mechanicnumber")
stmt = stmt & "', '" & request("paytype")
stmt = stmt & "', '" & request("showtechlist")
stmt = stmt & "', '" & request("candelete")
stmt = stmt & "', '" & request("changeshopnotice")
stmt = stmt & "', '" & request("deletecustomer")
stmt = stmt & "')"
'response.write stmt & "<BR>"
con.execute(stmt)


response.redirect "employees.asp"
%>