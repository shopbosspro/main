<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
'on error resume next
shopid = request.cookies("shopid")

stmt = "SELECT table_name FROM information_schema.tables WHERE table_schema = 'shopboss' AND table_name = 'partsregistry-" & shopid & "' LIMIT 1"
'response.write stmt
set rs = con.execute(stmt)
if not rs.eof then
	preg = "`partsregistry-" & shopid & "`"
else
	preg = "partsregistry"
end if

		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1


'connect to database and update information
if len(request("discount")) > 0 then
	disamt = request("discount")
else
	disamt = 0
end if
if len(request("net")) > 0 then
	netamt = request("net")
else
	netamt = 0
end if
if len(request("comid")) > 0 then
	comid = request("comid")
else
	comid = 0
end if
stmt = "insert into parts (partid,shopid, discount, net, complaintid, PartNumber, PartDesc, PartPrice, Quantity, ROID, Supplier, bin, Cost,"
stmt=stmt& " PartInvoiceNumber, PartCode, LineTTLPrice, LineTTLCost, `date`, PartCategory,tax, overridematrix,ponumber) values (" & newpartid & ",'" & request.cookies("shopid") & "', "
stmt=stmt& disamt & ", "
stmt=stmt& netamt & ", "
stmt=stmt& comid & ", '"
stmt=stmt& replace(replace(request("PartNumber"),"'","''"),"%","")
stmt=stmt&"', '" & replace(request("PartDesc"),"'","''")
if isnumeric(request("price")) then
	stmt = stmt & "', " & request("price")
else
	stmt = stmt & "', 0"
end if
if isnumeric(request("quantity")) then
	stmt=stmt&", " & request("Quantity")
else
	stmt=stmt&", 1"
end if
stmt=stmt&", " & request("roid")
stmt=stmt&", '" & replace(request("Supplier"),"'","''")
if request("bin") = "'" then
	thebin = ""
else
	thebin = request("bin")
end if
stmt=stmt&"', '" & replace(thebin,"'","''")
stmt=stmt&"', " & request("ReOrder")
stmt=stmt&", '" & replace(request("InvoiceNo"),"'","''")
stmt=stmt&"', '" & request("Codes")
stmt=stmt&"', " & request("TotalPrice")
stmt=stmt&", " & request("MaxReOrder")
stmt=stmt&", '" & dconv(date)
stmt=stmt&"', '" & request("Category")

if len(request("ponumber")) > 0 then
	cstmt = "select * from po where id = " & request("ponumber") & " and shopid = '" & request.cookies("shopid") & "'"
	set rs = con.execute(cstmt)
	if not rs.eof then
		ponumber = cdbl(rs("ponumber"))
	else
		ponumber = 0
	end if
else
	ponumber = 0
end if
stmt=stmt&"', '" & request("tax") & "', '" & request("overridematrix") & "', " & ponumber
stmt=stmt&")"
'response.write stmt & "<BR>"

' use the text file log in case the issue is the database server has gone away
'set fso = server.createobject("scripting.filesystemobject")
'set mfile = fso.opentextfile(server.mappath("stmtlog.txt"),8)
'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt


con.execute stmt

recordAudit "Add Part", "Added Part Number " & ucase(replace(replace(request("PartNumber"),"'","''"),"%","")) & " to RO#" & request("roid")

'update parts registry
if request("sp") = "addpartfound.asp" then
	stmt = "Update " & preg & " set PartDesc = '" & replace(request("PartDesc"),"'","''")
	stmt = stmt & "', tax = '" & request("tax")
	stmt = stmt & "', PartCode = '" & request("Codes")
	stmt = stmt & "', PartSupplier = '" & replace(request("Supplier"),"'","''")
	stmt = stmt & "', bin = '" & replace(thebin,"'","''") & "'"
	if isnumeric(request("price")) then
		stmt = stmt & ", partprice = " & request("price")
	else
		
	end if
	stmt = stmt & ", PartCategory = '" & request("Category")
	stmt = stmt & "', overridematrix = '" & request("overridematrix")
	stmt = stmt & "' where PartNumber = '" & replace(replace(request("PartNumber"),"'","''"),"%","") & "' and shopid = '" & request.cookies("shopid") & "'"
	response.write stmt & "<BR>"
	'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt
	con.execute(stmt)
end if
if request("sp") = "addpart.asp" then
	stmt = "insert into " & preg & " (shopid, PartNumber, PartDesc, PartPrice, PartCode, PartCost, tax, PartSupplier,bin, PartCategory,overridematrix) values ('" & request.cookies("shopid") & "', '"
	stmt=stmt& replace(replace(request("PartNumber"),"'","''"),"%","")
	stmt=stmt&"', '" & replace(request("PartDesc"),"'","''")
	stmt=stmt&"', " & request("Price")
	stmt=stmt&", '" & request("Codes")
	stmt=stmt&"', " & request("ReOrder")
	stmt=stmt&", '" & request("tax")
	stmt=stmt&"', '" & replace(request("Supplier"),"'","''")
	stmt=stmt&"', '" & replace(thebin,"'","''")
	stmt=stmt&"', '" & request("Category")
	stmt=stmt&"', '" & request("overridematrix")
	stmt=stmt&"')"
	'response.write stmt
	'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt
	con.execute(stmt)
end if

'increase inventory if ordered separtely
if request("usenew") = "on" then
	q = request("quantity")
	response.write "q:" & q & "<BR>"
	shopid = request.cookies("shopid")
	qtoadd = request("quantityonseparateorder")
	if qtoadd > 0 then q = qtoadd
	
	response.write "qtoadd:" & qtoadd & "<BR>"
	stmt = "update partsinventory set onhand = onhand + " & q & ", netonhand = netonhand + " & q & " where shopid = '" & shopid & "' and partnumber = '" & request("partnumber") & "'"
	response.write stmt
	'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt
	con.execute(stmt)
end if

' add to inventory if requested
if request("addtoinventory") = "Yes" then
	stmt = "insert into partsinventory (shopid,partnumber,partdesc,partprice,partcode,partcost,partsupplier,bin,onhand,netonhand,maintainstock,partcategory," _
	& "invoicenumber,tax,overridematrix) values ('" _
	& request.cookies("shopid") & "', '" _
	& replace(replace(request("PartNumber"),"'","''"),"%","") & "', '" _
	& replace(request("PartDesc"),"'","''") & "', " _
	& request("Price") & ", '" _
	& request("Codes") & "', " _
	& request("ReOrder") & ", '" _
	& replace(request("Supplier"),"'","''") & "', '" _
	& replace(thebin,"'","''") & "', " _
	& request("Quantity") & ", " _
	& request("Quantity") & ", '" _
	& "YES" & "', '" _
	& request("Category") & "', '" _
	& request("InvoiceNo") & "', '" _
	& request("tax") & "', '" & request("overridematrix") & "')"
	'response.write stmt
	'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt
	con.execute(stmt)
end if

'update inventory to reduce onhand by qty and increase allocated by qty
'check for in stock
ustmt = "select * from partsinventory where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & request("PartNumber") & "'"
set urs = con.execute(ustmt)
if not urs.eof then
	urs.movefirst
	if urs("NetOnHand") > 0 then
		newal = urs("Allocatted") + request("Quantity")
		newnet = urs("OnHand") - newal
		if newnet < 0 then newnet = 0
		upstmt = "update partsinventory set Allocatted = " & newal & ", NetOnHand = " & newnet & " where "
		upstmt = upstmt & " shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & request("PartNumber") & "'"
		'mfile.writeline request.cookies("shopid") & "|" & now & "|" & stmt
		con.execute(upstmt)
	end if
end if

if len(request("corecharge")) > 0 then
	d = replace(replace(request("corecharge"),",",""),"'","''")
	if isnumeric(d) and d > 0 then
		stmt = "insert into cores (shopid,partnumber,partdesc,corecharge,supplier,roid) values ('" & request.cookies("shopid") & "','" & replace(replace(request("PartNumber"),"'","''"),"%","")
		stmt = stmt & "','" & replace(request("PartDesc"),"'","''") & "'," & d & ",'" & replace(request("Supplier"),"'","''") & "'," & request("roid") & ")"
		con.execute(stmt)
	end if
end if

if request("addanother") = "Yes" then
	response.redirect "addpart1.asp?comid=" & request("comid") & "&roid=" & request("roid")
else
	'if request.cookies("shopid") <> "1734" then
		response.redirect "ro.asp?ap=yes&roid=" & request("roid")
	'end if
end if
'mfile.close
'set fso = nothing

%>
