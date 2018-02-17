<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
shopid = request.cookies("shopid")
stmt = "SELECT table_name FROM information_schema.tables WHERE table_schema = 'shopboss' AND table_name = 'partsregistry-" & shopid & "' LIMIT 1"
'response.write stmt & "<BR>"
set rs = con.execute(stmt)
'response.write rs.eof & "<BR>"
if not rs.eof then
	preg = "`partsregistry-" & shopid & "`"
else
	preg = "partsregistry"
end if

'verify all information
for each thing in request.form
	if len(request.form(thing)) < 1 then
		redirstmt = request("s") & "?err=yes" & "&pid=" & request("pid")
		for each item in request.form
			redirstmt = redirstmt & "&" & item & "=" & request.form(item)
		next
	end if
next
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
if isnumeric(request("discount")) then
	disc = request("discount")
else
	disc = 0
end if
stmt = "insert into psdetail (invoicenumber,shopid,psid,pnumber,`desc`,price,qty,partid,discount,ext,supplier,partcode,cost,tax,category) values ('"
stmt = stmt & request("invoicenumber") & "', '"
stmt = stmt & request.cookies("shopid") & "', "
stmt = stmt & request("psid")
stmt = stmt & ", '" & request("partnumber")
stmt = stmt & "', '" & replace(request("partdesc"),"'","''")
stmt = stmt & "', " & request("price")
stmt = stmt & ", " & request("quantity")
stmt = stmt & ", " & request("partid")
stmt = stmt & ", " & disc
stmt = stmt & ", " & request("totalprice")
stmt = stmt & ", '" & replace(request("supplier"),"'","''")
stmt = stmt & "', '" & request("codes")
stmt = stmt & "', " & request("reorder")
stmt = stmt & ", '" & request("tax")
stmt = stmt & "', '" & request("category")
stmt = stmt & "')"
response.write stmt & "<BR>"
con.execute stmt
'update parts registry
if request("sp") = "pspartfound.asp" then
	stmt = "Update " & preg & " set PartDesc = '" & replace(request("PartDesc"),"'","''")
	stmt = stmt & "', PartPrice = " & request("Net")
	stmt = stmt & ", PartCode = '" & request("Codes")
	stmt = stmt & "', PartCost = " & request("ReOrder")
	stmt = stmt & ", PartSupplier = '" & replace(request("supplier"),"'","''")
	stmt = stmt & "', tax = '" & request("tax")
	stmt = stmt & "', PartCategory = '" & request("Category")
	stmt = stmt & "' where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & request("PartNumber") & "'"
	response.write stmt & "<BR>"

	con.execute(stmt)
end if
if request("sp") = "psaddpart.asp" then
	stmt = "insert into " & preg & " (shopid, PartNumber, PartDesc, PartPrice, PartCode, PartCost, tax, PartSupplier, PartCategory) values ('"
	stmt=stmt & request.cookies("shopid") & "', '"
	stmt=stmt& request("PartNumber")
	stmt=stmt&"', '" & replace(request("PartDesc"),"'","''")
	stmt=stmt&"', " & request("Price")
	stmt=stmt&", '" & request("Codes")
	stmt=stmt&"', " & request("ReOrder")
	stmt=stmt&", '" & request("tax")
	stmt=stmt&"', '" & replace(request("supplier"),"'","''")
	stmt=stmt&"', '" & request("Category")
	stmt=stmt&"')"
	response.write stmt

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
		upstmt = "update partsinventory set Allocatted = " & newal & ", NetOnHand = " & newnet & " where shopid = '" & request.cookies("shopid") & "' and "
		upstmt = upstmt & " PartNumber = '" & request("PartNumber") & "'"
		con.execute(upstmt)
	end if
end if
if request("addanother") = "Yes" then
	response.redirect "pspartsfind.asp?psid=" & request("psid")
else
	response.redirect "ps.asp?psid=" & request("psid")
end if
%>
