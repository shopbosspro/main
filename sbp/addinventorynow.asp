<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
'connect to database and update information
stmt = "insert into partsinventory (PartNumber, tax, overridematrix, PartDesc, PartPrice, netonhand, PartCode, invoicenumber, bin, PartCost, PartSupplier, OnHand, ReOrderLevel, MaxOnHand, MaintainStock, PartCategory, shopid) values ('"
stmt=stmt& replace(replace(replace(ucase(request("PartNumber")),"'","''"),"%",""),"&"," and ")
stmt=stmt&"', '" & ucase(request("tax"))
stmt=stmt&"', '" & ucase(request("overridematrix"))
stmt=stmt&"', '" & replace(replace(ucase(request("PartDesc")),"'","''"),"&"," and ")
stmt=stmt&"', " & request("PartPrice")
stmt=stmt&", " & request("onhand")
stmt=stmt&", '" & ucase(request("PartCode"))
stmt=stmt&"', '" & ucase(request("invoicenumber"))
stmt=stmt&"', '" & request("bin")
stmt=stmt&"', " & request("PartCost")
stmt=stmt&", '" & ucase(replace(request("PartSupplier"),"'","''"))
stmt=stmt&"', " & request("OnHand")
stmt=stmt&", " & request("ReOrderLevel")
stmt=stmt&", " & request("MaxOnHand")
stmt=stmt&", '" & ucase(request("MaintainStock"))
stmt=stmt&"', '" & ucase(request("PartCategory"))
stmt=stmt&"', '" & request.cookies("shopid")
stmt=stmt&"')"
response.write stmt & "<BR>"

con.execute stmt

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

'update parts registry
cstmt = "select PartNumber from " & preg & " where shopid = '" & request.cookies("shopid") & "' and PartNumber = '" & replace(request("PartNumber"),"&"," and ") & "'"
set crs = con.execute(cstmt)
if crs.eof then
	stmt = "insert into " & preg & " (PartNumber, overridematrix, PartDesc, PartPrice, PartCode,bin, PartCost, PartSupplier, PartCategory, shopid, tax) values ('"
	stmt=stmt& replace(replace(replace(ucase(request("PartNumber")),"'","''"),"%",""),"&"," and ")
	stmt=stmt&"', '" & ucase(request("overridematrix"))
	stmt=stmt&"', '" & replace(replace(ucase(request("PartDesc")),"'","''"),"&"," and ")
	stmt=stmt&"', " & request("PartPrice")
	stmt=stmt&", '" & ucase(request("PartCode"))
	stmt=stmt&"', '" & request("bin")
	stmt=stmt&"', " & request("PartCost")
	stmt=stmt&", '" & ucase(replace(request("PartSupplier"),"'","''"))
	stmt=stmt&"', '" & ucase(request("PartCategory"))
	stmt=stmt&"', '" & request.cookies("shopid")
	stmt=stmt&"', '" & ucase(request("tax"))
	stmt=stmt&"')"
	response.write stmt

	con.execute(stmt)
else
	stmt = "Update " & preg & " set PartDesc = '" & replace(replace(ucase(request("PartDesc")),"'","''"),"&"," and ") & "', "
	stmt = stmt & "tax = '" & ucase(request("tax")) & "', "
	stmt = stmt & "PartPrice = " & request("PartPrice") & ", "
	stmt = stmt & "PartCode = '" & ucase(request("PartCode")) & "', "
	stmt = stmt & "PartCost = " & request("PartCost") & ", "
	stmt = stmt & "overridematrix = '" & ucase(request("overridematrix")) & "', "
	stmt = stmt & "bin = '" & request("bin") & "', "
	stmt = stmt & "PartSupplier = '" & ucase(replace(request("PartSupplier"),"'","''")) & "', "
	stmt = stmt & "PartCategory = '" & ucase(request("PartCategory")) & "' where shopid = '" & request.cookies("shopid") & "' and PartNumber = '"
	stmt = stmt & replace(replace(replace(ucase(request("PartNumber")),"'","''"),"%",""),"&"," and ") & "'"

	response.write stmt
	con.execute(stmt)
end if
response.redirect "inventory.asp"
%>
