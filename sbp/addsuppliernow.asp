<!-- #include file=aspscripts/conn.asp -->
<%
'get newid
set rs = con.execute("Select SupplierID from supplier where shopid = '" & request.cookies("shopid") & "' order by SupplierID DESC")
if not rs.eof then
	rs.movefirst
	newid = cdbl(rs("SupplierID")) + 1
else
	newid = 1
end if
stmt = "insert into supplier (shopid,SupplierID, SupplierName, SupplierAddress, SupplierCity, SupplierState, SupplierZip, SupplierPhone, SupplierFax, SupplierContact, Active,accountnumber,terms,taxexempt,notes)"
stmt = stmt & " values ('" & request.cookies("shopid") & "', " & newid & ", '" & ucase(replace(request("SupplierName"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("SupplierAddress"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("SupplierCity"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("SupplierState"),"'","''"))
stmt = stmt & "', '" & replace(request("SupplierZip"),"'","''")
stmt = stmt & "', '" & request("sac")&request("spre")&request("sl4")
stmt = stmt & "', '" & request("sfac")&request("sfpre")&request("sfl4")
stmt = stmt & "', '" & ucase(replace(request("SupplierContact"),"'","''"))
stmt = stmt & "', '" & replace(request("Active"),"'","''")
stmt = stmt & "', '" & ucase(replace(request("accountnumber"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("terms"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("taxexempt"),"'","''"))
stmt = stmt & "', '" & ucase(replace(request("notes"),"'","''"))
stmt = stmt & "')"
con.execute(stmt)


max=99999
min=11111
Randomize
r = Int((max-min+1)*Rnd+min)




if request("sp") = "addpart.asp" then
	response.redirect request("sp") & "?r=" & r & "&roid=" & request("roid") & "&comid=" & request("comid")
end if
if request("sp") = "suppliers.asp" then
	response.redirect "suppliers.asp?r=" & r
end if
%>
