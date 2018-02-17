<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%
srostmt = "select SubletID from sublet where shopid = '" & request.cookies("shopid") & "' order by SubletID DESC"
set rs = con.execute(srostmt)
if not rs.eof then
	rs.movefirst
	newid = cdbl(rs("SubletID")) + 1
else
	newid = 1
end if
set rs = nothing
if len(request("comid")) > 0 then
	comid = request("comid")
else
	comid = 0
end if
stmt = "insert into sublet (shopid,subletid,roid,subletdesc,subletprice,subletcost,subletinvoiceno,subletsupplier,complaintid,ponumber) values ('" & request.cookies("shopid") & "', " & newid
stmt=stmt&", " & request("roid")
stmt=stmt&", '" & replace(request("Sublet"),"'","''")
stmt=stmt&"', " & replace(request("Price"),",","")
stmt=stmt&", " & replace(request("Cost"),",","")
stmt=stmt&", '" & replace(request("Invoice"),"'","''")

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


stmt=stmt&"', '" & replace(request("Supplier"),"'","''") & "', " & comid & ", " & ponumber & ")" 
response.write stmt

con.execute stmt
response.redirect "ro.asp?roid=" & request("roid")
%>