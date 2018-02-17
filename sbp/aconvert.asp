<!-- #include file=aspscripts/conn.asp -->
<%
'convert the appt

stmt = "select * from schedule where shopid = '" & request.cookies("shopid") _
& "' and id = " & request("id")
set rs = con.execute(stmt)
custid = rs("customerid")

'********* get customerid **********
if len(custid) > 0 then
	if len(request("origshopid")) > 0 then
		response.redirect "customer2.asp?da=yes&aid=" & request("id") & "&customerid=" & custid & "&origshopid=" & request("origshopid")
	else
		response.redirect "customer2.asp?da=yes&aid=" & request("id") & "&customerid=" & custid
	end if
else
	response.redirect "addcust.asp"
end if
	
%>
