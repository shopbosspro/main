<!-- #include file=aspscripts/conn.asp -->
<%
on error resume next


stmt = "update customer set billto = '" & request("billto") & "', "
stmt = stmt & "billtoaddress = '" & request("billtoaddress") & "', "
stmt = stmt & "billtocity = '" & request("billtocity") & "', "
stmt = stmt & "billtostate = '" & request("billtostate") & "', "
stmt = stmt & "billtophone = '" & request("billtophone") & "', "
stmt = stmt & "billtozip = '" & request("billtozip") & "', "
stmt = stmt & "shippingaddress = '" & request("shippingaddress") & "', "
stmt = stmt & "shippingcity = '" & request("shippingcity") & "', "
stmt = stmt & "shippingstate = '" & request("shippingstate") & "', "
stmt = stmt & "shippingzip = '" & request("shippingzip") & "', "
stmt = stmt & "shippingphone = '" & request("shippingphone") & "', "
stmt = stmt & "shippingto = '" & request("shippingto") & "', "
stmt = stmt & "creditlimit = " & request("creditlimit") & " "
stmt = stmt & "where customerid = " & request("customerid") & " and shopid = '" & request("shopid") & "'"

con.execute(stmt)

if err.number <> 0 then
	response.write err.description
else
	response.write "success"
end if

%>