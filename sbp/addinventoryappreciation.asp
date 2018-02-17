<!-- #include file=aspscripts/conn.asp -->
<%
response.cookies("shopid") = request("shopid")
shopid = request("shopid")

on error resume next
stmt = "insert into partsadjustment (shopid,adate,partnumber,partid,adjamount) values ('" & shopid _
& "', '" & dconv(date) & "', '" & request("partnumber") & "', " & request("partid") & ", " & request("amt") _
& ")"
con.execute(stmt)

if err.number <> 0 then
	response.write err.description
else
	response.write "success"
end if
%>
