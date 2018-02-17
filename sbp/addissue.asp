<!-- #include file=aspscripts/conn.asp -->
<%
on error resume next
stmt = "insert into jobs(shopid,shortdescription,`description`) values ('" & request.cookies("shopid") _
& "', '" & replace(pcase(request("sd")),"'","''") & "', '" & replace(pcase(request("ld")),"'","''") & "')"
con.execute(stmt)
'response.write stmt
if err.number = 0 then
	'pass back value
	response.write "success," & replace(request("ld"),","," ")
else
	response.write err.description
end if
%>
