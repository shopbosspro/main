<%
'test for cross site scripting
mylist = "<,>,(,),],[,},{,--,=,|"
myar = split(mylist,",")
for each x in request.querystring
	for j = lbound(myar) to ubound(myar)
		if instr(x,myar(j)) then response.redirect "bad.asp"
		
		if instr(request.querystring(x),myar(j)) then
			'response.write request(x) & " : " & myar(j) & "<BR>"
			'redirect to home page
			response.redirect "default.asp"
		end if
	next
next

for each x in request.form
	for j = lbound(myar) to ubound(myar)
		if instr(x,myar(j)) then response.redirect "bad.asp"
		if instr(request.form(x),myar(j)) then
			'redirect to home page
			'response.write x & ":" & myar(j)
			response.redirect "default.asp"
		end if
	next
next

function dconv(d)

	dar = split(d,"/")
	dconv = dar(2) & "-" & dar(0) & "-" & dar(1)


end function

on error resume next
' connect to primary DB Server
set con = server.createobject("adodb.connection")
con.open "DRIVER={MySQL ODBC 5.2w Driver};Server=mysql-lb.shopbosspro.com;port=3306;Database=shopboss;UID=rootsbp;pwd=ashley92;charset=ucs2;"



%>
