<!-- #include file=aspscripts/conn.asp -->

<%
stmt = "insert into reminders (shopid,customerid,vehid,reminderservice,reminderdate) values ('" & request.cookies("shopid") & "', " _
& request("customerid") & ", " & request("vehid") & ", '" & request("reminder") & "', '" & dconv(request("reminderdate")) & "')"
con.execute(stmt)
response.redirect "ro.asp?roid=" & request("roid")
%>
