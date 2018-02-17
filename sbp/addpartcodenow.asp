<!-- #include file=aspscripts/conn.asp -->
<%
stmt = "insert into codes (shopid,`codes`) values ('" & request.cookies("shopid") & "', '" & request("Code") & "')"
con.execute(stmt)
response.redirect "partcodes.asp"
%>
