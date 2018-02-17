<!-- #include file=aspscripts/conn.asp -->
<%
stmt = "insert into rotype (shopid,rotype) values ('" & request.cookies("shopid") & "', '" & request("rotype") & "')"
con.execute(stmt)
response.redirect "rotypes.asp"
%>
