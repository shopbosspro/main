<!-- #include file=aspscripts/conn.asp -->
<%
stmt = "insert into source (shopid,source) values ('" & request.cookies("shopid") & "', '" & request("source") & "')"
con.execute(stmt)
response.redirect "sources.asp"
%>
