<!-- #include file=aspscripts/conn.asp -->
<%
stmt = "insert into jobdesc (shopid,jobdesc) values ('" & request.cookies("shopid") & "', '" & request("JobDesc") & "')"
con.execute(stmt)
response.redirect "jobdescs.asp"
%>
