<!-- #include file=aspscripts/connwoshopid.asp -->
<%
stmt = "update schedule set display = 'no' where shopid = '" & request("shopid") & "' and id = " & request("id")
con.execute(stmt)
response.write "success"
%>
