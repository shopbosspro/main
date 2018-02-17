<!-- #include file=aspscripts/conn.asp -->
<%
'add xref number
stmt = "insert into xref (shopid,partnumber,xref) values ('" & request.cookies("shopid") & "', '" & request("PartNumber") & "', '" & request("XREF") & "')"
con.execute(stmt)
response.redirect "editinventory.asp?PartNumber=" & request("PartNumber")
%>
