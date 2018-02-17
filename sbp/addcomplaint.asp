<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->

<head>
<style type="text/css">
.style1 {
	text-align: center;
}
</style>
</head>

<%
'add complaints
'get new laborid
srostmt = "select complaintid from complaints where shopid = '" & request.cookies("shopid") & "' order by complaintid DESC limit 1"
set rs = con.execute(srostmt)
if not rs.eof then
	rs.movefirst
	newid = cdbl(rs("complaintid")) + 1
else
	newid = 1
end if
set rs = nothing
for j = 1 to 10
	c = "c" & j
	t = "t" & j
	if len(request(c)) > 0 then
		nc = replace(request(c),"'","''")
		stmt = "insert into complaints (complaintid, roid, complaint, shopid, displayorder) values (" & newid
		stmt=stmt&", " & request("roid") & ", '" & nc & "', '" & request.cookies("shopid") & "', " & j & ")"
		'response.write stmt
		con.execute(stmt)
		newid = newid + 1
		cc = cc & nc & ", "
	end if
next
stmt = "update repairorders set tagnumber = '" & request("tagnumber") & "', majorcomplaint = '" & cc & "' where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
con.execute(stmt)
redir = "ro.asp?roid=" & request("roid")
'response.redirect "ro.asp?roid=" & request("roid")
%>
<html>
<script language="javascript">
<%
if len(request.cookies("quoteid")) = 0 then
%>
<%if request.cookies("interface") <> "2" then%>window.top.<%end if%>location.href = "ro.asp?roid=<%=request("roid")%>&recalc=y"
<%
else
%>
<%if request.cookies("interface") <> "2" then%>window.top.<%end if%>location.href = "addcomplaintfromquote.asp?roid=<%=request("roid")%>"
<%
end if
%>
</script>


