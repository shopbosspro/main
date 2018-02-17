<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
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
stmt = "insert into complaints (complaintid, roid, complaint, techreport,shopid) values (" & newid
stmt = stmt & ", " & request("roid") & ", '" & replace(request("complaint"), "'", "`") & "', '" & replace(request("techreport"), "'", "`") & "', '"
stmt = stmt & request.cookies("shopid") & "')"
'response.write stmt
con.execute(stmt)
redir = "ro.asp?roid=" & request("roid")
%>
<html>
<script language="javascript">
window.top.location.href = "complaintaddupdate.asp?roid=<%=request("roid")%>"
</script>