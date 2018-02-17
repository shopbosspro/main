<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=aspscripts/auth.asp -->
<%
if len(request("hourlyrate")) > 0 then
	'add all the labor to the labor table for the ro
	
	'check for multiple
	if instr(request("cannedjob"),",") then
	
		' its a multiselect
		lar = split(request("cannedjob"),",")
		for j = lbound(lar) to ubound(lar)
			cannedid = trim(lar(j))
			addCannedJob(cannedid)
		next
	else
		addCannedJob(request("cannedjob"))
	end if
	response.redirect "ro.asp?roid=" & request("roid")
end if

function addCannedJob(id)	

	' **** first check for flat pricing on the canned job
	stmt = "select * from cannedjobs where id = " & id
	set trs = con.execute(stmt)
	if trs("flatprice") = 0 then

		stmt = "select * from cannedlabor where shopid = '" & request.cookies("shopid") & "' and cannedjobsid = " & id & " order by id asc"
		set rs = con.execute(stmt)
		if not rs.eof then
			do until rs.eof
				nstmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,complaintid) values " _
				& "('" & request.cookies("shopid") & "', " _
				& request("roid") & ", " _
				& request("hourlyrate") & ", " _
				& rs("laborhours") & ", '" _
				& rs("labor") & "', '" _
				& request("tech") & "', " _
				& request("hourlyrate") * rs("laborhours") & ", " _
				& request("comid") & ")"
				'response.write nstmt & "<BR>"
				con.execute(nstmt)
				recordAudit "Add Labor", "Added Labor " & replace(ucase(rs("Labor")),"'","''") & " (" & rs("laborhours") & " hours) to RO#" & request("roid")
				rs.movenext
			loop
		end if
		
		
		'now get the parts
		stmt = "select * from cannedparts where shopid = '" & request.cookies("shopid") & "' and cannedjobsid = " & id & " order by id asc"
		set rs = con.execute(stmt)
		if not rs.eof then
			do until rs.eof
				'get the next partid
				ttstmt = "select partid from parts order by partid desc limit 1"
				set ttrs = con.execute(ttstmt)
				lpartid = ttrs("partid")
				newpartid = lpartid + 1

				q = rs("qty")
				nstmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,cost,partcode," _
				& "linettlprice,linettlcost,`date`,complaintid,net,tax"
				if len(rs("supplier")) > 0 then
					nstmt = nstmt & ", supplier "
				end if
				if len(rs("partcategory")) > 0 then
					nstmt = nstmt & ", partcategory "
				end if
				nstmt = nstmt & ") values (" & newpartid & ",'"
				nstmt = nstmt & request.cookies("shopid") & "', '" _
				& rs("partnumber") & "', '" _
				& replace(rs("partdescription"),"'","''") & "', " _
				& rs("partprice") & ", " _
				& rs("qty") & ", " _
				& request("roid") & ", " _
				& rs("partcost") & ", '" _
				& rs("partcode") & "', " _
				& rs("partprice") * rs("qty") & ", " _
				& rs("partcost") * rs("qty") & ", '" _
				& dconv(date) & "', " _
				& request("comid") & ", " _
				& rs("partprice") * rs("qty") & ", '" & rs("tax") & "'"
				if len(rs("supplier")) > 0 then
					nstmt = nstmt & ",'" & replace(rs("supplier"),"'","''") & "'"
				end if
				if len(rs("partcategory")) > 0 then
					nstmt = nstmt & ",'" & rs("partcategory") & "'"
				end if
				nstmt = nstmt & ")"
				response.write nstmt & "<BR>"
				con.execute(nstmt)
				recordAudit "Add Part", "Added Part Number " & ucase(replace(replace(rs("PartNumber"),"'","''"),"%","")) & " to RO#" & request("roid")
				
				' now adjust the inventory allocated
				astmt = "select * from partsinventory where shopid = '" & request.cookies("shopid") _
				& "' and partnumber = '" & rs("partnumber") & "'"
				'response.write stmt
				set ars = con.execute(astmt)
				if not ars.eof then
					ars.movefirst
					oh = ars("OnHand")
					al = ars("allocatted")
					noh = ars("netonhand")
					aid = ars("partid")
					'update the inventory by reducing the on hand and increasing the allocatted
					bstmt = "update partsinventory set onhand = " & oh - q
					bstmt = bstmt & ", allocatted = " & al + q
					bstmt = bstmt & ", netonhand = " & noh - q & " where shopid = '" & request.cookies("shopid") & "' and partid = " & aid
					'response.write stmt
					con.execute(bstmt)
				end if
				rs.movenext
			loop
		end if
	elseif cdbl(trs("flatprice")) > 0 then
		 '****  add the parts and labor without totals and modify to show as part of the 
		 
		stmt = "select * from cannedlabor where shopid = '" & request.cookies("shopid") & "' and cannedjobsid = " & id & " order by id asc"
		set rs = con.execute(stmt)
		if not rs.eof then
			do until rs.eof
				nstmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,complaintid,memorate) values " _
				& "('" & request.cookies("shopid") & "', " _
				& request("roid") & ", 0" _
				& ", " & rs("laborhours") _
				& ", '" _
				& rs("labor") & " (Included in " & trs("jobname") & ")', '" _
				& request("tech") & "', 0" _
				& ", " _
				& request("comid") & "," & request("hourlyrate") & ")"
				con.execute(nstmt)
				recordAudit "Add Labor", "Added Labor as part of a Canned Job - " & replace(ucase(rs("Labor")),"'","''") & " (" & rs("laborhours") & " hours) to RO#" & request("roid")
				rs.movenext
			loop
		end if

		'now get the parts
		stmt = "select * from cannedparts where shopid = '" & request.cookies("shopid") & "' and cannedjobsid = " & id & " order by id asc"
		set rs = con.execute(stmt)
		if not rs.eof then
			do until rs.eof
				'for each x in request.form
					'response.write x & ":" & request(x) & "<BR>"
				'next
				'get the next partid
				ttstmt = "select partid from parts order by partid desc limit 1"
				set ttrs = con.execute(ttstmt)
				lpartid = ttrs("partid")
				newpartid = lpartid + 1

				q = rs("qty")
				nstmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,cost,partcode," _
				& "linettlprice,linettlcost,`date`,complaintid,net,tax"
				if len(rs("supplier")) > 0 then
					nstmt = nstmt & ", supplier "
				end if
				if len(rs("partcategory")) > 0 then
					nstmt = nstmt & ", partcategory "
				end if
				nstmt = nstmt & ") values (" & newpartid & ",'"
				nstmt = nstmt & request.cookies("shopid") & "', '" _
				& rs("partnumber") & "', '" _
				& rs("partdescription") & " (Included in " & trs("jobname") & ")', 0" _
				& ", " _
				& rs("qty") & ", " _
				& request("roid") & ", " _
				& rs("partcost") & ", '" _
				& "New" & "', 0" _
				& ", " _
				& rs("partcost") * rs("qty") & ", '" _
				& dconv(date) & "', " _
				& request("comid") & ", 0" _
				& ", '" & rs("tax") & "'"
				if len(rs("supplier")) > 0 then
					nstmt = nstmt & ",'" & replace(rs("supplier"),"'","''") & "'"
				end if
				if len(rs("partcategory")) > 0 then
					nstmt = nstmt & ",'" & rs("partcategory") & "'"
				end if
				nstmt = nstmt & ")"
				response.write nstmt & "<BR>"
				con.execute(nstmt)
				recordAudit "Add Part", "Added Part Number as part of a Canned Job - " & ucase(replace(replace(rs("PartNumber"),"'","''"),"%","")) & " to RO#" & request("roid")
				
				' now adjust the inventory allocated
				astmt = "select * from partsinventory where shopid = '" & request.cookies("shopid") _
				& "' and partnumber = '" & rs("partnumber") & "'"
				'response.write stmt
				set ars = con.execute(astmt)
				if not ars.eof then
					ars.movefirst
					oh = ars("OnHand")
					al = ars("allocatted")
					noh = ars("netonhand")
					aid = ars("partid")
					'update the inventory by reducing the on hand and increasing the allocatted
					bstmt = "update partsinventory set onhand = " & oh - q
					bstmt = bstmt & ", allocatted = " & al + q
					bstmt = bstmt & ", netonhand = " & noh - q & " where shopid = '" & request.cookies("shopid") & "' and partid = " & aid
					'response.write stmt
					con.execute(bstmt)
				end if
				rs.movenext
			loop
		end if
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1

		 ' ***** now add the canned job with a price
		nstmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,cost,partcode," _
		& "linettlprice,linettlcost,`date`,partcategory,complaintid,net,tax"
		nstmt = nstmt & ", supplier,displayorder) values (" & newpartid & ",'"
		nstmt = nstmt & request.cookies("shopid") & "', '" _
		& "JOB', '" & trs("jobname") _
		& "', " _
		& trs("flatprice") & ", " _
		& "1" & ", " _
		& request("roid") & ", " _
		& "0" & ", '" _
		& "New" & "', " & trs("flatprice") _
		& ", " _
		& "0" & ", '" _
		& dconv(date) & "', '" _
		& "GENERAL PART" & "', " _
		& request("comid") & ", " _
		& trs("flatprice") & ", '" & ucase(trs("taxable"))
		nstmt = nstmt & "', '" & replace(request.cookies("shopname"), "'", "''") & "',0)"
		response.write nstmt & "<BR>"
		con.execute(nstmt)
		recordAudit "Add Canned Job", "Added Canned Job " & ucase(replace(replace(trs("jobname"),"'","''"),"%","")) & " to RO#" & request("roid")

	end if
end function
%>
<head>
<meta content="text/html; charset=windows-1252" http-equiv="Content-Type" />
<title>Untitled 1</title>
<link rel="stylesheet" href="css/main.css" />
<style type="text/css">
.style1 {
	text-align: left;
}
.style2 {
	text-align: center;
}
.style3 {
	border: 1px solid #0155BB;
}
</style>
<script type="text/javascript">

function checkForm(){

	if(document.theform.cannedjob.value.length == 0){
		alert("You must select a Job")
		return
	}
	if(!isNumber(document.theform.hourlyrate.value)){
		alert("You must enter an Hourly Rate")
		return
	}
	document.theform.submit()

}

function isNumber(num) {   
	return (typeof num == 'string' || typeof num == 'number') && !isNaN(num - 0) && num !== ''; 
}; 

</script>
</head>

<body >

<table style="width: 100%" cellpadding="3" cellspacing="0">
	<tr>
		<td class="sbheader" colspan="2">Add a Canned Job</td>
	</tr>
</table>
	<form action="addcannedjobtoro.asp" name="theform" method="post">
	<input name="roid" type="hidden" value="<%=request("roid")%>" />
	<input name="comid" type="hidden" value="<%=request("comid")%>" />

	<%=e%>
	<%
	stmt = "select vehyear,vehmake,vehmodel,vehengine,drivetype from repairorders where shopid = '" & request.cookies("shopid") & "' and roid = " & request("roid")
	set rs = con.execute(stmt)
	if not rs.eof then
		y = rs("vehyear")
		m = rs("vehmake")
		mdl = rs("vehmodel")
		eng = rs("vehengine")
		dt = rs("drivetype")
	end if
		
	%>
		<table style="width:60%" cellpadding="3" cellspacing="0" align="center" class="style3">
	<tr>
		<td style="font-size:14px;" valign="top" class="style2" colspan="2"><b><%=y & " " & m & " " & mdl & " " & eng & " " & dt%></b><br/><br/>
		&nbsp;</td>
	</tr>
	<tr>
		<td style="font-size:14px;" valign="top" class="style1" colspan="2">
		Click a job to select it.&nbsp; To select multiple jobs, hold the CTRL 
		key on your keyboard while clicking on the jobs you want.&nbsp; You can 
		de-select using the same method.&nbsp; For MAC, use the command key 
		while clicking</td>
	</tr>
	<tr>
		<td style="font-size:14px;" valign="top" class="style1">
		<strong>Select Job</strong></td>
		<td valign="top" class="style1">		
		<%
		stmt = "select * from cannedjobs where shopid = '" & request.cookies("shopid") & "' order by jobname"
		set rs = con.execute(stmt)
		if not rs.eof then
		%>

		<select multiple="multiple" name="cannedjob" size="6">
		<%
			do until rs.eof
		%>
		<option value="<%=rs("id")%>"><%=rs("jobname")%></option>
		<%
				rs.movenext
			loop
		%>
		</select>
		<%
		else
			response.write "You have no canned jobs setup.  <a href='cannedjobs/cannedjobs.asp'>Click here to create one</a>"
		end if
		%>
		</td>
	</tr>
	<tr>
		<td style="font-size:14px;" valign="top" align="center" class="style1">
		<strong>Select a Technician</strong></td>
		<td valign="top" class="style1">
            <select size="1" name="tech">
			<%
			strsql = "Select EmployeeLast, EmployeeFirst from employees where showtechlist = 'yes' and shopid = '" & request.cookies("shopid") & "' and Active = 'yes'"
			set rs = con.execute(strsql)
			if not rs.eof then
				rs.movefirst
				while not rs.eof
					response.write "<option value='" & rs("EmployeeLast") & ", " & rs("EmployeeFirst")
					response.write "'>" & rs("EmployeeLast") & ", " & rs("EmployeeFirst") & "</option>"
					rs.movenext
				wend
			end if
			set rs = nothing
			set rs = con.execute("select HourlyRate,hourlyrate2,hourlyrate3,hourlyrate4,hourlyrate5,hourlyrate6,hourlyrate1label,hourlyrate2label,hourlyrate3label,hourlyrate4label,hourlyrate5label,hourlyrate6label from company where shopid = '" & request.cookies("shopid") & "'")
			if not rs.eof then
				rs.movefirst
				hrate = rs("HourlyRate")
				hrate2 = rs("hourlyrate2")
				hrate3 = rs("hourlyrate3")
				hrate4 = rs("HourlyRate4")
				hrate5 = rs("hourlyrate5")
				hrate6 = rs("hourlyrate6")
				
				hlabel1 = rs("hourlyrate1label")
				hlabel2 = rs("hourlyrate2label")
				hlabel3 = rs("hourlyrate3label")
				hlabel4 = rs("hourlyrate4label")
				hlabel5 = rs("hourlyrate5label")
				hlabel6 = rs("hourlyrate6label")

			end if
			set rs = nothing
			%>
             </select>
		
		&nbsp;</td>
	</tr>
	<tr>
		<td style="font-size:14px;" valign="top" class="style1"><%=e%>
		<strong>Hourly Rate</strong></td>
		<td valign="top" class="style1">
            <%
            if len(hlabel2) > 0 then
            %>
            <select style="text-transform:uppercase" name="hourlyrate" class="style15">
				<option value="<%=hrate%>"><%=hlabel1 & " - " & hrate%></option>
				<option value="<%=hrate2%>"><%=hlabel2 & " - " & hrate2%></option>
				<%
				if len(hlabel3) >= 0 then
				%>
				<option value="<%=hrate3%>"><%=hlabel3 & " - " & hrate3%></option>
				<%
				end if
				%>
				<%
				if len(hlabel4) >= 0 then
				%>
				<option value="<%=hrate4%>"><%=hlabel4 & " - " & hrate4%></option>
				<%
				end if
				%>
				<%
				if len(hlabel5) >= 0 then
				%>
				<option value="<%=hrate5%>"><%=hlabel5 & " - " & hrate5%></option>
				<%
				end if
				%>
				<%
				if len(hlabel6) >= 0 then
				%>
				<option value="<%=hrate6%>"><%=hlabel6 & " - " & hrate6%></option>
				<%
				end if
				%>

			</select>
			<%
			else
			%>
			<input value="<%=hrate%>" onblur="javascript:this.style.backgroundColor='white';document.theform.calcrate.value=document.theform.hours.value*document.theform.rate.value" onfocus="javascript:this.style.backgroundColor='yellow'" type="text" name="hourlyrate" size="11" style="font-variant: small-caps; border-style: solid; border-color: gray;height:24px;width:200px" class="style14">
           <%
           end if
           %>

	</tr>

	<tr>
		<td style="font-size:14px;" valign="top" colspan="2" class="style2">
		<input onclick="checkForm()" name="Button1" type="button" value="Add Job" />
		<input onclick="location.href='ro.asp?roid=<%=request("roid")%>'" name="Button2" type="button" value="Cancel" /></td>
	</tr>

		</table>
</form>


</body>

</html>
