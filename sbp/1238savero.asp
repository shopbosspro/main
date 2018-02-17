<!-- #include file=aspscripts/adovbs.asp -->
<!-- #include file=aspscripts/conn.asp -->
<%

'on error resume next
shopid = request.cookies("shopid")
useaccounting = request.cookies("useaccounting")

'******************************first get the finaldate from this ro***************************
rttl = 0
stmt = "select finaldate,howpaid,howpaid2,amtpaid1,amtpaid2,totalro from repairorders where shopid = '" & shopid _
& "' and roid = " & request("roid")
response.write stmt & ";<BR>"
set rs = con.execute(stmt)
cfinaldate = rs("finaldate")
ttlro = rs("totalro")
if ucase(rs("howpaid")) <> "ACCOUNT" and ucase(rs("howpaid")) <> "PURCHASE ORDER" then
	rttl = rttl + rs("amtpaid1")
end if
if ucase(rs("howpaid2")) <> "ACCOUNT" and ucase(rs("howpaid")) <> "PURCHASE ORDER" then
	rttl = rttl + rs("amtpaid2")
end if
set rs = nothing
'****************************** end final date ***********************************************

'**************************Update ro with all fields from the ro page ************************
stmt = "update repairorders set "
for each thing in request.form
	if left(thing,2) = "db" then
		fldname = mid(thing,3,len(thing)-1)
		fldvalue = request.form(thing)
		typeoffield = "fldtype" + fldname
		fldtype = request(typeoffield)
		'response.write fldname & "*" & fldvalue & "*" & fldtype & "<BR>"
		select case fldtype
			case 3,4,5,131
				if len(fldvalue) > 0 then
					fldvalue = fldvalue
				else
					fldvalue = 0
				end if
				fldvalue = replace(fldvalue,",","")
				if fldvalue = "NaN" then fldvalue = 0
				stmt = stmt & fldname & " = " & fldvalue & ","
			case 202,203
				stmt = stmt & fldname & " = '" & replace(fldvalue,"'","''") & "',"
			case 133
				if len(fldvalue) > 0 then
					stmt = stmt & fldname & " = '" & dconv(fldvalue) & "',"
				else
					stmt = stmt & fldname & " = '0000-00-00',"
				end if
		end select
				
	end if
next


if right(stmt,1) = "," then
	stmt = left(stmt,len(stmt)-1)
else
	stmt = stmt
end if
stmt = stmt & " where shopid = '" & shopid & "' and ROID = " & request("roid")
response.write stmt & ";<br>"
con.execute(stmt)

'*********************************  end update of ro from ro page *********************************

'*********************************  setup the redirection after saving ****************************

if instr(request("dir"), "history.asp") then
	redirpage = request("dir")
else
	redirpage = request("dir") & "sid="&request("sid")&"&pid="&request("pid")
	redirpage=redirpage & "&itype="&request("itype")&"&lid="&request("lid")&"&roid="&request("roid")
end if
'********************************* end the redirection setup ***************************************

' ********************************  update from other table info ***********************************
pcstmt = "select coalesce(sum(cost*quantity),0) as pcost from parts where deleted = 'no' and shopid = '" & shopid & "' and ROID = " & request("roid")
response.write pcstmt & ";<br>"
set pcrs = con.execute(pcstmt)
if not pcrs.eof then
	pcrs.movefirst
	pcost = pcrs("pcost")
	if isnull(pcost) then pcost = 0
else
	pcost = 0
end if
' ***********************************  end parts cost from parts table *****************************

stmt = "select sum(amt) as prrt from accountpayments where ucase(ptype) != 'PURCHASE ORDER' and ucase(ptype) != 'ACCOUNT' and shopid = '" & shopid & "' and roid = " & request("roid")
response.write stmt & ";<br>"
set rs = con.execute(stmt)
rs.movefirst
if len(rs("prrt")) > 0 then
	prrt = rs("prrt")
else
	prrt = 0
end if
tpmts = (rttl + prrt)
newbalance = (ttlro - tpmts)

if isnull(newbalance) then
	newbalance = 0
else
	newbalance = formatnumber(newbalance,2)
end if
newbalance = cstr(newbalance)
newbalance = replace(newbalance,",","")
'response.write newbalance

if request("dbrotype") = "No Charge" then
	newbalance = "0"
end if

if request("dbrotype") = "No Approval" then
	newbalance = "0"
end if


upcstmt = "Update repairorders set balance = " & newbalance & ", totalro = totalprts+totallbr+totalsublet+totalfees-discountamt+salestax, PartsCost = " & pcost & " where shopid = '" & shopid & "' and ROID = " & request("roid")
response.write upcstmt & ";<br>"
con.execute(upcstmt)

' ******************************** end parts cost update ********************************************

' ********************************* update inventory ************************************************
'check for inventory part used
if ucase(request("Status")) = "CLOSED" then


	if useaccounting = "yes" then
		' remove all previous parts expenses, then post new amounts
		stmt = "delete from unpostedexpenses where shopid = '" & shopid _
		& "' and roid = " & request("roid")
		con.execute(stmt)
	end if


	istmt = "select `shopid`,`PartID`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax`,`overridematrix`,`posted` from parts where deleted = 'no' and shopid = '" & shopid & "' and ROID = " & request("roid")
	response.write istmt & ";<br>"
	set irs = con.execute(istmt)
	if not irs.eof then
		irs.movefirst
		while not irs.eof
		
		
			' post the parts cost to the unpostedexpenses table
			if useaccounting = "yes" and lcase(irs("posted")) <> "yes" then
				stmt = "insert into unpostedexpenses (shopid,paidto,amount,category,memo,udate,roid,partid) values ('" _
				& shopid & "','" & replace(irs("supplier"),"'","''") & "', " & irs("linettlcost") & ", 'Cost of Goods Sold','Parts cost for RO# " & request("roid") _
				& "', '" & dconv(request("statusdate")) & "', " & request("roid") & "," & irs("partid") & ")"
				con.execute(stmt)
			end if
		
			'now check for part in inventory
			iistmt = "select partnumber,allocatted,onhand,partid from partsinventory where shopid = '" & shopid & "' and PartNumber = '" & irs("PartNumber") & "'"
			response.write iistmt & ";<br>"
			set iirs = con.execute(iistmt)
			if not iirs.eof then
				iirs.movefirst
				'check to part to have allocated qty
				if iirs("Allocatted") > 0 then
					newal = iirs("Allocatted") - irs("Quantity")
				else
					newal = 0
				end if
				'check for onhand qty
				if iirs("OnHand") > 0 then
					newoh = iirs("OnHand") - irs("Quantity")
				else
					newoh = 0
				end if
				'calculate netonhand
				newnet = newoh - newal
				upstmt = "update partsinventory set OnHand = " & newoh & ", Allocatted = " & newal & ", NetOnHand = " & newnet & " where shopid = '" & shopid & "' and PartNumber = '" & iirs("PartNumber") & "'"
				response.write upstmt & ";<br>"
				con.execute(upstmt)
			end if
			irs.movenext
		wend
	end if
end if

' *** set displayorder 
	
' run the list of complaints and use the id to change the display order
'stmt = "select complaintid from complaints where cstatus = 'no' and shopid = '" & request.cookies("shopid") _
'& "' and roid = " & request("roid")
'response.write stmt & ";<BR>"
'set rs = con.execute(stmt)
'if not rs.eof then
'	do until rs.eof
'		
'		'get the value of the appropriate displayorder 
'		newdisplayorder = request("displayorder"&rs("complaintid"))
'		if len(newdisplayorder) = 0 then newdisplayorder = 0
'		if instr(newdisplayorder,",") > 0 then
'			newar = split(newdisplayorder,",")
'			newdisplayorder = newar(0)
'		end if
'		response.write "newdisplayorder:" & newdisplayorder & "<BR>"
'		stmt = "update complaints set displayorder = " & newdisplayorder & " where shopid = '" _
'		& request.cookies("shopid") & "' and complaintid = " & rs("complaintid") & " and roid = " & request("roid")
'		response.write "<BR><BR>" & stmt & ";<BR>	<BR>"
'		con.execute(stmt)
'		
'		rs.movenext
'	loop
'end if

if len(request.querystring("comid")) > 0 then
	stmt = "update complaints set acceptdecline = '" & request("comstat") & "' where complaintid = " _
	& request("comid") & " and shopid = '" & shopid & "' and roid = " & request("roid")
	response.write stmt & ";<br>"
	con.execute(stmt)
	
	' if declined update the recommended repairs
	if request("comstat") = "Declined" then
		'get the complaint from the complaint table and ad it to recommended repairs
		stmt = "select complaint,techreport from complaints where roid = " & request("roid") & " and shopid = '" _
		& shopid & "' and complaintid = " & request("comid")
		response.write stmt & ";<br>"
		set rs = con.execute(stmt)
		
		' now add the issue to the recommend table
		comp = rs("complaint")
		techreport = rs("techreport")
		if not isnull(techreport) then
			techreport = replace(techreport,"'","''")
		end if
		if not isnull(comp) then
			comp = replace(comp,"'","''")
		end if
		stmt = "insert into recommend (shopid,roid,comid,`desc`,technotes) values ('" & shopid & "', " & request("roid") _
		& ", " & request("comid") & ", 'Customer Declined: " & comp & "', '" & techreport & "')"
		response.write stmt & ";<br>"
		con.execute(stmt)
		
		' now get the id
		stmt = "select id from recommend where shopid = '" & shopid & "' and roid = " & request("roid") & " and comid = " & request("comid")
		response.write stmt & ";<br>"
		set comrs = con.execute(stmt)
		recid = comrs("id")
		
		' add all parts and labor to the recommend parts and recommendlabor table
		stmt = "select `shopid`,`LaborID`,`ROID`,`HourlyRate`,`LaborHours`,`Labor`,`Tech`,`LineTotal`,`LaborOp`,`complaintid`,`techrate` from labor where deleted = 'no' and shopid = '" & shopid & "' and roid = " & request("roid") & " and complaintid = " & request("comid")
		response.write stmt & ";<br>"
		set trs = con.execute(stmt)
		if not trs.eof then
			do until trs.eof
				totalrec = totalrec + trs("linetotal")
				stmt = "insert into recommendlabor (recid,shopid,roid,comid,`desc`,rate,hours,`total`,tech) values (" & recid & ", '" _
				& shopid & "', " & request("roid") & ", " & trs("complaintid") & ", '" & trs("labor") & "', " & trs("hourlyrate") & ", " _
				& trs("laborhours") & ", " & trs("linetotal") _
				& ", '" & trs("tech") & "')"
				response.write stmt & "<BR>"
				con.execute(stmt)
				trs.movenext
			loop
		end if
		
		stmt = "select 	`shopid`,`PartID`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax`,`overridematrix`,`posted` from parts where deleted = 'no' and shopid = '" & shopid & "' and roid = " & request("roid") & " and complaintid = " & request("comid")
		response.write stmt & ";<br>"
		set trs = con.execute(stmt)
		if not trs.eof then
			do until trs.eof
				totalrec = totalrec + trs("linettlprice")
				if len(trs("date")) > 0 then
					trsdate = dconv(trs("date"))
				else
					trsdate = "0000-00-00"
				end if
				stmt = "insert into recommendparts (shopid,recid,partnumber,partdesc,partprice,quantity,roid,supplier,cost,partinvoicenumber,partcode," _
				& "linettlprice,linettlcost,`date`,partcategory,complaintid,discount,`net`,tax) values ('" & trs("shopid") & "', " & recid & ", '" & trs("partnumber") _
				& "', '" & trs("partdesc") & "', " & trs("partprice") & ", " & trs("quantity") & ", " & trs("roid") & ", '" & replace(trs("supplier"),"'","''") & "', " _
				& trs("cost") & ", '" & trs("partinvoicenumber") & "', '" & trs("partcode") & "', " & trs("linettlprice") & ", " & trs("linettlcost") _
				& ", '" & trsdate & "', '" & trs("partcategory") & "', " & trs("complaintid") & ", " & trs("discount") & ", " & trs("net") _
				& ", '" & trs("tax") & "')"
				response.write stmt & ";<BR>"
				con.execute(stmt)
				trs.movenext
			loop
		end if
		
		' now update the recommend with a total
		if len(totalrec) > 0 then
			if isnumeric(totalrec) then
				totalrec = totalrec
			else
				totalrec = 0
			end if
		else
			totalrec = 0
		end if
		stmt = "update recommend set totalrec = " & totalrec & " where id = " & recid
		con.execute(stmt)
		
		'now delete the parts and labor and sublet associated with the complaint
		stmt = "delete from parts where shopid = '" & shopid & "' and complaintid = " & request("comid") & " and roid = " & request("roid")
		con.execute(stmt)
		stmt = "delete from labor where shopid = '" & shopid & "' and complaintid = " & request("comid") & " and roid = " & request("roid")
		con.execute(stmt)
		stmt = "delete from sublet where shopid = '" & shopid & "' and complaintid = " & request("comid") & " and roid = " & request("roid")
		con.execute(stmt)

	end if
	if request("comstat") <> "Declined" then
		response.write "not declined<BR><BR>"
		comidd = request.querystring("comid")
		if isnumeric(comidd) then
			comidd = cdbl(comidd)
		else
			comidd = 0
		end if
		if comidd > 0 then
			'get the complaint from the complaint table and ad it to recommended repairs
			stmt = "select `shopid`,`complaintid`,`roid`,`complaint`,`techreport`,`tech`,`acceptdecline`,`issue`,`displayorder`,`id` from complaints where roid = " & request("roid") & " and shopid = '" _
			& shopid & "' and complaintid = " & request("comid")
			response.write stmt & ";<br>"
			set rs = con.execute(stmt)
			if not rs.eof then
							
				' now check for recommened repairs to restore
				stmt = "select `id`,`shopid`,`roid`,`comid`,`desc`,`totalrec`,`originalcomplaintid`,`originalroid`,`technotes` from recommend where shopid = '" & shopid & "' and comid = " _
				& rs("complaintid") & " and roid = " & rs("roid")
				response.write stmt & ";<br>"
				set trs = con.execute(stmt)
				if not trs.eof then
					
					' get the parts for the recommend
					stmt = "select recid, id, `shopid`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax` from recommendparts where recid = " & trs("id")
					response.write stmt & ";<br>"
					set prs = con.execute(stmt)
					if not prs.eof then
						do until prs.eof	
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1
							if len(prs("date")) > 0 then
								prsdate = dconv(prs("date"))
							else
								prsdate = "0000-00-00"
							end if
							stmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,supplier" _
							& ",cost,partinvoicenumber,partcode,linettlprice,linettlcost,`date`,partcategory," _
							& "complaintid,discount,net,tax) values (" & newpartid & ",'" & prs("shopid") & "', '" _
							& prs("partnumber") & "', '" _
							& prs("partdesc") & "', " _
							& prs("partprice") & ", " _
							& prs("quantity") & ", " _
							& prs("roid") & ", '" _
							& replace(prs("supplier"),"'","''") & "', " _
							& prs("cost") & ", '" _
							& prs("partinvoicenumber") & "', '" _
							& prs("partcode") & "', " _
							& prs("linettlprice") & ", " _
							& prs("linettlcost") & ", '" _
							& prsdate & "', '" _
							& prs("partcategory") & "', " _
							& prs("complaintid") & ", " _
							& prs("discount") & ", " _
							& prs("net") & ", '" _
							& prs("tax") & "')"
							response.write stmt & "<BR>"
							con.execute(stmt)
							stmt = "delete from recommendparts where id = " & prs("id")
							'response.write stmt
							con.execute(stmt)
							stmt = "delete from recommend where id = " & prs("recid")
							con.execute(stmt)
							prs.movenext
						loop
					end if
					
					' get the labor for the recommend
					stmt = "select `id`,`recid`,`shopid`,`roid`,`comid`,`desc`,`rate`,`hours`,`total`,`tech` from recommendlabor where recid = " & trs("id")
					response.write stmt & ";<br>"
					set lrs = con.execute(stmt)
					if not lrs.eof then
						do until lrs.eof
							stmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,complaintid) values ('" & shopid _
							& "', " & lrs("roid") _
							& ", " & lrs("rate") _
							& ", " & lrs("hours") _
							& ", '" & lrs("desc") _
							& "', '" & lrs("tech") _
							& "', " & lrs("total") _
							& ", " & lrs("comid") & ")"
							response.write stmt & "<BR>"
							con.execute(stmt)
							stmt = "delete from recommendlabor where id = " & lrs("id")
							'response.write stmt & "<BR>"
							con.execute(stmt)
							stmt = "delete from recommend where id = " & lrs("recid")
							con.execute(stmt)
							lrs.movenext
						loop
					end if
					stmt = "delete from recommend where id = " & trs("id")
					con.execute(stmt)
				end if
			
			else   ' ******** the original complaint must have been lost
			
				'  there is no complaint to restore, so we get the declination description and use that for the complaint
				response.write "lost orig complaint<BR><BR>"
				if len(request.querystring("recid")) > 0 then
					
					'create the complaint
					stmt = "select complaintid from complaints order by complaintid desc"
					response.write stmt & ";<br>"
					set rs = con.execute(stmt)
					lastcomid = cdbl(rs("complaintid")) + 1
					
					' get the complaint from the recommend table
					stmt = "select `desc`, technotes from recommend where id = " & request("recid")
					response.write stmt & ";<br>"
					set rs = con.execute(stmt)
					comdesc = rs("desc")
					technotes = replace(rs("technotes"),"'","''")
					if instr(comdesc,"Customer Declined: ") then
						comdesc = replace(comdesc,"Customer Declined: ","")
					end if
					comdesc = replace(comdesc,"'","''")
					
					stmt = "insert into complaints (shopid,complaintid,roid,complaint,acceptdecline,displayorder,techreport) values ('" & shopid _
					& "', " & lastcomid & ", " & request("roid") & ", '" & comdesc & "', 'Pending',1,'" & technotes & "')"
					con.execute(stmt)
					
					'get the complaint from the complaint table and ad it to recommended repairs
					stmt = "select * from complaints where roid = " & request("roid") & " and shopid = '" _
					& shopid & "' and complaintid = " & lastcomid
					response.write stmt & ";<br>"
					'response.write "<BR><BR>" & stmt & "<BR><BR>"
					set rs = con.execute(stmt)
					if not rs.eof then
									
						' now check for recommened repairs to restore
						stmt = "select `id`,`shopid`,`roid`,`comid`,`desc`,`totalrec`,`originalcomplaintid`,`originalroid`,`technotes` from recommend where id = " & request.querystring("recid")
						'response.write "<BR><BR><BR>" & stmt & "<BR><BR>"
						response.write stmt & ";<br>"
						set trs = con.execute(stmt)
						if not trs.eof then
							
							' get the parts for the recommend
							stmt = "select recid,id,`shopid`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax` from recommendparts where recid = " & trs("id")
							response.write stmt & ";<br>"
							response.write "<BR><BR><BR>" & stmt & "<BR><BR>"
							set prs = con.execute(stmt)
							if not prs.eof then
								do until prs.eof
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1

									stmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,supplier" _
									& ",cost,partinvoicenumber,partcode,linettlprice,linettlcost,`date`,partcategory," _
									& "complaintid,discount,net,tax) values (" & newpartid & ",'" & prs("shopid") & "', '" _
									& prs("partnumber") & "', '" _
									& prs("partdesc") & "', " _
									& prs("partprice") & ", " _
									& prs("quantity") & ", " _
									& prs("roid") & ", '" _
									& replace(prs("supplier"),"'","''") & "', " _
									& prs("cost") & ", '" _
									& prs("partinvoicenumber") & "', '" _
									& prs("partcode") & "', " _
									& prs("linettlprice") & ", " _
									& prs("linettlcost") & ", '" _
									& dconv(prs("date")) & "', '" _
									& prs("partcategory") & "', " _
									& lastcomid & ", " _
									& prs("discount") & ", " _
									& prs("net") & ", '" _
									& prs("tax") & "')"
									'response.write stmt & "<BR>"
									con.execute(stmt)
									stmt = "delete from recommendparts where id = " & prs("id")
									'response.write stmt
									con.execute(stmt)
									stmt = "delete from recommend where id = " & prs("recid")
									con.execute(stmt)
									prs.movenext
								loop
							end if
							
							' get the labor for the recommend
							stmt = "select `id`,`recid`,`shopid`,`roid`,`comid`,`desc`,`rate`,`hours`,`total`,`tech` from recommendlabor where recid = " & trs("id")
							response.write stmt & ";<br>"
							set lrs = con.execute(stmt)
							if not lrs.eof then
								do until lrs.eof
									stmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,complaintid) values ('" & shopid _
									& "', " & lrs("roid") _
									& ", " & lrs("rate") _
									& ", " & lrs("hours") _
									& ", '" & lrs("desc") _
									& "', '" & lrs("tech") _
									& "', " & lrs("total") _
									& ", " & lastcomid & ")"
									'response.write stmt & "<BR>"
									con.execute(stmt)
									stmt = "delete from recommendlabor where id = " & lrs("id")
									'response.write stmt & "<BR>"
									con.execute(stmt)
									stmt = "delete from recommend where id = " & lrs("recid")
									con.execute(stmt)
									lrs.movenext
								loop
							end if
							stmt = "delete from recommend where id = " & trs("id")
							con.execute(stmt)
						end if
					end if	
				end if

			end if	
		else
			
			'  there is no complaint to restore, so we get the declination description and use that for the complaint
			if len(request.querystring("recid")) > 0 then
				
				'create the complaint
				stmt = "select complaintid from complaints order by complaintid desc"
				response.write stmt & ";<br>"
				set rs = con.execute(stmt)
				lastcomid = cdbl(rs("complaintid")) + 1
				
				' get the complaint from the recommend table
				stmt = "select `desc`, technotes from recommend where id = " & request("recid")
				response.write stmt & ";<br>"
				set rs = con.execute(stmt)
				comdesc = rs("desc")
				technotes = replace(rs("technotes"),"'","''")
				if instr(comdesc,"Customer Declined: ") then
					comdesc = replace(comdesc,"Customer Declined: ","")
				end if
				comdesc = replace(comdesc,"'","''")
				
				stmt = "insert into complaints (shopid,complaintid,roid,complaint,acceptdecline,displayorder,techreport) values ('" & shopid _
				& "', " & lastcomid & ", " & request("roid") & ", '" & comdesc & "', 'Pending',1,'" & technotes & "')"
				con.execute(stmt)
				
				'get the complaint from the complaint table and ad it to recommended repairs
				stmt = "select * from complaints where roid = " & request("roid") & " and shopid = '" _
				& shopid & "' and complaintid = " & lastcomid
				response.write stmt & ";<br>"
				'response.write "<BR><BR>" & stmt & "<BR><BR>"
				set rs = con.execute(stmt)
				if not rs.eof then
								
					' now check for recommened repairs to restore
					stmt = "select `id`,`shopid`,`roid`,`comid`,`desc`,`totalrec`,`originalcomplaintid`,`originalroid`,`technotes` from recommend where id = " & request.querystring("recid")
					'response.write "<BR><BR><BR>" & stmt & "<BR><BR>"
					response.write stmt & ";<br>"
					set trs = con.execute(stmt)
					if not trs.eof then
						
						' get the parts for the recommend
						stmt = "select recid,id,`shopid`,`PartNumber`,`PartDesc`,`PartPrice`,`Quantity`,`ROID`,`Supplier`,`Cost`,`PartInvoiceNumber`,`PartCode`,`LineTTLPrice`,`LineTTLCost`,`Date`,`PartCategory`,`complaintid`,`discount`,`net`,`tax` from recommendparts where recid = " & trs("id")
						response.write stmt & ";<br>"
						response.write "<BR><BR><BR>" & stmt & "<BR><BR>"
						set prs = con.execute(stmt)
						if not prs.eof then
							do until prs.eof
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1

								stmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,supplier" _
								& ",cost,partinvoicenumber,partcode,linettlprice,linettlcost,`date`,partcategory," _
								& "complaintid,discount,net,tax) values (" & newpartid & ",'" & prs("shopid") & "', '" _
								& prs("partnumber") & "', '" _
								& prs("partdesc") & "', " _
								& prs("partprice") & ", " _
								& prs("quantity") & ", " _
								& prs("roid") & ", '" _
								& replace(prs("supplier"),"'","''") & "', " _
								& prs("cost") & ", '" _
								& prs("partinvoicenumber") & "', '" _
								& prs("partcode") & "', " _
								& prs("linettlprice") & ", " _
								& prs("linettlcost") & ", '" _
								& dconv(prs("date")) & "', '" _
								& prs("partcategory") & "', " _
								& lastcomid & ", " _
								& prs("discount") & ", " _
								& prs("net") & ", '" _
								& prs("tax") & "')"
								'response.write stmt & "<BR>"
								con.execute(stmt)
								stmt = "delete from recommendparts where id = " & prs("id")
								'response.write stmt
								con.execute(stmt)
								stmt = "delete from recommend where id = " & prs("recid")
								con.execute(stmt)
								prs.movenext
							loop
						end if
						
						' get the labor for the recommend
						stmt = "select `id`,`recid`,`shopid`,`roid`,`comid`,`desc`,`rate`,`hours`,`total`,`tech` from recommendlabor where recid = " & trs("id")
						response.write stmt & ";<br>"
						set lrs = con.execute(stmt)
						if not lrs.eof then
							do until lrs.eof
								stmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,complaintid) values ('" & shopid _
								& "', " & lrs("roid") _
								& ", " & lrs("rate") _
								& ", " & lrs("hours") _
								& ", '" & lrs("desc") _
								& "', '" & lrs("tech") _
								& "', " & lrs("total") _
								& ", " & lastcomid & ")"
								'response.write stmt & "<BR>"
								con.execute(stmt)
								stmt = "delete from recommendlabor where id = " & lrs("id")
								'response.write stmt & "<BR>"
								con.execute(stmt)
								stmt = "delete from recommend where id = " & lrs("recid")
								con.execute(stmt)
								lrs.movenext
							loop
						end if
						stmt = "delete from recommend where id = " & trs("id")
						con.execute(stmt)
					end if
				end if	
			end if

		
		end if
		
	end if
	'response.write redirpage
	response.redirect redirpage
else
	'response.write redirpage
	response.redirect redirpage
end if
%>
