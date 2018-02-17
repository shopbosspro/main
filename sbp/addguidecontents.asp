<!-- #include file=aspscripts/conn.asp -->
<%
on error resume next

'check for duplicate items in laboritemslist


'*** get the list of parts to add
if len(request("partitems")) > 1 then
	response.write request("partitems")
	par = split(request("partitems"),"~")
	for j = lbound(par) to ubound(par)-1
		'response.write replace(request("part*" & par(j) & "*pr"),"","")
		'get the next partid
		ttstmt = "select partid from parts order by partid desc limit 1"
		set ttrs = con.execute(ttstmt)
		lpartid = ttrs("partid")
		newpartid = lpartid + 1

		stmt = "insert into parts (partid,shopid,partnumber,partdesc,partprice,quantity,roid,supplier,cost,partcode,linettlprice,linettlcost,`date`,partcategory"
		stmt = stmt & ",complaintid,net,tax) values (" & newpartid & ",'" & request("shopid") & "', '"
		stmt = stmt & request("part*" & par(j) & "*pn") & "', '"
		stmt = stmt & request("part*" & par(j) & "*de") & "', "
		stmt = stmt & replace(request("part*" & par(j) & "*pr"),"$","") & ", "
		stmt = stmt & request("part*" & par(j) & "*qty") & ", "
		stmt = stmt & request("roid") & ", '"
		stmt = stmt & "unk" & "', "
		stmt = stmt & replace(request("part*" & par(j) & "*pr"),"$","") & ", '"
		stmt = stmt & "New" & "', "
		stmt = stmt & cint(request("part*" & par(j) & "*qty"))*cdbl(replace(request("part*" & par(j) & "*pr"),"$","")) & ", "
		stmt = stmt & cint(request("part*" & par(j) & "*qty"))*cdbl(replace(request("part*" & par(j) & "*pr"),"$","")) & ", '"
		stmt = stmt & dconv(date) & "', '"
		stmt = stmt & "GENERAL PART" & "', "
		stmt = stmt & request("comid") & ", "
		stmt = stmt & cint(request("part*" & par(j) & "*qty"))*cdbl(replace(request("part*" & par(j) & "*pr"),"$","")) & ",'yes')"
		response.write stmt & "<BR>"
		con.execute(stmt)
	
	next
end if

'*** get shop hourly rate
stmt = "select hourlyrate from company where shopid = '" & request("shopid") & "'"
set rs = con.execute(stmt)
shoprate = rs("hourlyrate")
set rs = nothing

if len(request("laboritems")) > 1 then
	response.write request("laboritems")& "<BR>"
	
	par = split(request("laboritems"),"~")
	for j = lbound(par) to ubound(par)-1
		
		stmt = "insert into labor (shopid,roid,hourlyrate,laborhours,labor,tech,linetotal,laborop,complaintid,techrate) values ('" _
		& request("shopid") & "', " _
		& request("roid") & ", " _
		& shoprate & ", " _
		& request("labor*" & par(j) & "*ti") & ", '" _
		& request("labor*" & par(j) & "*de") & "', '" _
		& request("labor*" & par(j) & "*tech") & "', " _
		& cdbl(shoprate)*cdbl(request("labor*" & par(j) & "*ti")) & ", '" _
		& "', " _
		& request("comid") & ", " _
		& "0" & ")" 
		'response.write stmt & "<BR>"
		con.execute(stmt)
	next
end if

response.redirect "ro.asp?roid=" & request("roid")
%>
