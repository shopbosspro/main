<!-- #include file=aspscripts/conn.asp -->
<%
newroid = request("roid")
origroid =  request("origroid")
origcomid = request("origcomid")

' get the recommend info from the recommend table
stmt = "select * from recommend where shopid = '" & request.cookies("shopid") & "' and roid = " & origroid
response.write stmt & "<Br>"
set trs = con.execute(stmt)
c = 0
if not trs.eof then
	do until trs.eof
		stmt = "insert into recommend (shopid,roid,`desc`,totalrec,originalcomplaintid,originalroid,technotes) values ('" & request.cookies("shopid") _
		& "', " & newroid & ", '" & replace(trs("desc"),"'","''") & "', " & trs("totalrec") & ", " & trs("comid") & ", " & origroid
		technotes = trs("technotes")
		if not isnull(technotes) then
			technotes = replace(technotes,"'","''")
		else
			technotes = ""
		end if
		stmt = stmt & ", '" & technotes & "')"
		response.write stmt & "<BR><BR>"
		con.execute(stmt)
		
		' now get the id for the item just added
		stmt = "select * from recommend where shopid = '" & request.cookies("shopid") & "' and roid = " & newroid 
		response.write stmt & "<BR><BR>"
		set rs = con.execute(stmt)
		newrecid = cdbl(rs("id")) + c
		
		' get the recommend labor from the original recommend
		stmt = "select * from recommendlabor where recid = " & trs("id")
		response.write stmt & "<BR>"
		set rs = con.execute(stmt)
		if not rs.eof then
			do until rs.eof
		
				' add labor to the recommendlabor table 
				stmt = "insert into recommendlabor (recid,shopid,roid,`desc`,rate,hours,total,tech) values (" & newrecid & ", '" & request.cookies("shopid") _
				& "', " & newroid & ", '" & replace(rs("desc"),"'","''") & "', " & rs("rate") & ", " & rs("hours") & ", " & rs("total") & ", '" & rs("tech") & "')"
				response.write stmt & "<BR>"
				con.execute(stmt)
				rs.movenext
			loop
		end if
		
		' get the parts from the original recommend
		stmt = "select * from recommendparts where recid = " & trs("id")
		set rs = con.execute(stmt)
		response.write stmt & "<BR>"
		if not rs.eof then
			do until rs.eof
		
				' add part to the recommendpart table 
				stmt = "insert into recommendparts (shopid,recid,partnumber,partdesc,partprice,quantity,roid,supplier,cost,partinvoicenumber," _
				& "partcode,linettlprice,linettlcost,`date`,partcategory,discount,net,tax) values ('" & request.cookies("shopid") _
				& "', " & newrecid & ", '" & rs("partnumber") & "', '" & replace(rs("partdesc"),"'","''") & "', " & rs("partprice") _
				& ", " & rs("quantity") & ", " & newroid & ", '" & rs("supplier") & "', " & rs("cost") & ", '" & rs("partinvoicenumber") & "', '" _
				& rs("partcode") & "', " & rs("linettlprice") & ", " & rs("linettlcost") & ", '" & dconv(date) & "', '" & rs("partcategory") _
				& "', " & rs("discount") & ", " & rs("net") & ", '" & rs("tax") & "')"
				response.write stmt & "<BR>"
				con.execute(stmt)
				rs.movenext
			loop
		end if
		response.write "LOOP<BR>"
		c = c + 1
		trs.movenext
	loop
end if
%>
