<!-- #include file=aspscripts/auth.asp --> 
<!-- #include file=aspscripts/conn.asp -->
<!-- #include file=javascripts/checksubcust2.js -->
<!-- #include file=aspscripts/adovbs.asp -->
<%
for each x in request.form
	'response.write x & ":" & request(x) & "<BR>"
next
'where is first slash
dpart = split(request("d"),"/")
mmonth = dpart(0)
dday = right("0" & dpart(1),2)
'if dday <= 9 then dday = "0" & dday
yyear = right(dpart(2),2)

' now convert the time to 12 hour
if len(request("t")) > 0 then
	myhour = cint(hour(request("t")))
	if myhour >= 13 then
		hour12 = myhour - 12
		mymin = cstr(minute(request("t")))
		if len(mymin) = 1 then mymin = "0" & mymin
		selecttime = hour12 & ":" & mymin  & " PM"
	else
		hour12 = myhour
		mymin = cstr(minute(request("t")))
		if len(mymin) = 1 then mymin = "0" & mymin
		selecttime = hour12 & ":" & mymin & " AM"
	end if
else
	selecttime = ""
end if
'response.write selecttime

'get customer info if sent
if request("custid") then
if len(request("origshopid")) > 0 then
	stmt = "select * from customer where shopid = '" & request("origshopid") & "' and CustomerID = " & request("custid")
else
	stmt = "select * from customer where shopid = '" & request.cookies("shopid") & "' and CustomerID = " & request("custid")
end if
set rs = con.execute(stmt)
if not rs.eof then
	rs.movefirst
	uf = "yes"
	lastname = rs("LastName")
	firstname = rs("FirstName")
	ac = left(rs("HomePhone"),3)
	pf = mid(rs("HomePhone"),4,3)
	last = right(rs("HomePhone"),4)
	
	wac = left(rs("workPhone"),3)
	wpf = mid(rs("workPhone"),4,3)
	wlast = right(rs("workPhone"),4)
	
	cac = left(rs("cellPhone"),3)
	cpf = mid(rs("cellPhone"),4,3)
	clast = right(rs("cellPhone"),4)

	
	
	vnum = request("vehnum")
	yearfname = "year" & vnum
	yearf=request(yearfname)
	makefname = "make" & vnum
	make = request(makefname)
	modelfname="model" & vnum
	model=request(modelfname)
	vidfname="vid" & vnum
	vid=request(vidfname)
	mcontact = rs("contact")
else
	uf = "no"
end if
end if
%>
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/main.css">
<title><%=request.cookies("shopname")%></title>
<script type="text/javascript">


function checkForm(){
	d = document.theform
	if (d.time.value == "none"){
		alert("You must select a time for the appointment")
		return false;
	}
	if(d.lastname.value.length == 0){
		alert("Last Name is required")
		return false;
	}
	if(d.year.value.length == 0){
		d.year.value = " "
	}
	if(d.make.value.length == 0){
		d.make.value = " "
	}
	if(d.model.value.length == 0){
		d.model.value = " "
	}
	if(d.d1.value.length == 0){
		alert("Date is required")
		return false;
	}
	if(d.d2.value.length == 0){
		alert("Date is required")
		return false;
	}
	if(d.d3.value.length == 0){
		alert("Date is required")
		return false;
	}
	if(d.hours.value.length == 0){
		alert("Approximate Number of Hours is required")
		return false;
	}
	if(d.problem.value.length == 0){
		alert("Vehicle Issue is required")
		return false;
	}

	if(d.reminder.value == "email"){
		if(!validateEmail(d.email.value)){
			alert("If you want an email reminder to be sent, you must have a valid email address")
			return valse;
		}
		
		// make sure one of the boxes is checked
		oneday = document.getElementById("oneday").checked
		//sameday = document.getElementById("sameday").checked
		if(!oneday){
			alert("To send a reminder, you must check a Reminder Time")
			document.getElementById("remindertimes").style.backgroundColor = "yellow"
			return false
		}

	}
	if(d.reminder.value == "sms"){
		cphone = d.cac.value+d.cpf.value+d.clast.value
		if(cphone.length != 10){
			alert("To send a text message reminder, you must have a 10 digit Cell Phone number entered")
			return false;
		}
		
		// make sure one of the boxes is checked
		oneday = document.getElementById("oneday").checked
		if(!oneday){
			alert("To send a reminder, you must check a Reminder Time")
			document.getElementById("remindertimes").style.backgroundColor = "yellow"
			return false
		}
	}
	
	if(d.reminder.value == "both"){
		cphone = d.cac.value+d.cpf.value+d.clast.value
		email = d.email.value
		if(cphone.length != 10){
			alert("To send a text message reminder, you must have a 10 digit Cell Phone number entered")
			return false;
		}
		if(!validateEmail(email)){
			alert("If you want an email reminder to be sent, you must have a valid email address")
			return false;
		}
		// make sure one of the boxes is checked
		oneday = document.getElementById("oneday").checked
		if(!oneday){
			alert("To send a reminder, you must check a Reminder Time")
			document.getElementById("remindertimes").style.backgroundColor = "yellow"
			return false
		}
	}

	document.theform.submit()
}

function validateEmail(email) 
{
    var re = /\S+@\S+\.\S+/;
    return re.test(email);
}

function showReminderTimes(v){

	if (v == "No"){
		document.getElementById("remindertimes").style.display = "none"
	}else{
		document.getElementById("remindertimes").style.display = ""
	}

}
</script>
<style type="text/css">
.auto-style1 {
	font-size: small;
}
</style>
</head>

<body >

<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                      <td height="11" >
                      <p align="center">&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="11">
                      <form name=theform action=addapptnow.asp method=post>
                      <input name="ret" value="<%=request("ret")%>" type="hidden" >
                      <input type=hidden name=cid value="<%=request("custid")%>">
                      <input type=hidden name=origshopid value="<%=request("origshopid")%>">
                        <table border="0" width="100%" cellspacing="0" cellpadding="2">
                          <tr>
                            <td width="100%" valign="top" colspan="3" class="sbheader">
								Add New Appointment</td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Last 
							Name&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
                            <input value="<%if uf = "yes" then response.write rs("LastName") else response.write request("LastName")%>" type="text" name="lastname" size="16" class="namebox" tabindex="1"></td>
                            <td width="73%" rowspan="14" style="width: 36%" valign="top">
                <%
                ' get any recommended repairs from previous ro's
                ' get last ro where customer and vehicle, then look for recommends
                if len(request("custid")) > 0 and len(vid) > 0 then
                	if len(request("origshopid")) > 0 then
                		stmt = "select roid from repairorders where shopid = '" & request("origshopid") & "' and customerid = " & request("custid") & " and vehid = " & vid & " order by roid desc limit 1"
                	else
	                	stmt = "select roid from repairorders where shopid = '" & request.cookies("shopid") & "' and customerid = " & request("custid") & " and vehid = " & vid & " order by roid desc limit 1"
	                end if
	                set xrs = con.execute(stmt)
	                if not xrs.eof then
	                	roid = xrs("roid")
	                	
	                	' now get any recommendations for that ro
	                	stmt = "select * from recommend where shopid = '" & request.cookies("shopid") & "' and roid = " & xrs("roid")
	                	set xtrs = con.execute(stmt)
	                	if not xtrs.eof then
	                 %>
	                 <p style="font-size:medium;font-weight:bold;"> **** There 
					 are recommended repairs from a previous repair order. The 
					 recommendation(s) are listed below:<br><br>
	                 <%
	                 		do until xtrs.eof
	                 			response.write "<img alt='' src='newimages/rightarrows.png'> &nbsp;" & xtrs("desc") & "<br>"
	                 			xtrs.movenext
	                 		loop
	                 	end if
	                 %>
	                 <br>You will be prompted to add these recommendations when 
					 a Repair Order is created.
	                 <%
	                 end if
	             end if
                 %>

                            </td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>First 
							Name&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input value="<%if uf = "yes" then response.write rs("FirstName")%>" class="namebox" type="text" name="firstname" size="16" tabindex="2" ></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Home 
							Phone&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input onkeyup="javascript:if(document.theform.ac.value.length == 3){document.theform.pf.focus()};" class="datePhonebox" type="text" name="ac" size="5" value="<%if uf = "yes" then response.write ac%>" tabindex="3">
                              - 
							<input onkeyup="javascript:if(document.theform.pf.value.length == 3){document.theform.last.focus()};" class="datePhonebox" type="text" name="pf" size="5" value="<%if uf = "yes" then response.write pf%>" tabindex="4"> 
							- 
							<input onkeyup="javascript:if(document.theform.last.value.length == 4){document.theform.wac.focus()};" class="datePhonebox" type="text" name="last" size="5" value="<%if uf = "yes" then response.write last%>" tabindex="5"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><strong>
							Work Phone&nbsp;&nbsp;&nbsp; </strong></td>
                            <td style="width: 33%">
							<input onkeyup="javascript:if(document.theform.wac.value.length == 3){document.theform.wpf.focus()};" class="datePhonebox" type="text" name="wac" size="5" value="<%if uf = "yes" then response.write wac%>" tabindex="6">
                              - 
							<input onkeyup="javascript:if(document.theform.wpf.value.length == 3){document.theform.wlast.focus()};" class="datePhonebox" type="text" name="wpf" size="5" value="<%if uf = "yes" then response.write wpf%>" tabindex="7"> 
							- 
							<input onkeyup="javascript:if(document.theform.wlast.value.length == 4){document.theform.cac.focus()};" class="datePhonebox" type="text" name="wlast" size="5" value="<%if uf = "yes" then response.write wlast%>" tabindex="8"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><strong>
							Cell Phone&nbsp;&nbsp;&nbsp; </strong></td>
                            <td style="width: 33%">
							<input onkeyup="javascript:if(document.theform.cac.value.length == 3){document.theform.cpf.focus()};" class="datePhonebox" type="text" name="cac" size="5" value="<%if uf = "yes" then response.write cac%>" tabindex="9">
                              - 
							<input onkeyup="javascript:if(document.theform.cpf.value.length == 3){document.theform.clast.focus()};" class="datePhonebox" type="text" name="cpf" size="5" value="<%if uf = "yes" then response.write cpf%>" tabindex="10"> 
							- 
							<input onkeyup="javascript:if(document.theform.clast.value.length == 4){document.theform.year.focus()};" class="datePhonebox" type="text" name="clast" size="5" value="<%if uf = "yes" then response.write clast%>" tabindex="11"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><strong>
							Email Address&nbsp;&nbsp;&nbsp; </strong></td>
                            <td style="width: 33%">
							<input value="<%if uf = "yes" then response.write rs("EMail")%>" name="email" style="width: 286px" type="text" tabindex="12"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><strong>
							Send Appt. Reminder&nbsp;&nbsp;&nbsp; </strong></td>
                            <td style="width: 33%">
							<select name="reminder" tabindex="13">
							<option selected="" value="No">No</option>
							<option value="email">Email</option>
							<option value="sms">Text message</option>
							<option value="both">Send Both</option>
							</select> <strong><span class="auto-style1">
							(Reminder will be send the day before the 
							appointment)</span></strong></td>
                          </tr>
                          <tr style="display:none" id="remindertimes">
                            <td valign="top" align="right" style="width: 33%"><strong>
							Select Reminder Times</strong>&nbsp;&nbsp;&nbsp; </td>
                            <td style="width: 33%">
							<input name="oneday" id="oneday" checked type="checkbox">One 
							Day before appt<br></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Year&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input class="datePhonebox" type="text" name="year" size="6" value="<%if uf = "yes" then response.write yearf%>" tabindex="14"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Make&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input class="namebox" type="text" name="make" size="11" value="<%if uf = "yes" then response.write make%>" tabindex="15"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>
							Model&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input class="namebox" type="text" name="model" size="11" value="<%if uf = "yes" then response.write model%>" tabindex="16" ></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Date&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<input onkeyup="javascript:if(document.theform.d1.value.length == 2){document.theform.d2.focus()};" class="datePhonebox" type="text" name="d1" size="5" value="<%=mmonth%>" tabindex="17">
                              / 
							<input onkeyup="javascript:if(document.theform.d2.value.length == 2){document.theform.d3.focus()};" class="datePhonebox" type="text" name="d2" size="5" value="<%=dday%>" tabindex="18"> 
							/ 
							<input onkeyup="javascript:if(document.theform.d3.value.length == 2){document.theform.time.focus()};" class="datePhonebox" type="text" name="d3" size="5" value="<%=yyear%>" tabindex="19"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Time&nbsp;&nbsp;</b></td>
                            <td style="width: 33%">
							<select size="1" name="time" tabindex="20">
							<option value="none">Select Time</option>
<option value="5:45 AM">5:45 AM</option>						
<option value="6:00 AM">6:00 AM</option>
<option value="6:15 AM">6:15 AM</option>
<option value="6:30 AM">6:30 AM</option>
<option value="6:45 AM">6:45 AM</option>						
<option value="7:00 AM">7:00 AM</option>
<option value="7:15 AM">7:15 AM</option>
<option value="7:30 AM">7:30 AM</option>
<option value="7:45 AM">7:45 AM</option>
<option value="8:00 AM">8:00 AM</option>
<option value="8:15 AM">8:15 AM</option>
<option value="8:30 AM">8:30 AM</option>
<option value="8:45 AM">8:45 AM</option>
<option value="9:00 AM">9:00 AM</option>
<option value="9:15 AM">9:15 AM</option>
<option value="9:30 AM">9:30 AM</option>
<option value="9:45 AM">9:45 AM</option>
<option value="10:00 AM">10:00 AM</option>
<option value="10:15 AM">10:15 AM</option>
<option value="10:30 AM">10:30 AM</option>
<option value="10:45 AM">10:45 AM</option>
<option value="11:00 AM">11:00 AM</option>
<option value="11:15 AM">11:15 AM</option>
<option value="11:30 AM">11:30 AM</option>
<option value="11:45 AM">11:45 AM</option>
<option value="12:00 PM">12:00 PM</option>
<option value="12:15 PM">12:15 PM</option>
<option value="12:30 PM">12:30 PM</option>
<option value="12:45 PM">12:45 PM</option>
<option value="1:00 PM">1:00 PM</option>
<option value="1:15 PM">1:15 PM</option>
<option value="1:30 PM">1:30 PM</option>
<option value="1:45 PM">1:45 PM</option>
<option value="2:00 PM">2:00 PM</option>
<option value="2:15 PM">2:15 PM</option>
<option value="2:30 PM">2:30 PM</option>
<option value="2:45 PM">2:45 PM</option>
<option value="3:00 PM">3:00 PM</option>
<option value="3:15 PM">3:15 PM</option>
<option value="3:30 PM">3:30 PM</option>
<option value="3:45 PM">3:45 PM</option>
<option value="4:00 PM">4:00 PM</option>
<option value="4:15 PM">4:15 PM</option>
<option value="4:30 PM">4:30 PM</option>
<option value="4:45 PM">4:45 PM</option>
<option value="5:00 PM">5:00 PM</option>
<option value="5:15 PM">5:15 PM</option>
<option value="5:30 PM">5:30 PM</option>
<option value="5:45 PM">5:45 PM</option>
<option value="6:00 PM">6:00 PM</option>
<option value="6:15 PM">6:15 PM</option>
<option value="6:30 PM">6:30 PM</option>
<option value="6:45 PM">6:45 PM</option>
<option value="7:00 PM">7:00 PM</option>
<option value="7:15 PM">7:15 PM</option>
<option value="7:30 PM">7:30 PM</option>
<option value="7:45 PM">7:45 PM</option>
<option value="8:00 PM">8:00 PM</option>
                              </select></td>
                          </tr>
                          <%
                          stmt = "select * from colorcoding where shopid = '" & request.cookies("shopid") & "'"
                          set trs = con.execute(stmt)
                          if not trs.eof then
                          	fc = "#" & trs("colorhex")
                          %>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><strong>
							Appointment Category&nbsp;</strong></td>
                            <td style="width: 33%">
                            <select onchange="document.getElementById('colorshow').style.backgroundColor='#'+this.value" id="category" name="category" tabindex="21">
                            	<%
                            	do until trs.eof
                            		if trs("title") = request("c") then
                            			s = " selected='selected' "
                            		else
                            			s = ""
                            		end if
                            	%>
								<option <%=s%> value="<%=trs("colorhex")%>"><%=trs("title")%></option>
								<%
									trs.movenext
								loop
								%>
							</select>
							&nbsp; <span id="colorshow" style="width:100px;background-color:<%=fc%>;border:1px black solid">
							&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
                          </tr>
                          <%
                          end if
                          %>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>Approx 
							# of Hours&nbsp;&nbsp;</b></td>
                            <td colspan="2" style="width: 67%">
							<input class="datePhonebox" type="text" name="hours" size="10" tabindex="22" value="0"></td>
                          </tr>
                          <tr>
                            <td valign="top" align="right" style="width: 33%"><b>
							Vehicle Issue&nbsp;
                              </b></td>
                            <td colspan="2" style="width: 67%">
							<textarea name="problem" tabindex="23"></textarea></td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="3">
                              <p align="center"><input type="button" onclick="checkForm()" value="Add Appt." name="B1">
                              <input type="button" onclick="location.href='calendar.asp?ed=<%=request("d")%>'" value="Cancel" name="B2"></p>
                            </td>
                          </tr>
                        </table>
                      <input type="hidden" name="vid" value="<%=vid%>">
                      </form>
                      </td>
                    </tr>
                  </table>
                </div>
                 
                 
                 </p>
                <p align="center"><b><font size="4" color="#FF0000">All Fields 
				are required</font></b></td>
            </tr>
            <tr>
              <td valign="top" width="100%"></td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
<p></p>
<script type="text/javascript" language="javascript" src="jquery/jquery-1.10.2.min.js"></script>
<script language="javascript">
var timar = new Array('5:45 AM','6:00 AM','6:15 AM','6:30 AM','6:45 AM','7:00 AM','7:15 AM','7:30 AM','7:45 AM','8:00 AM','8:15 AM','8:30 AM','8:45 AM','9:00 AM','9:15 AM','9:30 AM','9:45 AM','10:00 AM','10:15 AM','10:30 AM','10:45 AM','11:00 AM','11:15 AM','11:30 AM','11:45 AM','12:00 PM','12:15 PM','12:30 PM','12:45 PM','1:00 PM','1:15 PM','1:30 PM','1:45 PM','2:00 PM','2:15 PM','2:30 PM','2:45 PM','3:00 PM','3:15 PM','3:30 PM','3:45 PM','4:00 PM','4:15 PM','4:30 PM','4:45 PM','5:00 PM','5:15 PM','5:30 PM','5:45 PM','6:00 PM','6:15 PM','6:30 PM','6:45 PM','7:00 PM')
for (j=0;j<=timar.length;j++){
	if(timar[j] == '<%=selecttime%>'){
		document.theform.time.selectedIndex = j
		//document.theform.time.value = j
		document.getElementById('category').onchange()
	}
}
//
document.theform.lastname.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>