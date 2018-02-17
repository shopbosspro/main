<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- #include file=aspscripts/conn.asp -->
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>Untitled 1</title>
<style type="text/css">
body{
	font-family:Arial, Helvetica, sans-serif
}
.style1 {
	color: #FFFFFF;
	font-weight: bold;
}
.style2 {
	background-image: url('newimages/wipheader.jpg');
}

.auto-style1 {
	text-align: left;
}
.auto-style2 {
	text-align: right;
}
.auto-style3 {
	text-align: center;
}

</style>
<script type="text/javascript">

function runReport(){

	// first make sure a category is selected and dates are selected
	sd = document.getElementById("sd").value
	ed = document.getElementById("ed").value
	cat = document.getElementById("cats").value

	if(sd.length == 0 || ed.length == 0 || cat == "select"){
		alert("You must select a category, start and end dates")
	}else{
		document.getElementById("auditlist").src = "activitydata.asp?sd="+sd+"&ed="+ed+"&cat="+cat
	}
}

function getROActivity(){

	ro = prompt("Enter the RO number")
	document.getElementById("auditlist").src = "activitydataro.asp?roid="+ro

}
</script>
</head>

<body>
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 51px">
                    <tr>
                      <td height="11" class="style2">
                        <p align="center" class="style1">Shop Boss Audit Log</p></td>
                    </tr>
                  </table>

<p align="center">Here you can view the log files for activity in Shop Boss Pro.&nbsp; 
Currently we log activities in Repair Orders such as adding and deleting parts, 
creating repair orders, etc.</p>
<p align="center">
<table style="width: 100%">
	<tr>
		<td class="auto-style2">Select Category: 
</td>
		<td class="auto-style1"> 
<select id="cats" name="cats">
<option value="select">Select Category</option>
<%
shopid = request.cookies("shopid")
stmt = "select distinct category from audit where shopid = '" & shopid & "' order by category"
set rs = con.execute(stmt)
do until rs.eof
%>
<option value="<%=rs("category")%>"><%=rs("category")%></option>
<%
	rs.movenext
loop
%>
</select></td>
	</tr>
	<tr>
		<td class="auto-style2">Start Date: </td>
		<td class="auto-style1"> <input type="text" id="sd" name="sd" value="" 
						style="width: 100px" /></td>
	</tr>
	<tr>
		<td class="auto-style2">End Date:</td>
		<td class="auto-style1"> <input type="text" id="ed" name="ed" value="" 
						style="width: 100px" size="20" /></td>
	</tr>
	<tr>
		<td class="auto-style3" colspan="2">
		<input name="Button1" type="button" onclick="runReport()" value="Run Report" />
		<input name="Button2" type="button" onclick="location.href='company.asp'" value="Done" />
		<br />
		<input name="Button3" type="button" onclick="getROActivity()" value="Get ALL Log Activity for Specific RO" /></td>
	</tr>
</table>
</p>
<iframe style="width:100%;min-height:500px;height:600px;" id="auditlist" name="auditlist">
				Your browser does not support inline frames or is currently configured not to display inline frames.
			</iframe>
<link href="jquery/jquery-ui.css" rel="stylesheet" />
<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.6.0/jquery.min.js"></script>
<script src="javascripts/jquery-ui.js"></script>

<script language="javascript">
$('#sd').datepicker();
$('#ed').datepicker();
$("#sd").attr('readonly', 'readonly');
$("#ed").attr('readonly', 'readonly');
</script>


</body>

</html>
