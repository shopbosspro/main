<!-- #include file=aspscripts/auth.asp --> 
<html>
<!-- Copyright 2011 - Boss Software Inc. --><head><meta name="robots" content="noindex,nofollow">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title><%=request.cookies("shopname")%></title>
<script type="text/javascript" src="javascripts/killenter.js" ></script>
<style type="text/css">
<!--

P, TD, TH, LI { font-size: 10pt; font-family: Verdana, Arial, Helvetica }

.style1 {
	border-color: #000000;
	border-width: 0;
}
.style2 {
	font-weight: bold;

}
.style3 {

}

.style4 {
	text-align: center;
	color: #FFFFFF;
		background-image: url('newimages/wipheader.jpg');
}

.auto-style1 {
	text-align: left;
}

-->
</style>
<script type="text/javascript">
function checkForm(){
	
	if (document.theform.Category.value == ""){
		alert("You must enter a category")
		return false;
	}
	if (document.theform.Start.value == ""){
		alert("You must enter a Starting Part Cost")
		return false;
	}
	if (document.theform.End.value == ""){
		alert("You must enter an Ending Part Cost")
		return false;
	}
	if (document.theform.Factor.value == ""){
		alert("You must enter a Markup Factor")
		return false;
	}
	document.theform.submit()
	
}

function calcMargin(){

	e = document.theform.End.value
	f = document.theform.Factor.value
	// calc selling price from end value
	if (f.length > 0 && e >= 0){
		sp = e * f
		gp = sp - e
		if (gp >= 0){
			gpm = Math.round((gp / sp)*100)
			document.getElementById("gpm").innerHTML = gpm + "%"
		}
	}
	

}
</script>
</head>

<body  link="#800000" vlink="#800000" alink="#800000"  topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
    <tr>
      <td width="100%" valign="top">
        <div align="left">
          <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
            <tr>
              <td valign="top" width="100%">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="height: 50px">
                    <tr>
                      <td height="11" class="style4">
                       <strong>Add new category here.&nbsp;&nbsp; All Fields
                       are required. </strong>
                        </td>
                    </tr>
                  </table>
                </div>
               <form action=addmatrixnow.asp method=post name=theform>
               <div align="center">
                <center>
               <table width="60%" cellspacing="0" cellpadding="3" class="style1">
                <tr>
                 <td width="39%" align="right" class="style3"><b>Category</b> </td>
                 <td width="61%" class="style2"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="Category" size="20" style="font-variant: small-caps; font-family: Arial; font-size: 12pt;" value="<%=request("Category")%>"></td>
                </tr>
                <tr>
                 <td width="39%" align="right" class="style2">Start Part Cost</td>
                 <td width="61%" class="style2"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="Start" size="9" style="font-variant: small-caps; font-family: Arial; font-size: 12pt; " value="<%=request("Start")%>">(ie.
                  0.01)</td>
                </tr>
                <tr>
                 <td width="39%" align="right" class="style2">Ending Part 
				 Cost</td>
                 <td width="61%" class="style2"><input  onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="End" size="9" style="font-variant: small-caps; font-family: Arial; font-size: 12pt; " value="<%=request("End")%>">(ie.
                  2.99)</td>
                </tr>
                <tr>
                 <td width="39%" align="right" class="style2">Factor</td>
                 <td width="61%" class="style2"><input onkeyup="calcMargin()" onfocus=this.style.backgroundColor='yellow' onblur=this.style.backgroundColor='white' type="text" name="Factor" size="9" style="font-variant: small-caps; font-family: Arial; font-size: 12pt; " value="<%=request("Factor")%>">(ie.
                  3.75)</td>
                </tr>
                <tr>
                 <td width="39%" align="right" class="style2">Gross Profit 
				 Margin (calculated)</td>
                 <td id="gpm" width="61%" class="style2"></td>
                </tr>
                <tr>
                 <td width="100%" align="right" colspan="2">
                  <p align="center"><input onmouseover=this.style.cursor='hand' onclick="checkForm()" type="button" value="Add to Matrix" name="B1">
				  <input onclick="location.href='partpricematrix.asp'" name="Button1" type="button" value="Cancel"></td>
                </tr>
                <tr>
                 <td width="100%" colspan="2" class="auto-style1">
                  &nbsp;</td>
                </tr>
               </table>
                </center>
               </div>
               </form>
              </td>
            </tr>
            </table>
        </div>
      </td>
    </tr>
  </table>
  </center>
</div>
<p>The cost of a part will be multiplied by the 
	factor to arrive at your default selling price.&nbsp; You can also set a 
	Minimum Factor to prevent parts from being sold below a certain markup 
	without a password override.&nbsp; If minimum factor is 0 (zero), then parts 
	can be sold for any price without a password override.</p>
<script language="javascript">
document.theform.Category.focus()
</script>


</html>
<%
'Copyright 2011 - Boss Software Inc.
%>
