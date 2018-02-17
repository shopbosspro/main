<%
'generate a random number to break the cache
intHighNumber = 10000
intLowNumber = 1

For i = 1 to 2
    Randomize
    intNumber = Int((intHighNumber - intLowNumber + 1) * Rnd + intLowNumber)
Next

'url = "http://www.vinquery.com/ws_POQCXTYNO1D/NX2PD7QK.aspx?accessCode=951a4321-8d16-44ef-a071-4728048a2275&vin=" & request("vin") & "&reportType=3&rand=" & intNumber
url = "http://ws.vinquery.com/restxml.aspx?accessCode=30bfea68-6f2f-4c47-9e4a-50928f5af5e1&vin=" & request("vin") & "&reportType=3&rand=" & intNumber
'Send the transaction info as part of the querystring
set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")
xml.open "GET", url, true
xml.send()

'Turn off error handling
On Error Resume Next

'Wait for up to 3 seconds if we've not gotten the data yet
If xml.readyState <> 4 then
	xml.waitForResponse 6
End If

'Did an error occur?  If so, use a default value for our data
If Err.Number <> 0 then
	strData = "some default text..."
Else
	If (xml.readyState <> 4) Or (xml.Status <> 200) Then
		'Abort the XMLHttp request
		xml.Abort
		strData = "Problem communicating with remote server..."
    Else
		strData = xml.ResponseText
    End If
End If

tar = split(strdata,chr(10))

for j = lbound(tar) to ubound(tar)
	linev = trim(tar(j))
	linev = replace(linev,"<Item Key=""","")
	linev = replace(linev," Unit="""" />","")
	linev = replace(linev,"""","")
	linev = replace(linev,"Value=",",")
	mar = split(linev,",")
	rvar = ""
	if instr(mar(0),"Year") then
		rvar = rvar & "year," & trim(mar(1))
	end if
	if instr(mar(0),"Make") then
		rvar = rvar & "*make," & trim(mar(1))
	end if
	if instr(mar(0),"Model") then
		if instr(mar(0),"Year") then
			'skip it
		else
			rvar = rvar & "*model," & trim(mar(1))
		end if
	end if
	if instr(mar(0),"Trim Level") then
		rvar = rvar & "*trim," & trim(mar(1))
	end if
	if instr(mar(0),"Body") then
		rvar = rvar & "*body," & trim(mar(1))
	end if
	if instr(mar(0),"Engine") then
		rvar = rvar & "*engine," & trim(mar(1))
	end if
	response.write rvar
next





%>