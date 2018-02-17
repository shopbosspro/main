<!-- #include file=aspscripts/connwoshopid.asp -->
<%

Function Base64Decode(ByVal base64String)
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function


function validateInfo()

	' verify correct information and return success or failure
	phonevalidate = "no"
	addressvalidate = "no"
	stmt = "select customeraddress ca,customerphone ch,customerwork cw,cellphone cp from repairorders where shopid = '" & shopid & "' and roid = " & roid
	set rs = con.execute(stmt)
	if not rs.eof then
		
		' validate the phone number
		if right(rs("ch"),4) = l4 or right(rs("cw"),4) = l4 or right(rs("cp"),4) = l4 then
			phonevalidate = "yes"
		else
			phonevalidate = "no"
			validateInfo = "phone error"
			exit function
		end if

		'validate the address
		if lcase(rs("ca")) = a then
			addressvalidate = "yes"
		else
			validateInfo = "address error"
			addressvalidate = "no"
			exit function
		end if
		
		if phonevalidate = "yes" and addressvalidate = "yes" then
			validateInfo = "yes|" & rs("ch") & "," & rs("cw") & "," & rs("cp") & "|" & l4 & "|" & lcase(rs("ca")) & "|" & a
		else
			validateInfo = "no"
		end if
	else
		validateInfo = "recordset error"
	end if

end function

'on error resume next

shopid = base64decode(request("shopid"))
roid = base64decode(request("roid"))
l4 = request("ph")
a = lcase(request("a"))
sig = request("sig")

x = validateInfo()
'response.write "x:" & x & "<BR>"

if left(x,3) = "yes" then

	'on error resume next
	
	Function GetRandom()
	    
	    'get the number of seconds between 1999 and now
	    numsecs = cdbl(datediff("s","1/1/1999 00:00:00 AM",now))
	    
	    
	    GetRandom = numsecs
	    
	End Function
	
	set fso = server.createobject("scripting.filesystemobject")


	currdisplayedro = request("firstro")
	currarray = split(currdisplayedro,"pdfinvoices")
	
	currpath = server.mappath(".") & "\pdfinvoices" & replace(currarray(1),"/","\")
	'response.write currpath & "<BR>"

	Set Pdf = Server.CreateObject("Persits.Pdf")
	Set Doc = Pdf.OpenDocument(currpath)
	Set Font = Doc.Fonts("Helvetica-Bold")
	for each page in doc.pages
		set tfont = doc.fonts.loadfromfile("C:\Windows\Fonts\segoepr.ttf")
		page.canvas.drawtext request("sig") & " (e-signature on " & now & " MT)","x=70;y=47;height=18;width=500; size=12;color=black",tfont
	next
	
	shopuploadpath = "d:\savedinvoices\" & shopid
	if not fso.folderexists(shopuploadpath) then
		fso.createfolder shopuploadpath
		fso.createfolder shopuploadpath & "\" & roid
	end if
	
	rouploadpath = shopuploadpath & "\" & roid
	if not fso.folderexists(rouploadpath) then
		fso.createfolder shopuploadpath & "\" & roid
	end if
	
	newdirname = shopuploadpath & "\" & roid & "\"
	newfilename = shopid & "_" & roid & "_" & getrandom & ".pdf"
	
	doc.save(newdirname & newfilename)
	
	response.write "/savedinvoices/" & shopid & "/" & roid & "/" & newfilename
	
	
	'response.write "savedinvoices" & "/" & shopid & "/" & roid & "/" & newfilename
	
	recordAudit "Remote E-Signature|" & shopid & "|no", "RO E-signed by customer on RO#" & roid
	
	'send an email to the shop to alert them.
	stmt = "select companyname,companyemail,timezone from company where shopid = '" & shopid & "'"
	set rs = con.execute(stmt)
	shopemail = rs("companyemail")
	shopname = rs("companyname")
	tz = rs("timezone")
	
	set mail = server.createobject("persits.mailsender")
	mail.host = "smtp.emailsrvr.com"
	mail.addaddress shopemail
	mail.port = 587
	mail.from = "no-reply@shopbosspro.com"
	mail.username = "no-reply@shopbosspro.com"
	mail.password = "45ley92!"
	mail.fromname = "Shop Boss Pro"
	mail.subject = "E-Signature completed"
	mb = "An E-Signature was received at " & formatTimeClockTime(getCurrentLocalTime(tz)) & " for RO number " & roid & ".  To view the signed RO, please go into the RO and click on Print RO.  You will see it in the printed history" & chr(10) & chr(10)
	mail.body = mb
	mail.send

	
    if err.number <> 0 then
    	'response.write err.description
    end if
    
else
	response.write "validationerror"
end if

if err.number <> 0 then
	response.write "programerror"
	'response.write err.description
end if
%>
