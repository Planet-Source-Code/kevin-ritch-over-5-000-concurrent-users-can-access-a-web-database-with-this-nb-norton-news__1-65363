<%
 CID = HexToASCII(Request.Form("ContactUID"))
 ContactUID = CLng(CID)
 FieldNumber = HexToASCII(Request.Form("FieldNumber"))
 FldNum = CLng(FieldNumber) -1
 NewData = HexToASCII(Request.Form("NewData")) & " "
 HDSN = HexToASCII(Request.Form("HDSN")) & " "
 HDSN = TRIM(HDSN)
 Tab = Chr(9)
 If FldNum = 13 Then ' Block Create Date Alteration
  FldNum = 255
  Response.Write "Hmmm - This should block that!"
 End If
 If FldNum = 23 Then ' Block Hard Disk Serial Number Changes! 
  FldNum = 255
  Response.Write "Hmmm - This should block that!"
 End If
'==============================================================
'This Server is in California, so Add 8 HOURS to convert to GMT
'==============================================================
 LeDate = DateAdd("h",8,Now)
 MyMin = DatePart("n",LeDate) : NN = "00" & trim(MyMin)
 MyHour = DatePart("h",LeDate) : HH = "00" & trim(MyHour)
 MyDay = DatePart("d",LeDate) : DD = "00" & trim(MyDay)
 MyMon = DatePart("m",LeDate) : MM = "00" & trim(MyMon)
 MyYear = DatePart("yyyy",LeDate) : YY = "0000" & trim(MyYear)
 DAYT = right(YY,4) & Right(MM,2) & right(DD,2) & Right(HH,2) & Right(NN,2)
 Database = "BusinessCards.mdb"
 Set DataConn = Server.CreateObject("ADODB.Connection")
 DataConn.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("\db\" & Database)
 Set cmdTemp = Server.CreateObject("ADODB.Command")
 Set rstContacts = Server.CreateObject("ADODB.Recordset")
 cmdTemp.CommandText = "Select * From Contacts Where UID=" & ContactUID
 cmdTemp.CommandType = 1
 Set cmdTemp.ActiveConnection = DataConn
 rstContacts.Open cmdTemp, , 1, 3
'====================
'Ownership Protection
'====================
 If TRIM(rstContacts("HardDiskSerialNumber")) = "SHARED!!!" Then
  HDSN = "SHARED!!!"
 End If
 If TRIM(rstContacts("HardDiskSerialNumber")) <> HDSN Then
  Result = "{PERMISSION DENIED - YOU DO NOT OWN THIS RECORD}"
 Else
  Select Case FldNum
   Case 14,23 ' DELETEFLAG & PRIVATERECORD
    If TRIM(UCASE(NewData)) = "TRUE" Then
     rstContacts(FldNum) = True
    Else
     rstContacts(FldNum) = False
    End If
   Case Else 
    rstContacts(FldNum) = Newdata
  End Select
  rstContacts.Update
  Result = "{UPDATED}"
 end If
 rstContacts.Close
 DataConn.Close
 Response.Write AsciiToHex(Result)    
%>
<%
 Function HexToAscii (TheString)
  tmp = ""
  Bytes = LEN(TheString)
  FOR i = 1 to Bytes
    HV =  "&H" & MID(TheString,i,2) 
    HB = CInt(HV)
    tmp = tmp & CHR(HB)
    i = i + 1
  NEXT
  HexToAscii = tmp
 End Function
%>
<%
 Function AsciiToHex(TheString)
  tmp = ""
  For i = 1 To Len(TheString)
   n = Asc(Mid(TheString, i, 1))
   HV = "00" & Hex(n)
   HV = Right(HV, 2)
   tmp = tmp & HV
  Next
  AsciiToHex = tmp
 End Function
%>
