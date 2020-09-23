<%
 Company = HexToASCII(Request.Form("Company")) & " "
 Contact = HexToASCII(Request.Form("CONTACT")) & " "
 HDSN = HexToASCII(Request.Form("HDSN")) & " "
 PrivateRecord = HexToASCII(Request.Form("PrivateContact"))
 Tab = Chr(9)
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
 cmdTemp.CommandText = "Select * From Contacts Where UID=0"
 cmdTemp.CommandType = 1
 Set cmdTemp.ActiveConnection = DataConn
 rstContacts.Open cmdTemp, , 1, 3
 rstContacts.AddNew
 rstContacts("Company") = Company
 rstContacts("Person") = Contact
 rstContacts("Created") = DAYT
 If UCASE(TRIM(PrivateRecord)) = "TRUE" Then
  rstContacts("PrivateRecord") = True
 End If
 rstContacts("HardDiskSerialNumber") = TRIM(HDSN)
 rstContacts.Update
 Result = CStr(rstContacts("UID")) & Tab
 rstContacts.Close
 DataConn.Close
 Result = Result & "{EOF}"
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
