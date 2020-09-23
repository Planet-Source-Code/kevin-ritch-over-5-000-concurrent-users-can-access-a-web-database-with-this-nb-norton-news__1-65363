<%
 SQLStr = HexToASCII(Request.Form("SQL"))
 HDSN = HexToASCII(Request.Form("HDSN")) & " "
 Tab = Chr(9)
 Result = ""
 Database = "BusinessCards.mdb"
 Set DataConn = Server.CreateObject("ADODB.Connection")
 DataConn.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("\db\" & Database)
 Set cmdTemp = Server.CreateObject("ADODB.Command")
 Set rstContacts = Server.CreateObject("ADODB.Recordset")
 cmdTemp.CommandText = SQLStr
 cmdTemp.CommandType = 1
 Set cmdTemp.ActiveConnection = DataConn
 rstContacts.Open cmdTemp, , 1, 3
 Found = rstContacts.RecordCount 
 If Found then
  rstContacts.Move CLng(Request("Record"))
  While Not rstContacts.EOF
   If rstContacts("DeleteFlag") = False then 
    PermitSend = True
    If rstContacts("PrivateRecord") = True then
     If TRIM(HDSN) <> TRIM(rstContacts("HardDiskSerialNumber")) Then
      PermitSend = False
     End If
    End If
    If PermitSend = True Then
     Result = Result & rstContacts("UID") & Tab 
     Result = Result & rstContacts("Company") & Tab 
     Result = Result & rstContacts("Person") & Tab 
     Result = Result & rstContacts("Title") & Tab 
     Result = Result & rstContacts("Phone") & Tab 
     Result = Result & rstContacts("Fax") & Tab 
     Result = Result & rstContacts("EMail") & Tab 
     Result = Result & rstContacts("Address1") & Tab 
     Result = Result & rstContacts("Address2") & Tab 
     Result = Result & rstContacts("City") & Tab 
     Result = Result & rstContacts("State") & Tab 
     Result = Result & rstContacts("Zip") & Tab 
     Result = Result & rstContacts("WebSite") & Tab 
     Result = Result & rstContacts("Created") & Tab 
     Result = Result & rstContacts("Mobile") & Tab 
     Result = Result & rstContacts("PrivateRecord") & Tab 
    End If
   End If
   Result = Result & vbCrLf
   rstContacts.MoveNext
  Wend
 End if
 Result = Result & "{EOF}"
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
