<%
 SQLStr = HexToASCII(Request.Form("SQL"))
 Tab = Chr(9)
 Database = "BusinessCards.mdb"
 Set DataConn = Server.CreateObject("ADODB.Connection")
 DataConn.Open "Driver=Microsoft Access Driver (*.mdb);DBQ=" & Server.MapPath("\db\" & Database)
 Set cmdTemp = Server.CreateObject("ADODB.Command")
 Set rstContacts = Server.CreateObject("ADODB.Recordset")
 cmdTemp.CommandText = SQLStr
 cmdTemp.CommandType = 1
 Set cmdTemp.ActiveConnection = DataConn
 rstContacts.Open cmdTemp, , 1, 3
 Result = rstContacts("Detail") & Tab 
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
