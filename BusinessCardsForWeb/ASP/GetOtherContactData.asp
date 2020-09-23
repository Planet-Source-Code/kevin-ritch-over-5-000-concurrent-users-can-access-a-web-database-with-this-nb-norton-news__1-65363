<%
 SQLStr = HexToASCII(Request.Form("SQL"))
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
 ContactID = rstContacts("UID")
 Result = Result & rstContacts("Home_Address1") & Tab 
 Result = Result & rstContacts("Home_Address2") & Tab 
 Result = Result & rstContacts("Home_City") & Tab 
 Result = Result & rstContacts("Home_State") & Tab 
 Result = Result & rstContacts("Home_Zip") & Tab 
 Result = Result & rstContacts("Home_Phone") & Tab 
 Result = Result & rstContacts("Spouse") & Tab 
 Result = Result & vbCrLf
 rstContacts.Close
'=======================
'GET THE CONTACT'S NOTES
'=======================
 cmdTemp.CommandText = "Select * From Notes Where ContactUID=" & ContactID
 cmdTemp.CommandType = 1
 Set cmdTemp.ActiveConnection = DataConn
 rstContacts.Open cmdTemp, , 1, 3
 Found = rstContacts.RecordCount 
 If Found then
  rstContacts.Move CLng(Request("Record"))
  While Not rstContacts.EOF
   If rstContacts("DeleteFlag") = False then 
    Result = Result & rstContacts("NoteUID") & Tab 
    Result = Result & rstContacts("Created") & Tab 
    Result = Result & rstContacts("Subject") & Tab 
   End If
   Result = Result & vbCrLf
   rstContacts.MoveNext
  Wend
 End if
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
