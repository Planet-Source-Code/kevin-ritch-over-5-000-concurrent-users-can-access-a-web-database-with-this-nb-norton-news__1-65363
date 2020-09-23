<%
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
 Response.Write DAYT    
%>
