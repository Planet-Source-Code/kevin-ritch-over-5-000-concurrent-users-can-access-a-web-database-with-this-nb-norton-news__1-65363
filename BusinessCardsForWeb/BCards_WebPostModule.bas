Attribute VB_Name = "WebPostModule"
Global NewContact As Boolean
Global NewPrivateContact$
Global NewContactName$
Global NewCompanyName$
Global NewNote As Boolean
Global NewSubject$
Global NewDetail$
Global StringtoPost As String
Global SnStr As String
Global Loading As Boolean
Global NewSharedRecord As Boolean
Global SiteASP$
Global WaitForWeb As Boolean
Global WebResult$
 
Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer _
    As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
    lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function AsciiToHex(TypedData$) As String
 On Error Resume Next
 tmp$ = ""
 For i = 1 To Len(TypedData$)
  n = Asc(Mid$(TypedData$, i, 1))
  HV$ = "00" & Hex$(n)
  HV$ = Right$(HV$, 2)
  tmp$ = tmp$ & HV$
 Next i
 AsciiToHex$ = tmp$
End Function

Public Function HexToASCII(TheString) As String
 On Error Resume Next
 BYTES = Len(TheString)
 For i = 1 To BYTES
  HV = "&H" & Mid(TheString, i, 2)
  HB = CInt(HV)
  TmpStr = TmpStr & Chr(HB)
   i = i + 1
 Next
 HexToASCII = TmpStr
End Function
Public Function GetPostSource(TheURL As String) As String
 S = InStr(TheURL$, "?")
 If S = 0 Then
  Exit Function
 End If
 SiteASP$ = Left$(TheURL, S - 1)
 StringtoPost = Right$(TheURL, Len(TheURL) - S)
 WaitForWeb = True
 WebPostForm.Timer1.Enabled = True
 While WaitForWeb
  DoEvents
 Wend
 GetPostSource = WebResult$
End Function
Sub HardDiskSerial()
 Dim volname As String   ' receives volume name of C:
 Dim sn As Long          ' receives serial number of C:
 Dim maxcomplen As Long  ' receives maximum component length
 Dim sysflags As Long    ' receives file system flags
 Dim sysname As String   ' receives the file system name
 Dim retval As Long      ' return value
 volname = Space(256)
 sysname = Space(256)
 retval = GetVolumeInformation("C:\", volname, Len(volname), sn, maxcomplen, sysflags, sysname, Len(sysname))
 SnStr = Trim(Hex(sn))
 SnStr = String(8 - Len(SnStr), "0") & SnStr
 SnStr = Left(SnStr, 4) & "-" & Right(SnStr, 4)
End Sub
