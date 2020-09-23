VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form WebPostForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WebPostForm"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   Icon            =   "BCards_WebPostForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   20
      Left            =   3360
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3360
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      Picture         =   "BCards_WebPostForm.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   1508
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "WebPostForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim bytpostdata() As Byte
 Dim strPostData As String
 Dim strHeader As String
 Dim varPostData As Variant
'====================================
'Pack the post data into a byte array
'====================================
 strPostData = StringtoPost
 BuildPostData bytpostdata(), strPostData
'=============================
'Write the byte into a variant
'=============================
 varPostData = bytpostdata
'=================
'Create the Header
'=================
 strHeader = "Content-Type: application/x-www-form-urlencoded" + Chr(10) + Chr(13)
'=============
'Post the data
'=============
On Error Resume Next
 WebBrowser1.Navigate2 SiteASP$, 0, "", varPostData, strHeader
End Sub
Private Sub Form_Load()
 WebBrowser1.Width = 1
 WebBrowser1.Height = 1
 WebBrowser1.Top = -20
 WebBrowser1.Left = -20
 Me.Width = 1
 Me.Height = 1
 Me.Left = -20
 Me.Top = -20
 Me.Visible = False
 Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
 Timer1.Enabled = False
 Call Command1_Click
End Sub

Private Sub Timer2_Timer()
 Timer2.Enabled = False
 Me.Visible = False
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 On Error Resume Next
 a$ = WebBrowser1.Document.Body.Innertext
 WebResult$ = HexToASCII(a$)
 If Trim$(a$) <> "" Then
  WaitForWeb = False
 End If
End Sub

Private Sub BuildPostData(ByteArray() As Byte, ByVal strPostData As String)
'========================================================
'NB: This Sub is available from a variety of VB websites.
'Please know that I am NOT claiming to have invented it.
'========================================================
 Dim intNewBytes As Integer
 Dim strCH As String
 Dim i As Integer
 intNewBytes = Len(strPostData) - 1
 If intNewBytes < 0 Then
  Exit Sub
 End If
 ReDim ByteArray(intNewBytes)
 For i = 0 To intNewBytes
  strCH = Mid$(strPostData, i + 1, 1)
  If strCH = Space(1) Then
   strCH = "+"
  End If
  ByteArray(i) = Asc(strCH)
 Next
End Sub
