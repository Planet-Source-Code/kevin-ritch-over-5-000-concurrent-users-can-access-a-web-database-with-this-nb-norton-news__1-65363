VERSION 5.00
Begin VB.Form NewNoteForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Note"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8610
   Icon            =   "BCards_NewNoteForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BCards_NewNoteForm.frx":5D52
   ScaleHeight     =   3225
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   8415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "NewNoteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Trim$(Text2.Text) = "" Then
  MsgBox "Sorry, I need a SUBJECT.", vbApplicationModal + vbExclamation, "MISSING INFOMATION"
  Exit Sub
 End If
 If Trim$(Text1.Text) = "" Then
  MsgBox "Sorry, I need some DETAIL.", vbApplicationModal + vbExclamation, "MISSING INFOMATION"
  Exit Sub
 End If
 NewSubject$ = Trim$(Text2.Text)
 NewDetail$ = Trim$(Text1.Text)
 NewNote = True
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub
