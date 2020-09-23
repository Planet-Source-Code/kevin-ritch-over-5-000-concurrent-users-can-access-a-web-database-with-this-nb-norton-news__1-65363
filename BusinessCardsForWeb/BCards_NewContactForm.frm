VERSION 5.00
Begin VB.Form NewContactForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Contact"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "BCards_NewContactForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "BCards_NewContactForm.frx":5D52
   ScaleHeight     =   4065
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ViewableContact 
      BackColor       =   &H0066EAEA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      MaskColor       =   &H0057C8E3&
      TabIndex        =   12
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox SharedContact 
      BackColor       =   &H0066EAEA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      MaskColor       =   &H0057C8E3&
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.CheckBox PrivateContact 
      BackColor       =   &H0066EAEA&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      MaskColor       =   &H0057C8E3&
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
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
      Left            =   3720
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   3615
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
      Left            =   2280
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Viewable"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Anybody can VIEW this record."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1800
      TabIndex        =   13
      Top             =   1560
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For your eyes only."
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
      Left            =   1800
      TabIndex        =   11
      Top             =   1080
      Width           =   1995
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A SHARED Contact is one that is PUBLIC (not Private) and it may be edited or deleted by any other user."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Shared"
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
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Private"
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
      Left            =   240
      TabIndex        =   7
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact"
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
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
End
Attribute VB_Name = "NewContactForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 If Trim$(Text2.Text) = "" Then
  MsgBox "Sorry, I need a CONTACT NAME.", vbApplicationModal + vbExclamation, "MISSING INFOMATION"
  Exit Sub
 End If
 NewCompanyName$ = Trim$(Text1.Text)
 NewContactName$ = Trim$(Text2.Text)
 NewPrivateContact$ = IIf(PrivateContact, "TRUE", "FALSE")
 NewContact = True
 NewSharedRecord = SharedContact.Value = 1
 Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub PrivateContact_Click()
 If PrivateContact.Value = 0 Then
  If SharedContact.Value = 0 Then
   ViewableContact.Value = 1
  End If
 End If
 If PrivateContact.Value Then
  ViewableContact.Value = 0
 End If
 If SharedContact Then
  If PrivateContact.Value Then
   PrivateContact.Value = 0
   ViewableContact.Value = 0
   Beep
  End If
 End If
End Sub

Private Sub SharedContact_Click()
 If SharedContact.Value Then
  If PrivateContact.Value Then
   SharedContact.Value = 0
   MsgBox "Please uncheck 'PRIVATE' first if you wish to make this new contact a SHARED one!", vbApplicationModal + vbInformation, "You cannot SHARE a Private Contact"
  End If
 End If
End Sub

Private Sub ViewableContact_Click()
 If ViewableContact.Value = 1 Then
  PrivateContact.Value = 0
 End If
 
 If ViewableContact.Value = 0 Then
  SharedContact.Value = 0
  DoEvents
  PrivateContact.Value = 1
 End If
End Sub
