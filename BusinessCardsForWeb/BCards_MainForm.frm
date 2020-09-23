VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form MainForm 
   Caption         =   "A Database for Business Cards & Personal Contact Details : Stored on the Internet ~ by Kevin Ritch, May 2006"
   ClientHeight    =   8010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "BCards_MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "BCards_MainForm.frx":0442
   ScaleHeight     =   8010
   ScaleWidth      =   11865
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SystemMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3960
      ScaleHeight     =   1065
      ScaleWidth      =   4305
      TabIndex        =   24
      Top             =   1680
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Submitting Data Request to the Web Server"
         Height          =   195
         Left            =   840
         TabIndex        =   26
         Top             =   480
         Width           =   3120
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "BCards_MainForm.frx":72FD
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         Caption         =   "System Information"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   4320
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "BUSINESS CARDS"
      TabPicture(0)   =   "BCards_MainForm.frx":773F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MainWindow"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SET UP YOUR OWN PERSONAL WEB SERVER"
      TabPicture(1)   =   "BCards_MainForm.frx":775B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture3"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture3 
         Height          =   6975
         Left            =   -74880
         Picture         =   "BCards_MainForm.frx":7777
         ScaleHeight     =   6915
         ScaleWidth      =   11115
         TabIndex        =   27
         Top             =   480
         Width           =   11175
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "That's it! - Jolly Simple"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   6
            Left            =   1320
            TabIndex        =   64
            Top             =   5400
            Width           =   8355
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Need Help?  Feel free to contact me!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   555
            Index           =   5
            Left            =   1320
            TabIndex        =   63
            Top             =   6120
            Width           =   8355
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4. Change the variable Website$ in Form_Load"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   4560
            Width           =   10755
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3. Save the ASP Pages into the ""root"" folder"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   3600
            Width           =   10005
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2. Save the MDB into the /db/ folder (off ""root"")"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   2
            Left            =   120
            TabIndex        =   60
            Top             =   2640
            Width           =   10635
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1. Get a Domain Name && an ASP Hosted Site"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   1560
            Width           =   10365
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set up your own Web Server in 4 Easy Steps"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   0
            Left            =   480
            TabIndex        =   58
            Top             =   480
            Width           =   10200
         End
      End
      Begin VB.PictureBox MainWindow 
         Height          =   6975
         Left            =   120
         Picture         =   "BCards_MainForm.frx":E632
         ScaleHeight     =   6915
         ScaleWidth      =   11115
         TabIndex        =   1
         Top             =   480
         Width           =   11175
         Begin VB.PictureBox Picture2 
            Height          =   2775
            Left            =   120
            ScaleHeight     =   2715
            ScaleWidth      =   10845
            TabIndex        =   8
            Top             =   3240
            Width           =   10900
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   20
               Left            =   1320
               TabIndex        =   66
               Tag             =   "15~25"
               Top             =   1560
               Width           =   1935
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   6
               Left            =   1320
               TabIndex        =   23
               Tag             =   "13~13"
               Top             =   2280
               Width           =   4095
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   5
               Left            =   1320
               TabIndex        =   22
               Tag             =   "7~7"
               Top             =   1920
               Width           =   4095
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   4
               Left            =   3480
               TabIndex        =   21
               Tag             =   "6~6"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   3
               Left            =   1320
               TabIndex        =   20
               Tag             =   "5~5"
               Top             =   1200
               Width           =   1935
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   2
               Left            =   1320
               TabIndex        =   19
               Tag             =   "4~4"
               Top             =   840
               Width           =   2415
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   1
               Left            =   1320
               TabIndex        =   18
               Tag             =   "3~3"
               Top             =   480
               Width           =   4095
            End
            Begin VB.TextBox ContactField 
               Height          =   285
               Index           =   0
               Left            =   1320
               TabIndex        =   17
               Tag             =   "2~2"
               Top             =   120
               Width           =   4095
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   2535
               Left            =   5520
               TabIndex        =   28
               Top             =   120
               Width           =   5205
               _ExtentX        =   9181
               _ExtentY        =   4471
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "Business"
               TabPicture(0)   =   "BCards_MainForm.frx":154ED
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label1(7)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label1(8)"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Label1(9)"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Label1(10)"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "Label1(11)"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "Label1(12)"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "ContactField(7)"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "ContactField(8)"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "ContactField(9)"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "ContactField(10)"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "ContactField(11)"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "ContactField(12)"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "PrivateContact"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).ControlCount=   13
               TabCaption(1)   =   "Personal"
               TabPicture(1)   =   "BCards_MainForm.frx":15509
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Label1(13)"
               Tab(1).Control(1)=   "Label1(14)"
               Tab(1).Control(2)=   "Label1(15)"
               Tab(1).Control(3)=   "Label1(16)"
               Tab(1).Control(4)=   "Label1(17)"
               Tab(1).Control(5)=   "Label1(18)"
               Tab(1).Control(6)=   "Label1(19)"
               Tab(1).Control(7)=   "ContactField(13)"
               Tab(1).Control(8)=   "ContactField(14)"
               Tab(1).Control(9)=   "ContactField(15)"
               Tab(1).Control(10)=   "ContactField(16)"
               Tab(1).Control(11)=   "ContactField(17)"
               Tab(1).Control(12)=   "ContactField(18)"
               Tab(1).Control(13)=   "ContactField(19)"
               Tab(1).ControlCount=   14
               TabCaption(2)   =   "Notes"
               TabPicture(2)   =   "BCards_MainForm.frx":15525
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Command7"
               Tab(2).Control(1)=   "Command6"
               Tab(2).Control(2)=   "NListView"
               Tab(2).ControlCount=   3
               Begin VB.CheckBox PrivateContact 
                  Caption         =   "Private Contact"
                  Enabled         =   0   'False
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
                  Left            =   2940
                  TabIndex        =   67
                  Tag             =   "16~24"
                  Top             =   1920
                  Width           =   2055
               End
               Begin VB.CommandButton Command7 
                  Height          =   570
                  Left            =   -74160
                  Picture         =   "BCards_MainForm.frx":15541
                  Style           =   1  'Graphical
                  TabIndex        =   57
                  Top             =   1890
                  Width           =   615
               End
               Begin VB.CommandButton Command6 
                  Height          =   570
                  Left            =   -74880
                  Picture         =   "BCards_MainForm.frx":15983
                  Style           =   1  'Graphical
                  TabIndex        =   56
                  Top             =   1890
                  Width           =   615
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   19
                  Left            =   -71280
                  TabIndex        =   53
                  Tag             =   "22~22"
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   18
                  Left            =   -73920
                  TabIndex        =   51
                  Tag             =   "21~21"
                  Top             =   1920
                  Width           =   1335
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   17
                  Left            =   -71280
                  TabIndex        =   49
                  Tag             =   "20~20"
                  Top             =   1560
                  Width           =   1335
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   16
                  Left            =   -73920
                  TabIndex        =   47
                  Tag             =   "19~19"
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   15
                  Left            =   -73920
                  TabIndex        =   45
                  Tag             =   "18~18"
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   14
                  Left            =   -73920
                  TabIndex        =   43
                  Tag             =   "17~17"
                  Top             =   840
                  Width           =   3975
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   13
                  Left            =   -73920
                  TabIndex        =   41
                  Tag             =   "16~16"
                  Top             =   480
                  Width           =   3975
               End
               Begin VB.TextBox ContactField 
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   12
                  Left            =   1080
                  TabIndex        =   34
                  Tag             =   "14~14"
                  Top             =   1920
                  Width           =   1695
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   11
                  Left            =   3720
                  TabIndex        =   33
                  Tag             =   "12~12"
                  Top             =   1560
                  Width           =   1335
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   10
                  Left            =   1080
                  TabIndex        =   32
                  Tag             =   "11~11"
                  Top             =   1560
                  Width           =   2055
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   9
                  Left            =   1080
                  TabIndex        =   31
                  Tag             =   "10~10"
                  Top             =   1200
                  Width           =   2055
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   8
                  Left            =   1080
                  TabIndex        =   30
                  Tag             =   "9~9"
                  Top             =   840
                  Width           =   3975
               End
               Begin VB.TextBox ContactField 
                  Height          =   285
                  Index           =   7
                  Left            =   1080
                  TabIndex        =   29
                  Tag             =   "8~8"
                  Top             =   480
                  Width           =   3975
               End
               Begin MSComctlLib.ListView NListView 
                  Height          =   1365
                  Left            =   -74880
                  TabIndex        =   55
                  Top             =   480
                  Width           =   4965
                  _ExtentX        =   8758
                  _ExtentY        =   2408
                  SortKey         =   1
                  View            =   3
                  LabelEdit       =   1
                  SortOrder       =   -1  'True
                  Sorted          =   -1  'True
                  LabelWrap       =   -1  'True
                  HideSelection   =   0   'False
                  FullRowSelect   =   -1  'True
                  GridLines       =   -1  'True
                  _Version        =   393217
                  ForeColor       =   -2147483640
                  BackColor       =   -2147483643
                  BorderStyle     =   1
                  Appearance      =   1
                  NumItems        =   4
                  BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     Text            =   "NoteUID"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   1
                     Text            =   "TimeStamp"
                     Object.Width           =   0
                  EndProperty
                  BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   2
                     Text            =   "Time Stamp"
                     Object.Width           =   2928
                  EndProperty
                  BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                     SubItemIndex    =   3
                     Text            =   "Subject"
                     Object.Width           =   6121
                  EndProperty
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Spouse"
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
                  Index           =   19
                  Left            =   -72480
                  TabIndex        =   54
                  Top             =   1920
                  Width           =   810
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Phone"
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
                  Index           =   18
                  Left            =   -74880
                  TabIndex        =   52
                  Top             =   1920
                  Width           =   675
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Zip"
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
                  Index           =   17
                  Left            =   -71760
                  TabIndex        =   50
                  Top             =   1560
                  Width           =   345
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "State"
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
                  Index           =   16
                  Left            =   -74880
                  TabIndex        =   48
                  Top             =   1560
                  Width           =   555
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "City"
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
                  Index           =   15
                  Left            =   -74880
                  TabIndex        =   46
                  Top             =   1200
                  Width           =   405
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Addr2"
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
                  Index           =   14
                  Left            =   -74880
                  TabIndex        =   44
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  Index           =   13
                  Left            =   -74880
                  TabIndex        =   42
                  Top             =   480
                  Width           =   885
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Created"
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
                  Index           =   12
                  Left            =   120
                  TabIndex        =   40
                  Top             =   1920
                  Width           =   840
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Zip"
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
                  Index           =   11
                  Left            =   3240
                  TabIndex        =   39
                  Top             =   1560
                  Width           =   345
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "State"
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
                  Index           =   10
                  Left            =   120
                  TabIndex        =   38
                  Top             =   1560
                  Width           =   555
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "City"
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
                  Index           =   9
                  Left            =   120
                  TabIndex        =   37
                  Top             =   1200
                  Width           =   405
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Addr2"
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
                  Index           =   8
                  Left            =   120
                  TabIndex        =   36
                  Top             =   840
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address"
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
                  Index           =   7
                  Left            =   120
                  TabIndex        =   35
                  Top             =   480
                  Width           =   885
               End
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Website"
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
               Index           =   6
               Left            =   120
               TabIndex        =   15
               Top             =   2280
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "EMail"
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
               Index           =   5
               Left            =   120
               TabIndex        =   14
               Top             =   1920
               Width           =   600
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
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
               Index           =   4
               Left            =   120
               TabIndex        =   13
               Top             =   1560
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel && Fax"
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
               Index           =   3
               Left            =   120
               TabIndex        =   12
               Top             =   1200
               Width           =   1005
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title"
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
               Index           =   2
               Left            =   120
               TabIndex        =   11
               Top             =   840
               Width           =   480
            End
            Begin VB.Label Label1 
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
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   480
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
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   990
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   735
            Left            =   120
            ScaleHeight     =   675
            ScaleWidth      =   10845
            TabIndex        =   2
            Top             =   6120
            Width           =   10900
            Begin VB.CommandButton CreateEmptyLookupCommand 
               Height          =   690
               Left            =   9360
               Picture         =   "BCards_MainForm.frx":15DC5
               Style           =   1  'Graphical
               TabIndex        =   68
               ToolTipText     =   "CREATE AN EMPTY LOOKUP"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton Command8 
               Height          =   690
               Left            =   10130
               Picture         =   "BCards_MainForm.frx":16207
               Style           =   1  'Graphical
               TabIndex        =   65
               ToolTipText     =   "SYSTEM DATE AND TIME CONFIGURATION"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton SaveButton 
               Enabled         =   0   'False
               Height          =   690
               Left            =   1440
               Picture         =   "BCards_MainForm.frx":16511
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "SAVE CHANGES"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton Command4 
               Height          =   690
               Left            =   8640
               Picture         =   "BCards_MainForm.frx":16953
               Style           =   1  'Graphical
               TabIndex        =   7
               ToolTipText     =   "DELETE CONTACT"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton Command3 
               Height          =   690
               Left            =   2160
               Picture         =   "BCards_MainForm.frx":16C5D
               Style           =   1  'Graphical
               TabIndex        =   6
               ToolTipText     =   "ADD A NEW CONTACT"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton Command2 
               Height          =   690
               Left            =   720
               Picture         =   "BCards_MainForm.frx":1709F
               Style           =   1  'Graphical
               TabIndex        =   5
               ToolTipText     =   "VIEW THE CURRENTLY HIGHLIGHTED CONTACT IN THE LIST"
               Top             =   0
               Width           =   735
            End
            Begin VB.CommandButton Command1 
               Height          =   690
               Left            =   0
               Picture         =   "BCards_MainForm.frx":174E1
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "LOOKUP CONTACTS"
               Top             =   0
               Width           =   735
            End
         End
         Begin MSComctlLib.ListView CListView 
            Height          =   3015
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Width           =   10900
            _ExtentX        =   19235
            _ExtentY        =   5318
            SortKey         =   1
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   16
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "UID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Company"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Contact"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Title"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Phone"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Fax"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Email"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Address1"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Address2"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "City"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "State"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "Zip"
               Object.Width           =   1235
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "Website"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "Created"
               Object.Width           =   2910
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "Mobile"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "PrivateContact"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin VB.Menu LookupMenu 
      Caption         =   "&Lookup"
      Visible         =   0   'False
      Begin VB.Menu LookupCompany 
         Caption         =   "&Company"
      End
      Begin VB.Menu LookupContact 
         Caption         =   "Co&ntact"
      End
      Begin VB.Menu LookupCity 
         Caption         =   "Cit&y"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ListViewData$(255) ' (Only using 17 columns as it happens)
Dim ChangedCID$
Dim FieldNumberChanged$
Dim FieldDataChanged$
Dim FieldIndexChanged As Integer
Dim ClockFormat$
Dim GMTOffSet As Integer
Dim ShortDateFormat$
Dim Website$
Dim rstData$(255) ' Allows for up to 255 Access Fields
Dim CListItem As ListItem
Dim NListItem As ListItem
Dim myWidth As Integer
Dim MyHeight As Integer

Private Sub Form_Load()
 If App.PrevInstance Then
  MsgBox "The program is already running.", vbApplicationModal + vbExclamation, "Whoops!"
  End
 End If
'=======================================
'Set this to YOUR Website at EasyCGI.com
'=======================================
      Website$ = "ACTBrowser.com"
'=======================================
 On Error Resume Next
 MkDir "c:\Program Files"
 MkDir "c:\Program Files\V8Software"
 On Error GoTo 0
 WebPostForm.Show
 Call HardDiskSerial
 If Screen.Width > (Me.Width + 1000) Then
  Me.Top = (Screen.Height - Me.Height) / 2
  Me.Left = (Screen.Width - Me.Width) / 2
 Else
  Me.Top = 0
  Me.Left = 0
 End If
 myWidth = Me.Width
 MyHeight = Me.Height
'==========================================================================
'INITIALIZE THE WEBBROWSER OBJECT (WILL FAIL ON FIRST ATTEMPT WHEN LOADING)
'==========================================================================
 PostURL$ = "http://" & Website$ & "/GMTinHex.asp?Whatever=" & AsciiToHex$("HELLO")
 HTTPText$ = GetPostSource$(PostURL$)
 T = Timer + 1
 While T > Timer
  DoEvents
 Wend
'=======================
 Call GetMyTimeZone
'=======================
End Sub

Private Sub CListView_DBLClick()
'===================================================================
'Place data in Text Boxes (Uses .TAG property for Recordset Columns)
'===================================================================
 Rows = CListView.ListItems.Count
 If Rows < 1 Then
  Exit Sub
 End If
 Row = CListView.SelectedItem.Index
 For Col = 1 To 16
  If Col = 1 Then
   rstData$(Col) = CListView.ListItems(Row) ' UID
  Else
   rstData$(Col) = Trim$(CListView.SelectedItem.SubItems(Col - 1))
  End If
 Next Col
 Loading = True
 For TextBoxIndex = 0 To 12
  Col = Val(ContactField(TextBoxIndex).Tag)
  ContactField(TextBoxIndex).Text = rstData$(Col)
 Next TextBoxIndex
'================
'Private Contact?
'================
 PrivateContact.Enabled = UCase$(Trim$(rstData$(16))) = "TRUE"
 PrivateContact.Value = IIf(PrivateContact.Enabled, 1, 0)
'=======================================
'COLLECT BALANCE OF FIELDS AND THE NOTES
'=======================================
 SQL$ = "Select * From Contacts Where UID = " & rstData$(1)
 Call ProcessRequestForMoreContactData(SQL$)
'======================================
'Populate the balance of the Text Boxes
'======================================
 For TextBoxIndex = 13 To 20
  Col = Val(ContactField(TextBoxIndex).Tag)
  ContactField(TextBoxIndex).Text = rstData$(Col)
 Next TextBoxIndex
 Loading = False
End Sub

Private Sub CListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'=====================================
'Additional Sorting could be done here
'=====================================
End Sub

Private Sub CListView_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Call CListView_DBLClick
 End If
End Sub

Private Sub Command1_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 PopupMenu LookupMenu
End Sub

Private Sub Command2_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 Call CListView_DBLClick
 On Error Resume Next
 CListView.SetFocus
End Sub

Private Sub Command3_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 NewContact = False
 NewCompanyName$ = ""
 NewContactName$ = ""
 NewPrivateContact$ = "FALSE"
 NewContactForm.Top = Me.Top
 NewContactForm.Left = Me.Left + (Me.Width - NoteForm.Width)
 NewContactForm.Show vbModal, Me
 If NewContact = False Then
  Exit Sub
 End If
 Company$ = Trim$(NewCompanyName$)
 Contact$ = Trim$(NewContactName$)
 TheHDSN$ = IIf(NewSharedRecord, "SHARED!!!", SnStr)
 PostURL$ = "http://" & Website$ & "/AddNewContact.asp?Company=" & AsciiToHex$(Company$) & "&HDSN=" & AsciiToHex$(TheHDSN$) & "&PrivateContact=" & AsciiToHex$(NewPrivateContact$) & "&Contact=" & AsciiToHex$(Contact$)
 HTTPText$ = GetPostSource$(PostURL$)
 CID = Val(HTTPText$)
 SQL$ = "Select * From Contacts Where UID=" & Trim$(CID)
 Call ProcessLookupRequest(SQL$)
End Sub

Private Sub Command4_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 Rows = CListView.ListItems.Count
 If Rows = 0 Then
  Exit Sub
 End If
 If MsgBox("The CURRENT Contact is:" & String$(2, 10) & UCase$(ContactField(0).Text) & String$(2, 10) & UCase$(ContactField(1).Text) & String$(3, 10) & "Are you sure that you wish to DELETE this record?", vbApplicationModal + vbDefaultButton2 + vbExclamation + vbYesNo, "Deleted Contacts can only be restored by Admin!") = vbNo Then
  Exit Sub
 End If
'===========================
'FIND CURRENT ITEM TO DELETE
'===========================
 For RowCheck = 1 To Rows
  If Val(rstData$(1)) = Val(CListView.ListItems(RowCheck)) Then
   Row = RowCheck
   Exit For
  End If
 Next RowCheck
'===================================
'UPDATE THE DELETEFLAG = TRUE IN MDB
'===================================
 ChangedCID$ = Trim$(rstData$(1))
 PostURL$ = "http://" & Website$ & "/EditContactField.asp?HDSN=" & AsciiToHex$(SnStr) & "&FieldNumber=" & AsciiToHex("15") & "&Newdata=" & AsciiToHex("TRUE") & "&ContactUID=" & AsciiToHex(ChangedCID$)
 HTTPText$ = GetPostSource$(PostURL$)
 If InStr(HTTPText$, "PERMISSION DENIED - YOU DO NOT OWN THIS RECORD}") Then
  MsgBox "Sorry, I am unable to let you change or delete this record." & String$(2, 10) & "The system security protects records owned by other users.", vbApplicationModal + vbExclamation, "PERMISSION DENIED."
  GoTo AllDone:
 End If
 CListView.ListItems.Remove Row
 Rows = CListView.ListItems.Count
 If Rows = 0 Then
  Call CreateEmptyLookupCommand_Click
  Exit Sub
 Else
  CListView.SelectedItem = CListView.ListItems(1)
  Call CListView_DBLClick
  On Error Resume Next
  CListView.SetFocus
 End If
AllDone:
End Sub

Private Sub ContactField_KeyPress(Index As Integer, KeyAscii As Integer)
 If Val(rstData$(1)) = 0 Then
  KeyAscii = 0
  Exit Sub
 End If
End Sub

Private Sub ContactField_LostFocus(Index As Integer)
 If SaveButton.Enabled Then
  Call SaveButton_Click
 End If
End Sub

Private Sub CreateEmptyLookupCommand_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 Loading = True
 PrivateContact.Value = 0
 PrivateContact.Enabled = False
 For i = 0 To 20
  ContactField(i).Text = ""
 Next i
 For i = 1 To 255
  rstData$(i) = ""
  ListViewData$(i) = ""
 Next i
 CListView.ListItems.Clear
 NListView.ListItems.Clear
 Loading = False
End Sub

Private Sub SaveButton_Click()
 If Val(ChangedCID$) = 0 Then
  SaveButton.Enabled = False
  Exit Sub
 End If
 Loading = True
 If Len(Trim$(FieldDataChanged$)) < 2 Then ' (Contact simply called "ED" is allowed)
  If Val(FieldNumberChanged$) = 3 Then
   MsgBox "Sorry, you cannot simply remove the Contact name!", vbApplicationModal + vbExclamation, "Everybody has a name!"
   ContactField(1).Text = rstData$(3)
   Loading = False
   SaveButton.Enabled = False
   Exit Sub
  End If
 End If
'=================================
'UPDATE THE DATABASE ON THE SERVER
'=================================
 PostURL$ = "http://" & Website$ & "/EditContactField.asp?HDSN=" & AsciiToHex$(SnStr) & "&FieldNumber=" & AsciiToHex(FieldNumberChanged$) & "&Newdata=" & AsciiToHex(FieldDataChanged$) & "&ContactUID=" & AsciiToHex(ChangedCID$)
 HTTPText$ = GetPostSource$(PostURL$)
 If InStr(HTTPText$, "PERMISSION DENIED - YOU DO NOT OWN THIS RECORD}") Then
  MsgBox "Sorry, I am unable to let you change or delete this record." & String$(2, 10) & "The system security protects records owned by other users.", vbApplicationModal + vbExclamation, "PERMISSION DENIED."
  GoTo AllDone:
 End If
 Rows = CListView.ListItems.Count
 For RowCheck = 1 To Rows
  If Val(rstData$(1)) = Val(CListView.ListItems(RowCheck)) Then
   Row = RowCheck
   Exit For
  End If
 Next RowCheck
'=================================
'Store List View Data for a moment
'=================================
 CListView.SelectedItem = CListView.ListItems(Row)
 For Col = 1 To 16
  If Col = 1 Then
   ListViewData$(Col) = CListView.ListItems(Row) ' UID
  Else
   ListViewData$(Col) = Trim$(CListView.SelectedItem.SubItems(Col - 1))
  End If
 Next Col
'=============================
'Update Array with Change Made
'=============================
 rstData$(Val(FieldNumberChanged)) = FieldDataChanged$
 FINDEX = Val(ContactField(FieldIndexChanged).Tag)
 ListViewData$(FINDEX) = Trim$(FieldDataChanged$)
'=================================
'Remove the item from the ListView
'=================================
 CListView.ListItems.Remove Row
'===================
'Update the ListView
'===================
 For i = 1 To 16
  DAT$ = Trim$(ListViewData$(i))
  DAT$ = IIf(Trim$(DAT$) = "", DAT$ & " ", DAT$)
  If i = 1 Then
   Set CListItem = CListView.ListItems.Add(, , DAT$)
  Else
   CListItem.SubItems(i - 1) = DAT$
  End If
 Next i
'========================================================================
'Re-Locate the item in the ListView (Sort may have changed it's location)
'========================================================================
 Rows = CListView.ListItems.Count
 For RowCheck = 1 To Rows
  If Val(rstData$(1)) = Val(CListView.ListItems(RowCheck)) Then
   Row = RowCheck
   Exit For
  End If
 Next RowCheck
'===============================
'Reset the CURRENT Row Selection
'===============================
 CListView.SelectedItem = CListView.ListItems(Row)
AllDone:
 SaveButton.Enabled = False
 Loading = False
End Sub

Private Sub Command7_Click()
 Call NListView_DblClick
End Sub

Private Sub Command8_Click()
 If SaveButton.Enabled Then
  Exit Sub
 End If
 G$ = Chr$(9)
 TG$ = String$(2, 9)
 MsgBox "Clock Format: " & TG$ & ClockFormat$ & String$(2, 10) & "Short Date Format: " & G$ & ShortDateFormat$ & String$(2, 10) & "GMT OffSet: " & TG$ & GMTOffSet

End Sub

Private Sub ContactField_Change(Index As Integer)
 If Loading Then
  Exit Sub
 End If
 If Val(rstData$(1)) = 0 Then
  Exit Sub
 End If
 a$ = ContactField(Index).Tag
 S = InStr(a$, "~")
 FieldIndexChanged = Index
 FieldNumberChanged$ = Right$(a$, Len(a$) - S)
 FieldDataChanged$ = ContactField(Index).Text
 ChangedCID$ = Trim$(rstData$(1))
 SaveButton.Enabled = True
End Sub

Private Sub ContactField_GotFocus(Index As Integer)
 ContactField(Index).SelStart = 0
 ContactField(Index).SelLength = Len(ContactField(Index).Text)
End Sub

Private Sub Form_Resize()
 If Me.WindowState = 0 Then
  Me.Height = MyHeight
  Me.Width = myWidth
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 MsgBox "Cheers for using my program!" & String$(2, 10) & "Kevin Ritch" & String$(1, 10) & "http://GreatCRM.com", vbApplicationModal + vbInformation, "Have a groovy day now, y'hear!"
 End
End Sub

Private Sub LookupCity_Click()
 Fynd$ = UCase$(Trim$(InputBox$("Enter the CITY to search for:", "LOOKUP BY CITY")))
 If Fynd$ = "" Then
  Exit Sub
 End If
 Fynd$ = Replace$(Fynd$, "'", "''")
 SQL$ = "Select * From Contacts Where City Like '" & Fynd$ & "%'"
 Call ProcessLookupRequest(SQL$)
End Sub

Private Sub LookupCompany_Click()
 Fynd$ = UCase$(Trim$(InputBox$("Enter any part of the COMPANY NAME to search for:", "LOOKUP BY COMPANY")))
 If Fynd$ = "" Then
  Exit Sub
 End If
 Fynd$ = Replace$(Fynd$, "'", "''")
 SQL$ = "Select * From Contacts Where Company Like '%" & Fynd$ & "%'"
 Call ProcessLookupRequest(SQL$)
End Sub

Private Sub LookupContact_Click()
 Fynd$ = UCase$(Trim$(InputBox$("Enter any part of the CONTACT NAME to search for:", "LOOKUP BY CONTACT")))
 If Fynd$ = "" Then
  Exit Sub
 End If
 Fynd$ = Replace$(Fynd$, "'", "''")
 SQL$ = "Select * From Contacts Where Person Like '%" & Fynd$ & "%'"
 Call ProcessLookupRequest(SQL$)
End Sub

Sub ProcessLookupRequest(SQL$)
 SystemMessage.Visible = True
 SystemMessage.Refresh
 PostURL$ = "http://" & Website$ & "/DoSearch.asp?HDSN=" & AsciiToHex$(SnStr) & "&SQL=" & AsciiToHex(SQL$)
 HTTPText$ = GetPostSource$(PostURL$)
 If InStr(HTTPText$, Chr$(9)) Then
 '==================
 'There are records!
 '==================
  CListView.ListItems.Clear
 Else
  SystemMessage.Visible = False
  SystemMessage.Refresh
  MsgBox "No match found!", vbApplicationModal + vbInformation, "Search Completed!"
  Exit Sub
 End If
 DF = FreeFile
 Loading = True
 Open "c:\Program Files\V8Software\Temp.txt" For Output As #DF
 Print #DF, HTTPText$;
 Close #DF
 Open "c:\Program Files\V8Software\Temp.txt" For Input As #DF
 While Not EOF(DF)
  Line Input #DF, a$
  If InStr(a$, Chr$(9)) Then
   For i = 1 To 16
    S = InStr(a$, Chr$(9))
    DAT$ = Left$(a$, S - 1)
   '=========================
   'Fill the CONTACT ListView
   '=========================
    Select Case i
     Case 1 ' UID
      Set CListItem = CListView.ListItems.Add(, , DAT$ & " ")
     Case Else
      If i = 14 Then
      '==================
      'Format Create Date
      '==================
       DAT$ = PrettyDateFormat$(DAT$)
      End If
      CListItem.SubItems(i - 1) = DAT$ & " "
    End Select
    a$ = Right$(a$, Len(a$) - S)
    rstData$(i) = DAT$
   Next i
  End If
 Wend
 Close #DF
 REX = CListView.ListItems.Count
 If REX Then
  CListView.SelectedItem = CListView.ListItems(1)
  Call CListView_DBLClick
 End If
 SystemMessage.Visible = False
 SystemMessage.Refresh
 On Error Resume Next
 CListView.SetFocus
 Loading = False
End Sub
Sub ProcessRequestForMoreContactData(SQL$)
 SystemMessage.Visible = True
 SystemMessage.Refresh
 PostURL$ = "http://" & Website$ & "/GetOtherContactData.asp?SQL=" & AsciiToHex(SQL$)
 HTTPText$ = GetPostSource$(PostURL$)
 If InStr(HTTPText$, Chr$(9)) = 0 Then
  SystemMessage.Visible = False
  SystemMessage.Refresh
  MsgBox "Please try again.", vbApplicationModal + vbExclamation, "Apparent Internet Error!"
  Exit Sub
 End If
 DF = FreeFile
 Open "c:\Program Files\V8Software\Temp.txt" For Output As #DF
 Print #DF, HTTPText$;
 Close #DF
 Open "c:\Program Files\V8Software\Temp.txt" For Input As #DF
 Line Input #DF, a$
 If InStr(a$, Chr$(9)) Then
  For i = 16 To 22
   S = InStr(a$, Chr$(9))
   DAT$ = Left$(a$, S - 1)
  '================================================================================
  'Use Select Case Framework just in case special treatment of field data is wanted
  '================================================================================
   Select Case i
    Case Else
   End Select
   a$ = Right$(a$, Len(a$) - S)
   rstData$(i) = DAT$
  Next i
 End If
'==================================
'Now Collect Notes One-To-Many data
'==================================
 NListView.ListItems.Clear
 While Not EOF(DF)
  Line Input #1, a$
  If InStr(a$, Chr$(9)) Then
   S = InStr(a$, Chr$(9))
   NotesUID$ = Left$(a$, S - 1)
   a$ = Right$(a$, Len(a$) - S)
   S = InStr(a$, Chr$(9))
   Created$ = Left$(a$, S - 1)
   a$ = Right$(a$, Len(a$) - S)
   PrettyCreateDate$ = PrettyDateFormat$(Created$)
   S = InStr(a$, Chr$(9))
   Subject$ = Left$(a$, S - 1)
  '=======================
  'Fill the NOTES ListView
  '=======================
   Set NListItem = NListView.ListItems.Add(, , NotesUID$ & " ")
  '=======================================================================================
  'This is in order to SORT the Notes records in Date & Time Order (Reverse Chronological)
  '=======================================================================================
   NListItem.SubItems(1) = Created$
  '============================
  'Now insert the other columns
  '============================
   NListItem.SubItems(2) = PrettyCreateDate$
   NListItem.SubItems(3) = Subject$
  End If
 Wend
 Close #DF
 REX = NListView.ListItems.Count
 If REX Then
  NListView.SelectedItem = NListView.ListItems(1)
 End If
 SystemMessage.Visible = False
 SystemMessage.Refresh
 On Error Resume Next
 CListView.SetFocus
End Sub

Private Sub NListView_DblClick()
 Rows = NListView.ListItems.Count
 If Rows < 1 Then
  Exit Sub
 End If
 Row = NListView.SelectedItem.Index
 NoteUID$ = Trim$(Val(NListView.ListItems(Row)))
 TimeStamp$ = Trim$(NListView.SelectedItem.SubItems(2))
 Subject$ = Trim$(NListView.SelectedItem.SubItems(3))
'=============================
'COLLECT DETAILS FOR THIS NOTE
'=============================
 SQL$ = "Select * From Notes Where NoteUID = " & NoteUID$
 Detail$ = NoteDetail(SQL$)
 If InStr(Detail$, Chr$(9)) = 0 Then
  MsgBox "Please try again.", vbApplicationModal + vbExclamation, "Apparent Internet Error!"
  Exit Sub
 End If
 S = InStr(Detail$, Chr$(9))
 Detail$ = Left$(Detail$, S - 1)
'========================================
'Display the FULL NOTE in a Pop-up Window
'========================================
 NoteForm.Caption = "Note : " & TimeStamp$
 NoteForm.SubjectLabel = Subject$
 NoteForm.Text1.Text = Detail$
 NoteForm.Top = Me.Top
 NoteForm.Left = Me.Left + (Me.Width - NoteForm.Width)
 NoteForm.Show vbModal, Me
End Sub
Function NoteDetail(SQL$) As String
 SystemMessage.Visible = True
 SystemMessage.Refresh
 PostURL$ = "http://" & Website$ & "/GetNoteDetail.asp?SQL=" & AsciiToHex(SQL$)
 NoteDetail = GetPostSource$(PostURL$)
 SystemMessage.Visible = False
 SystemMessage.Refresh
End Function
Private Sub Command6_Click()
 Rows = CListView.ListItems.Count
 If Rows < 1 Then
  Exit Sub
 End If
 NewNote = False
 NewSubject$ = ""
 NewBody$ = ""
 NewNoteForm.Top = Me.Top
 NewNoteForm.Left = Me.Left + (Me.Width - NoteForm.Width)
 NewNoteForm.Show vbModal, Me
 If NewNote = False Then
  Exit Sub
 End If
 CID$ = Trim$(rstData$(1))
 Subject$ = Trim$(NewSubject$)
 Detail$ = Trim$(NewDetail$)
 PostURL$ = "http://" & Website$ & "/AddNewNote.asp?ContactUID=" & AsciiToHex(CID$) & "&Subject=" & AsciiToHex$(Subject$) & "&Detail=" & AsciiToHex$(Detail$)
 HTTPText$ = GetPostSource$(PostURL$)
 Call Command2_Click
End Sub
Sub GetShortDateFormat()
 ShortDateFormat$ = IIf(Month(CVDate("6/7/2008")) = 6, "MM/DD/YY", "DD/MM/YY")
 ClockFormat$ = IIf(ShortDateFormat$ = "DD/MM/YY", "HH:MM", "HH:MM AMPM") ' UK = 24 Hour Clock
End Sub
Sub GetMyTimeZone()
 PostURL$ = "http://" & Website$ & "/GMTinHex.asp?Whatever=" & AsciiToHex$("HELLO")
 GMTTime$ = GetPostSource$(PostURL$)
 If Len(GMTTime$) <> 12 Then
  MsgBox "Sorry, there appears to be an Internet Error.", vbApplicationModal + vbExclamation, "Please try again!"
  End
 End If
'==============================================================================
'ALL TIME-STAMPS IN ACCESS MDB ARE STORED IN ANSI FORMAT USING GMT TIME.  THUS,
'NO MATTER WHERE THE CLIENT IS, THE TIME IS STORED AS UNIVERSAL "WORLD-TIME" !!
'TIME-STAMPS ARE GENERATED ON THE IIS SERVER IN CALIFORNIA. TIME = TIME +8 HRS.
'==============================================================================
 Call GetShortDateFormat
 Select Case ShortDateFormat$
  Case "DD/MM/YY"
    GMTNow = CVDate(Mid$(GMTTime$, 7, 2) & "/" & Mid$(GMTTime$, 5, 2) & "/" & Mid$(GMTTime$, 3, 2) & " " & Mid$(GMTTime$, 9, 2) & ":" & Mid$(GMTTime$, 11, 2))
  Case "MM/DD/YY"
    GMTNow = CVDate(Mid$(GMTTime$, 5, 2) & "/" & Mid$(GMTTime$, 7, 2) & "/" & Mid$(GMTTime$, 3, 2) & " " & Mid$(GMTTime$, 9, 2) & ":" & Mid$(GMTTime$, 11, 2))
 End Select
 Difference = GMTNow - Now
 iMinutes = DateDiff("n", GMTNow, Now)
'=============================================================================
'Adjust for time being off, say, by 14 minutes between the server and local PC
'=============================================================================
 Neg = iMinutes < 0
 iMinutes = Abs(iMinutes) + 14
 iHours = Fix(iMinutes / 60)
 GMTOffSet = IIf(Neg, iHours * -1, iHours)
End Sub
Function PrettyDateFormat$(GMTTime$)
'============================================================
'ANSIString looks something like : "200605141234" (GMT Based)
'Converts the GMT ANSI Time-Stamp into a "user-friendly" time
'that will satiate the user's probable format expectation :-)
'============================================================
 Select Case ShortDateFormat$
  Case "DD/MM/YY"
    GMTNow = CVDate(Mid$(GMTTime$, 7, 2) & "/" & Mid$(GMTTime$, 5, 2) & "/" & Mid$(GMTTime$, 3, 2) & " " & Mid$(GMTTime$, 9, 2) & ":" & Mid$(GMTTime$, 11, 2))
  Case "MM/DD/YY"
    GMTNow = CVDate(Mid$(GMTTime$, 5, 2) & "/" & Mid$(GMTTime$, 7, 2) & "/" & Mid$(GMTTime$, 3, 2) & " " & Mid$(GMTTime$, 9, 2) & ":" & Mid$(GMTTime$, 11, 2))
 End Select
 ShowTime = DateAdd("h", GMTOffSet, GMTNow)
 PrettyDateFormat$ = Format$(ShowTime, ShortDateFormat$ & " " & ClockFormat$)
End Function
