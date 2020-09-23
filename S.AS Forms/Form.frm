VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Caption         =   "Nick + Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   44
      Top             =   5040
      Width           =   3735
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   5
         TabIndex        =   46
         Text            =   "1000"
         Top             =   480
         Width           =   975
      End
      Begin VB.Timer Timer12 
         Left            =   1920
         Top             =   120
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Start The Number Machine"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   2400
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Nick + Time + Date Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   7920
      TabIndex        =   38
      Top             =   4440
      Width           =   3855
      Begin VB.Timer Timer11 
         Enabled         =   0   'False
         Left            =   3360
         Top             =   1320
      End
      Begin VB.Timer Timer10 
         Enabled         =   0   'False
         Left            =   3360
         Top             =   960
      End
      Begin VB.Timer Timer9 
         Left            =   3000
         Top             =   1440
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   43
         Text            =   "1000"
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Nick With Time + Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Nick With Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Nick With Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label8 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   360
         TabIndex        =   54
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "D . L Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   31
      Top             =   3120
      Width           =   3855
      Begin VB.CommandButton Command13 
         Caption         =   "Start D.L Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   37
         Top             =   240
         Width           =   2415
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Left            =   3120
         Top             =   720
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   36
         Text            =   "1000"
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Date Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7920
      TabIndex        =   30
      Top             =   1800
      Width           =   3855
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Left            =   2880
         Top             =   720
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   35
         Text            =   "1000"
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Start Date Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   34
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   240
         TabIndex        =   51
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Time Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7920
      TabIndex        =   29
      Top             =   0
      Width           =   3855
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Left            =   3120
         Top             =   1200
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   33
         Text            =   "1000"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Start Time Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   32
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   360
         TabIndex        =   50
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Nick Name Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4080
      TabIndex        =   5
      Top             =   6000
      Width           =   3735
      Begin VB.CommandButton Command10 
         Caption         =   "Speak My Nick Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   28
         Top             =   720
         Width           =   2535
      End
      Begin ACTIVEVOICEPROJECTLibCtl.DirectSS ad 
         Height          =   615
         Left            =   1440
         OleObjectBlob   =   "Form.frx":0000
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Nick Mover"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4080
      TabIndex        =   4
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   23
         Text            =   "3000"
         Top             =   2040
         Width           =   975
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Left            =   3240
         Top             =   960
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Left            =   3120
         Top             =   960
      End
      Begin VB.CommandButton r 
         Caption         =   "Right to Left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton l 
         Caption         =   "Left to Right"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "S . L Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4080
      TabIndex        =   3
      Top             =   2760
      Width           =   3735
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   1440
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "950"
         Top             =   1680
         Width           =   975
      End
      Begin VB.Timer Timer5 
         Left            =   3120
         Top             =   720
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Start The S . L Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label6 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Simple Nick Machine"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   5880
      Width           =   3975
      Begin VB.CommandButton Command8 
         Caption         =   "Show  Some Nick"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Change Nick"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   3735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Step Nick Machine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
      Begin VB.CommandButton Command6 
         Caption         =   "Clean"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   2640
         Width           =   735
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Left            =   360
         Top             =   2400
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "3000"
         Top             =   2640
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Start "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Load Form File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add To List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         ItemData        =   "Form.frx":0058
         Left            =   120
         List            =   "Form.frx":005A
         TabIndex        =   11
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label10 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   360
         TabIndex        =   56
         Top             =   2640
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Nick Machine"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Left            =   2880
         Top             =   2160
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         MaxLength       =   5
         TabIndex        =   9
         Text            =   "3000"
         Top             =   2160
         Width           =   975
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start The Machine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Timer Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2160
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog A 
      Left            =   6720
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "S.AE Nick Machine Made By Syed Adeel Hassan Rizvi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   7320
      Width           =   6495
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
      Begin VB.Menu line3 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents Msn As Messenger.MsgrObject
Attribute Msn.VB_VarHelpID = -1
Dim c As Integer
Dim x1 As String
Dim d
Private Sub About_Click()
frmAbout.Show
End Sub
Private Sub Command1_Click()
On Error Resume Next
If Command1.Caption = "Stop The Machine" Then
Timer1.Enabled = False
Command1.Caption = "Start The Machine"
Else
List1.Tag = 0
Timer1.Interval = Val(Text1.Text)
Timer1.Enabled = True
Command1.Caption = "Stop The Machine"
End If
End Sub
Private Sub Command10_Click()
ad.Speak Msn.Services.PrimaryService.FriendlyName
End Sub
Private Sub Command11_Click()
If Command11.Caption = "Stop The Machine" Then
Timer6.Enabled = False
Command11.Caption = "Start Time Machine"
Else
Timer6.Interval = Text9.Text
Timer6.Enabled = True
Command11.Caption = "Stop The Machine"
End If
End Sub

Private Sub Command12_Click()
If Command12.Caption = "Stop The Machine" Then
Timer7.Enabled = False
Command12.Caption = "Start Date Machine"
Else
Timer7.Interval = Text10.Text
Timer7.Enabled = True
Command12.Caption = "Stop The Machine"
End If
End Sub

Private Sub Command13_Click()
If Command13.Caption = "Stop The Machine" Then
Timer8.Enabled = False
Command13.Caption = "Start D . T Machine"
Else
Timer8.Interval = Text11.Text
Timer8.Enabled = True
Command13.Caption = "Stop The Machine"
End If
End Sub

Private Sub Command14_Click()
If Command14.Caption = "Stop" Then
Timer9.Enabled = False
Command14.Caption = "Nick With Date"
Else
Timer9.Interval = Text13.Text
Timer9.Enabled = True
Command14.Caption = "Stop"
End If
End Sub

Private Sub Command15_Click()
If Command15.Caption = "Stop" Then
Timer10.Enabled = False
Command15.Caption = "Nick With Time"
Else
Timer10.Interval = Text13.Text
Timer10.Enabled = True
Command15.Caption = "Stop"
End If
End Sub

Private Sub Command16_Click()
If Command16.Caption = "Stop" Then
Timer11.Enabled = False
Command16.Caption = "Nick With Time + Date"
Else
Timer11.Interval = Text13.Text
Timer11.Enabled = True
Command16.Caption = "Stop"
End If
End Sub

Private Sub Command17_Click()
Msn.Services.PrimaryService.FriendlyName = "0"
If Command17.Caption = "Stop  Machine" Then
Timer12.Enabled = False
Command17.Caption = "Start  Machine"
Else
Timer12.Interval = Text8.Item(1).Text
Timer12.Enabled = True
Command17.Caption = "Stop  Machine"
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
For i = 0 To Msn.List(MLIST_CONTACT).Count - 1
List1.AddItem Msn.List(MLIST_CONTACT).Item(i).FriendlyName
Next
End Sub
Private Sub Command3_Click()
On Error Resume Next
List2.AddItem Text2.Text
End Sub
Private Sub Command4_Click()
On Error Resume Next
Dim ok As Boolean
A.Filter = "Text File|*.txt"
A.ShowOpen
If A.FileName <> "" Then
Open A.FileName For Input As #1
While Not EOF(1)
Input #1, tt
List2.AddItem tt
Wend
End If
End Sub
Private Sub Command5_Click()
On Error Resume Next
If Command5.Caption = "Stop" Then
Timer2.Enabled = False
Command5.Caption = "Start"
Else
List2.Tag = 0
Timer2.Interval = Val(Text3.Text)
Timer2.Enabled = True
Command5.Caption = "Stop"
End If
End Sub
Private Sub Command6_Click()
List2.Clear
End Sub
Private Sub Command7_Click()
On Error Resume Next
If Text4.Text = "" Then
MsgBox "Please Type any Nick Name", vbCritical
End If
Msn.Services.PrimaryService.FriendlyName = Text4.Text
End Sub
Private Sub Command8_Click()
On Error Resume Next
Form2.Show
End Sub
Private Sub Command9_Click()
On Error Resume Next
If Command9.Caption = "Stop The S . L Machine" Then
Timer5.Enabled = False
Command9.Caption = "Start The S . L Machine"
Else
Timer5.Interval = Val(Text8.Item(0).Text)
c = 1
x1 = Text7.Text
Timer5.Enabled = True
Command9.Caption = "Stop The S . L Machine"
End If
End Sub
Private Sub Exit_Click()
End
End Sub
Private Sub Form_Load()
On Error Resume Next
origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
Set Msn = VBA.CreateObject("Messenger.MsgrObject")
Me.Hide
frmSplash.Show
Call main
DetectIE
mnuAddIE
Me.Caption = "S.AS Nick Machine - " & Msn.Services.PrimaryService.FriendlyName
End Sub
Private Sub Form_Resize()
On Error Resume Next

End Sub

Private Sub Help_Click()
On Error Resume Next
Shell "Notepad.EXE C:\Help.html", vbMaximizedFocus
End Sub
Private Sub l_Click()
On Error Resume Next
If l.Caption = "Stop The Mover" Then
Timer3.Enabled = False
l.Caption = "Left To Right"
Else
Timer3.Interval = Text6.Text
Timer3.Enabled = True
l.Caption = "Stop The Mover"
End If
End Sub
Private Sub Msn_OnLocalFriendlyNameChangeResult(ByVal hr As Long, ByVal pService As Messenger.IMsgrService, ByVal bstrPrevFriendlyName As String)
On Error Resume Next
sk.Caption = bstrPrevFriendlyName
Me.Caption = "S.AS Nick Machine - " & bstrPrevFriendlyName
End Sub
Private Sub Msn_OnLogoff()
MsgBox "The Program Is Closeing Please Reopen Program", vbCritical
Call main
End Sub
Private Sub Msn_OnLogonResult(ByVal hr As Long, ByVal pService As Messenger.IMsgrService)
MsgBox " Welcome " & Msn.Services.PrimaryService.FriendlyName
Call main
MsgBox "Now You Can able To Use S.AS Nick Machine", vbExclamation
End Sub
Private Sub Msn_OnServiceLogoff(ByVal hr As Long, ByVal pService As Messenger.IMsgrService)
Call main
End Sub
Private Sub Open_Click()
Me.WindowState = 2
Me.Show
End Sub
Private Sub r_Click()
On Error Resume Next
If r.Caption = "Stop The Mover" Then
Timer4.Enabled = False
r.Caption = "Right To Left"
Else
Timer4.Interval = Val(Text6.Text)
Timer4.Enabled = True
r.Caption = "Stop The Mover"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If List1.Tag > List1.ListCount - 1 Then List1.Tag = 0
Msn.Services.PrimaryService.FriendlyName = List1.List(Val(List1.Tag))
List1.ListIndex = Val(List1.Tag)
List1.Tag = Val(List1.Tag) + 1
End Sub

Private Sub Timer10_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Text12.Text & " - " & Time
End Sub

Private Sub Timer11_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Text12.Text & " - " & Time & " - " & Date
End Sub

Private Sub Timer12_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Msn.Services.PrimaryService.FriendlyName + 1
End Sub

Private Sub Timer13_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Msn.Services.PrimaryService.FriendlyName + Text14.Text & " - " & 1
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
If List2.Tag > List2.ListCount - 1 Then List2.Tag = 0
Msn.Services.PrimaryService.FriendlyName = List2.List(Val(List2.Tag))
List2.ListIndex = Val(List2.Tag)
List2.Tag = Val(List2.Tag) + 1
End Sub
Private Sub Timer3_Timer()
On Error Resume Next
If Timer3.Tag > Len(Text5.Text) - 1 Then Timer3.Tag = 0
Msn.Services.PrimaryService.FriendlyName = VBA.Right(Text5.Text, Len(Text5.Text) - Val(Timer3.Tag))
Timer3.Tag = Val(Timer3.Tag) + 1
End Sub
Private Sub Timer4_Timer()
On Error Resume Next
If Timer4.Tag > Len(Text5.Text) - 1 Then Timer4.Tag = 0
Msn.Services.PrimaryService.FriendlyName = VBA.Left(Text5.Text, Len(Text5.Text) - Val(Timer4.Tag))
Timer4.Tag = Val(Timer4.Tag) + 1
End Sub
Private Sub Timer5_Timer()
On Error Resume Next
If c > Len(x1) \ 2 Then c = 1
d = Split(x1, " ")
d(c) = UCase(d(c))
t = Join(d, " ")
Msn.Services.PrimaryService.FriendlyName = t
c = c + 1
Exit Sub
End Sub
Private Sub Timer6_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Time
End Sub
Private Sub Timer7_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Date
End Sub
Private Sub Timer8_Timer()
Msn.Services.PrimaryService.FriendlyName = Time & " - " & Date
End Sub
Private Sub Timer9_Timer()
On Error Resume Next
Msn.Services.PrimaryService.FriendlyName = Text12.Text & " - " & Date
End Sub
Private Sub try_MouseDblClick(Button As Integer, Id As Long)
On Error Resume Next
Me.WindowState = 2
Me.Show
End Sub
Private Sub try_MouseUp(Button As Integer, Id As Long)
On Error Resume Next
If Button = vbRightButton Then
Me.PopupMenu File
End If
End Sub
