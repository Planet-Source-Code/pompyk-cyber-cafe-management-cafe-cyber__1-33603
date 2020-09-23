VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form cafecyber 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8880
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   11970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   67
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtnood 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   8160
      TabIndex        =   64
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "C"
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
      Left            =   7800
      TabIndex        =   60
      Top             =   3000
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin VB.Data nodestats1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "=================================="
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "nodestatus"
      Top             =   6000
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "CYBER CAFE/LAB STATUS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7215
      Left            =   480
      TabIndex        =   30
      Top             =   480
      Width           =   4215
      Begin VB.CommandButton cmdseatno 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   5
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   7
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   8
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   9
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   10
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   11
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   12
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   13
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "15"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   14
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   15
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "17"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   16
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "18"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   17
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "19"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   18
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   19
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "21"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   20
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "22"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   21
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "23"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   22
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5160
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "25"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   24
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "26"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   25
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "27"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   26
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6120
         Width           =   855
      End
      Begin VB.CommandButton cmdseatno 
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   27
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6120
         Width           =   855
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   4560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Data cyberpeopled 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cyberpeople"
      Top             =   6480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      DataField       =   "nodeno"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   28
      Top             =   2520
      Width           =   735
   End
   Begin VB.Data nodestats 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "=================================="
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "nodestatus"
      Top             =   4440
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "password"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "username"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdin 
      Caption         =   "IN"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton cmdout 
      Caption         =   "OUT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   12
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "person"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "timeout"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "timein"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   1560
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "hours"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      DataField       =   "name"
      DataSource      =   "nodestats"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   8640
      Width           =   11970
      _ExtentX        =   21114
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "cafecyber"
            TextSave        =   "cafecyber"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5186
            MinWidth        =   5186
            Text            =   "Author: Somdutt Ganguly"
            TextSave        =   "Author: Somdutt Ganguly"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5821
            MinWidth        =   5821
            Text            =   "gangulysomdutt@yahoo.com"
            TextSave        =   "gangulysomdutt@yahoo.com"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   186
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "SEARCH BY NODE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   65
      Top             =   5280
      Width           =   3255
      Begin VB.CommandButton cmdgoo 
         Caption         =   "GO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   66
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lblipaddress 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   69
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "IP ADDRESS:"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   480
      TabIndex        =   68
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   63
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   9960
      TabIndex        =   62
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      Caption         =   "SCROLL DOWN..PLZ"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9480
      TabIndex        =   61
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label lbltime 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   7680
      TabIndex        =   59
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "NODE NO:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   12
      Left            =   7920
      TabIndex        =   29
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   27
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "DAMAGED / NOT AVAILABLE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   11
      Left            =   5400
      TabIndex        =   26
      Top             =   7200
      Width           =   3255
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGGED IN"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   10
      Left            =   5400
      TabIndex        =   25
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "IN"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   9
      Left            =   5400
      TabIndex        =   24
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   22
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "VACANT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   8
      Left            =   5400
      TabIndex        =   21
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   6120
      Width           =   495
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "USER NAME:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   7
      Left            =   4920
      TabIndex        =   18
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   17
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "AMOUNT:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   5
      Left            =   4920
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "PERSON:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "TIME OUT:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "TIME IN:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "HOURS:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Menu mnufile 
      Caption         =   "FILE"
      Visible         =   0   'False
      Begin VB.Menu mnuexit 
         Caption         =   "EXIT"
      End
   End
   Begin VB.Menu mnunodestatus 
      Caption         =   "NODESTATUS"
      Visible         =   0   'False
      Begin VB.Menu mnudamaged 
         Caption         =   "DAMAGED"
      End
      Begin VB.Menu mnuin 
         Caption         =   "IN"
      End
      Begin VB.Menu mnuvacant 
         Caption         =   "VACANT"
      End
      Begin VB.Menu mnugo 
         Caption         =   "GO"
      End
      Begin VB.Menu mnuassignip 
         Caption         =   "ASSIGN IP"
      End
   End
End
Attribute VB_Name = "cafecyber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author : Somdutt Ganguly
'email  : gangulysomdutt@yahoo.com
'Application name: cafe cyber
'Type : cyber cafe/lab management
'Time : april (2002)
'about me: I am right now in my TY Bachelor of
'computer Application.. from CPICA college, Gujarat
'university, Ahmedabad, gujarat, INDIA
'address: no 6, chandrodaya apt,
'bhaikaka nagar, Thaltej
'ahmedabad, Gujarat, India.. 380059

'***********************************************************************

'*   About this sofware: This is a application software                *
'*   for cyber cafes....(cyber labs in colleges)..This                 *
'*   project is fully network based...which using                      *
'*   winsock control (socket programming) using UDP                    *
'*   protocol..I view this program as a handsome                       *
'*   application in the sense that it's a very good                    *
'*   example of network programming, it's easy, automates              *
'*   your cyber cafe, automatically assign random password             *
'*   , u can view all sorts of cyber cafe status...                    *
'*   , backup the data etc etc....and lastly use this s/w              *
'*   using 28 client nodes..........uses powerful system               *
'*   analysis of the cyber cafe......also in our city (ahmedabad)      *
'*   recently no of cyber cafes have increased rapidly so this         *
'*   can be used to take atleast some of their strain...               *

'*   I have explained more about this program in the help              *
'*   section in the application....................                    *
'*   Remember setting up this software requires network                *
'*   skill.............................                                *
'*   Do u like this s/w........                                        *
'*   if u like it...or...use it to learn from it then                  *
'*   offcourse vote for me..(give me feedback)..since                  *
'*   it really took enery to think about it.........                   *
'*   yes, u are free to use it!!.....if u use the source               *
'*   of this then don't replace my name...thx...thx...                 *
'*   since i know some of my friends doing so...this is                *
'*   not a good practice in programming..thx                           *

'***********************************************************************




Dim ipindex As Integer
Dim gotstring As String

Private Sub cmdcalculate_Click()
On Error GoTo errhand
x = InputBox("Enter amount per hour:")
Text6.Text = Val(Text2.Text) * x
'update it
Call cmdin_Click
Exit Sub
errhand:
MsgBox Err.Description
End Sub



Private Sub cmdexit_Click()
On Error Resume Next
Unload Me
Unload frmaboutme
Unload frmbackupdata
Unload frmdatabase
Unload frminstruction
Unload frmlookup
Unload frmsplashscreen
Unload cafecyber
End
End
End Sub

Private Sub cmdgoo_Click()
On Error GoTo errhand
nodestats.Recordset.FindFirst "nodeno=" & txtnood
Exit Sub
errhand:
MsgBox "not found"
End Sub

Private Sub cmdin_Click()
'if the pc is damaged..or not there...i.e. violet
'in color...don't let in
If cmdseatno(Val(Text9.Text) - 1).BackColor = &HFFC0C0 Then
Text1.Text = ""
Text2.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
MsgBox "can't go in..damaged pc.."
Exit Sub
End If
'ready for getting in
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(0) = Val(Text9.Text) Then
nodestats.Recordset.Edit
nodestats.Recordset.Update
MsgBox "record updated"
nodestats.Recordset.MoveFirst
Exit Sub
End If
nodestats.Recordset.MoveNext
Wend
End Sub

Private Sub cmdout_Click()
If Text1.Text = " " Or Text1.Text = "" Or Text2.Text = "0" Or Text2.Text = "" Then
MsgBox "..can't update empty field...sorry"
Exit Sub
End If
'for backup...of the data i.e
'people who have visited and left the cyber cafe
cyberpeopled.Recordset.AddNew
cyberpeopled.Recordset.Fields(0) = Text1.Text
cyberpeopled.Recordset.Fields(1) = Text2.Text
cyberpeopled.Recordset.Fields(2) = Text3.Text
cyberpeopled.Recordset.Fields(3) = Text4.Text
cyberpeopled.Recordset.Fields(4) = Text5.Text
cyberpeopled.Recordset.Fields(5) = Val(Text9.Text)
cyberpeopled.Recordset.Fields(6) = Val(Text6.Text)

cyberpeopled.Recordset.Update
cyberpeopled.Refresh
nodestats.Recordset.Edit

nodestats.Recordset.Fields(2) = " "
nodestats.Recordset.Fields(3) = 0
nodestats.Recordset.Fields(4) = " "
nodestats.Recordset.Fields(5) = " "
nodestats.Recordset.Fields(6) = 0
nodestats.Recordset.Fields(7) = " "
nodestats.Recordset.Fields(8) = " "
nodestats.Recordset.Fields(1) = "vacant"
nodestats.Recordset.Update
nodestats.Refresh
Call changecoloronout
MsgBox "user data backed up successfully"
End Sub

Private Sub cmdseatno_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
holdbuttonindex = Index
mnugo_Click
If Button = vbLeftButton Then
PopupMenu mnunodestatus
End If
End Sub




'create random password
Private Sub Command1_Click()
'random password creation algorithm...by me
Dim values As Integer
Dim generatedpassword As String
Dim i As Integer
For i = 0 To 6
Randomize
'create random no between 0 to 25
values = (Int(Rnd(1) * 25))
If values = 0 Then
generatedpassword = generatedpassword & "a"
ElseIf values = 1 Then
generatedpassword = generatedpassword & "b"
ElseIf values = 2 Then
generatedpassword = generatedpassword & "c"
ElseIf values = 3 Then
generatedpassword = generatedpassword & "d"
ElseIf values = 4 Then
generatedpassword = generatedpassword & "e"
ElseIf values = 5 Then
generatedpassword = generatedpassword & "f"
ElseIf values = 6 Then
generatedpassword = generatedpassword & "g"
ElseIf values = 7 Then
generatedpassword = generatedpassword & "h"
ElseIf values = 8 Then
generatedpassword = generatedpassword & "i"
ElseIf values = 9 Then
generatedpassword = generatedpassword & "j"
ElseIf values = 10 Then
generatedpassword = generatedpassword & "k"
ElseIf values = 11 Then
generatedpassword = generatedpassword & "l"
ElseIf values = 12 Then
generatedpassword = generatedpassword & "m"
ElseIf values = 13 Then
generatedpassword = generatedpassword & "n"
ElseIf values = 14 Then
generatedpassword = generatedpassword & "o"
ElseIf values = 15 Then
generatedpassword = generatedpassword & "p"
ElseIf values = 16 Then
generatedpassword = generatedpassword & "q"
ElseIf values = 17 Then
generatedpassword = generatedpassword & "r"
ElseIf values = 18 Then
generatedpassword = generatedpassword & "s"
ElseIf values = 19 Then
generatedpassword = generatedpassword & "t"
ElseIf values = 20 Then
generatedpassword = generatedpassword & "u"
ElseIf values = 21 Then
generatedpassword = generatedpassword & "v"
ElseIf values = 22 Then
generatedpassword = generatedpassword & "w"
ElseIf values = 23 Then
generatedpassword = generatedpassword & "x"
ElseIf values = 24 Then
generatedpassword = generatedpassword & "y"
ElseIf values = 25 Then
generatedpassword = generatedpassword & "z"
End If
Next i
Text8.Text = generatedpassword
generatedpassword = ""
End Sub

Private Sub Form_Load()
'Text10.Text = Date
cafecyber.nodestats.DatabaseName = App.Path & "\cyber cafe1.mdb"
cafecyber.nodestats1.DatabaseName = App.Path & "\cyber cafe1.mdb"
cafecyber.cyberpeopled.DatabaseName = App.Path & "\cyber cafe1.mdb"

nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(1) = "vacant" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &H8000000F
ElseIf nodestats.Recordset.Fields(1) = "in" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFF&
ElseIf nodestats.Recordset.Fields(1) = "logged" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFF00&
ElseIf nodestats.Recordset.Fields(1) = "damaged" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFFC0C0
End If
nodestats.Recordset.MoveNext
Wend
nodestats.Recordset.MoveFirst

'load winsock control
For i = 1 To 27
On Error Resume Next
Load Winsock1(i)
Next i
'assign ip address..ports etc
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(9) <> "" Then
'close it before being used
'Winsock1(nodestats.Recordset.Fields(0) - 1).Close
Winsock1(nodestats.Recordset.Fields(0) - 1).RemoteHost = nodestats.Recordset.Fields(9)
Winsock1(nodestats.Recordset.Fields(0) - 1).RemotePort = 11111

Winsock1(nodestats.Recordset.Fields(0) - 1).LocalPort = nodestats.Recordset.Fields(0)
End If
nodestats.Recordset.MoveNext
Wend
nodestats.Recordset.MoveFirst

'binding port and ip together..for udp protocol
For i = 0 To 27
Winsock1(i).Bind Winsock1(i).LocalPort
Next i
On Error Resume Next
lblipaddress.Caption = Winsock1(0).LocalIP


End Sub

Private Sub Form_Unload(Cancel As Integer)
For i = 1 To 27
Unload Winsock1(i)
Next i
End Sub







Private Sub mnuassignip_Click()
Dim ip As String
ip = InputBox("Assign an ip address to this node")
nodestats.Recordset.Edit
On Error GoTo x
nodestats.Recordset.Fields(9) = ip
nodestats.Recordset.Update
Exit Sub
x:
MsgBox Err.Description
End Sub

Private Sub mnudamaged_Click()
cmdseatno(holdbuttonindex).BackColor = &HFFC0C0
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(0) = holdbuttonindex + 1 Then
nodestats.Recordset.Edit
nodestats.Recordset.Fields(1) = "damaged"
nodestats.Recordset.Update

Exit Sub
End If
nodestats.Recordset.MoveNext
Wend

End Sub

Private Sub mnugo_Click()
nodestats.Recordset.FindFirst "nodeno=" & holdbuttonindex + 1
End Sub

Private Sub mnuin_Click()
cmdseatno(holdbuttonindex).BackColor = &HFF&
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(0) = holdbuttonindex + 1 Then
nodestats.Recordset.Edit
nodestats.Recordset.Fields(1) = "in"
nodestats.Recordset.Update
MsgBox "u need to fill in the details in the corresponding field...plz..otherwise..make it vacant.."
Exit Sub
End If
nodestats.Recordset.MoveNext
Wend
End Sub

Private Sub mnuvacant_Click()
cmdseatno(holdbuttonindex).BackColor = &HC0C0C0
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(0) = holdbuttonindex + 1 Then
nodestats.Recordset.Edit
nodestats.Recordset.Fields(1) = "vacant"
nodestats.Recordset.Update

Exit Sub
End If
nodestats.Recordset.MoveNext
Wend
End Sub


Private Sub Timer1_Timer()
lbltime.Caption = Now
StatusBar1.Panels(1).Text = Time
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Winsock1(i).GetData gotstring, vbString
ipindex = i
If gotstring = "[extension]" Then
Call extensionupdate
ElseIf gotstring = "[cybersurfended]" Then
cmdseatno(ipindex).BackColor = &HFF&
Call inupdate
Else
processdata gotstring
End If
End Sub


Sub processdata(data As String)
Dim username As String
Dim password As String
Dim nodeno As Integer
'extracting username and password
password = Mid(data, InStr(1, data, "::") + 2)
nodeno = Right(password, 1)
username = Left(data, InStr(1, data, "::") - 1)
'verifying the data from database
nodestats1.Refresh
nodestats1.Recordset.MoveFirst
While nodestats1.Recordset.EOF <> True
If nodestats1.Recordset.Fields(0) = nodeno Then
If nodestats1.Recordset.Fields(7) = username And nodestats1.Recordset.Fields(8) = Mid(password, 1, Len(password) - 1) Then

cmdseatno(ipindex).BackColor = &HFF00&
'update the database
Call loggedupdate
Exit Sub
End If
Winsock1(ipindex).SendData ("[invalid login]")
Exit Sub
End If
nodestats1.Recordset.MoveNext
Wend
Winsock1(ipindex).SendData ("[invalid login]")

End Sub


Sub loggedupdate()
nodestats1.Refresh
nodestats1.Recordset.MoveFirst
While nodestats1.Recordset.EOF <> True
If nodestats1.Recordset.Fields(0) = ipindex + 1 Then
nodestats1.Recordset.Edit
nodestats1.Recordset.Fields(1) = "logged"
nodestats1.Recordset.Fields(4) = Now
nodestats1.Recordset.Fields(5) = DateAdd("h", Trim(Text2.Text), Now)
nodestats1.Recordset.Update
Winsock1(ipindex).SendData ("[success]") & nodestats1.Recordset.Fields(5)
Exit Sub
End If
nodestats1.Recordset.MoveNext
Wend
nodestats.Refresh
End Sub

Sub extensionupdate()
nodestats1.Refresh
nodestats1.Recordset.MoveFirst
While nodestats1.Recordset.EOF <> True
If nodestats1.Recordset.Fields(0) = ipindex + 1 Then
nodestats1.Recordset.Edit
nodestats1.Recordset.Fields(3) = nodestats1.Recordset.Fields(3) + 1
nodestats1.Recordset.Fields(5) = DateAdd("h", Trim(1), nodestats1.Recordset.Fields(5))
nodestats1.Recordset.Update
'sending the extended time to the client..
Winsock1(ipindex).SendData "[extensiongranted]" & nodestats1.Recordset.Fields(5)
nodestats1.Refresh
Exit Sub
End If
nodestats1.Recordset.MoveNext
Wend
End Sub


Sub inupdate()
cmdseatno(ipindex).BackColor = &HFF&
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(0) = ipindex + 1 Then
nodestats.Recordset.Edit
nodestats.Recordset.Fields(1) = "in"
nodestats.Recordset.Update
MsgBox "Now u can encash the visitor...and out him"
Exit Sub
End If
nodestats.Recordset.MoveNext
Wend
End Sub

Sub changecoloronout()
nodestats.Refresh
nodestats.Recordset.MoveFirst
While nodestats.Recordset.EOF <> True
If nodestats.Recordset.Fields(1) = "vacant" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &H8000000F
ElseIf nodestats.Recordset.Fields(1) = "in" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFF&
ElseIf nodestats.Recordset.Fields(1) = "logged" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFF00&
ElseIf nodestats.Recordset.Fields(1) = "damaged" Then
cmdseatno(nodestats.Recordset.Fields(0) - 1).BackColor = &HFFC0C0
End If
nodestats.Recordset.MoveNext
Wend
nodestats.Recordset.MoveFirst
End Sub
