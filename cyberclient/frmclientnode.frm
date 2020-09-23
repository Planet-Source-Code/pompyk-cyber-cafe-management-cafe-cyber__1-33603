VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclientnode 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsettings 
      Caption         =   "SETTING"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1935
   End
   Begin VB.Data serverinformad 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "serverinforma"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer changetime 
      Interval        =   1000
      Left            =   4680
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.TextBox txtpname 
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
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtuname 
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
      Height          =   495
      Left            =   5640
      TabIndex        =   0
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   5400
      TabIndex        =   4
      Top             =   2400
      Width           =   3615
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton cmdsubmit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Label lblcloseme 
      BackColor       =   &H00808080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11520
      TabIndex        =   17
      Top             =   8520
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LOCAL IP:"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   16
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblyourip 
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
      ForeColor       =   &H00C0C0FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME LEFT:"
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
      Left            =   5640
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT TIME:"
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
      Left            =   5640
      TabIndex        =   13
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LOGOUT TIME:"
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
      Index           =   0
      Left            =   5640
      TabIndex        =   12
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lbltimeleft 
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
      Left            =   7560
      TabIndex        =   11
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label lblserverip 
      BackStyle       =   0  'Transparent
      DataSource      =   "serverinformad"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label4 
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
      Left            =   7560
      TabIndex        =   8
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label3 
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
      Left            =   7560
      TabIndex        =   7
      Top             =   0
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00E0E0E0&
      Height          =   6615
      Left            =   1080
      Top             =   840
      Width           =   9375
   End
   Begin VB.Label Label2 
      Caption         =   "PASSWORD:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "USER NAME:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
   End
End
Attribute VB_Name = "frmclientnode"
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


Private Sub changetime_Timer()
Dim hh As Integer
Dim mm As Integer
Dim ss As Integer
Dim sss As Long
On Error Resume Next
'current time
Label4.Caption = Now
'conversion in hh:mm:ss...
sss = DateDiff("s", CDate(Label4.Caption), CDate(Label3.Caption))
hh = Int(sss / 3600)
mm = Int((sss Mod 3600) / 60)
ss = Int((sss Mod 60) Mod 60)
lbltimeleft.Caption = hh & ":" & mm & ":" & ss
'when 10 minutes 10 seconds remain show the dialog..for extension

If lbltimeleft.Caption = "1:59:45" Then
frmalert.Visible = True
End If

If Label3.Caption = Label4.Caption Then
Winsock1.SendData "[cybersurfended]"
Label3.Caption = ""
For i = 0 To 8130
Me.Height = Me.Height + 1
Next i
End If
End Sub

Private Sub cmdclear_Click()
txtuname.Text = ""
txtpname.Text = ""
End Sub

Private Sub cmdsettings_Click()
frmclientsetting.Visible = True
End Sub

Private Sub cmdsubmit_Click()
On Error GoTo errorhand
If txtuname.Text = "" Or txtpname = "" Then
MsgBox "fill in the fields plz"
Exit Sub
End If
Dim connectiontext As String
Call Form_Load
DoEvents
connectiontext = ""
connectiontext = Trim(txtuname.Text)
connectiontext = connectiontext & "::" & txtpname.Text & Winsock1.RemotePort
Winsock1.SendData (connectiontext)
Exit Sub
errorhand:
MsgBox Err.Description
End Sub



Private Sub Form_Load()
'retrieve the server's ip address in the cyber
'cafe
frmclientnode.serverinformad.DatabaseName = App.Path & "\cybercafeclient1.mdb"

serverinformad.Refresh
On Error Resume Next
serverinformad.Recordset.MoveFirst
lblserverip.Caption = serverinformad.Recordset.Fields(0)
Winsock1.Close
Winsock1.RemoteHost = serverinformad.Recordset.Fields(0)
Winsock1.LocalPort = 11111
Winsock1.RemotePort = serverinformad.Recordset.Fields(1)
Winsock1.Bind Winsock1.LocalPort
lblyourip.Caption = Winsock1.LocalIP
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
Unload frmalert
Unload frmclientsetting
End
End Sub

'end the program..
Private Sub lblcloseme_Click()
On Error Resume Next
Unload Me
Unload frmalert
Unload frmclientsetting
End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim x As String
Dim i As Integer
Winsock1.GetData x, vbString
If Left(x, 9) = "[success]" Then
txtuname.Text = ""
txtpname.Text = ""
For i = 0 To 8130
Me.Height = Me.Height - 1
Next i
Label3.Caption = Mid(x, 10)
ElseIf Left(x, 18) = "[extensiongranted]" Then
Label3.Caption = Mid(x, 19)
End If
End Sub


