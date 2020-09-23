VERSION 5.00
Begin VB.Form frminstruction 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frminstruction.frx":0000
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   5175
      Index           =   5
      Left            =   3360
      TabIndex        =   8
      Top             =   3720
      Width           =   8415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTS AND SERVER ARE UNIQUELY IDENTIFIED BY AN IP ADDRESS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   4
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Width           =   7815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EACH CLIENT HANDLES A SINGLE SOCKET USING UDP PROTOCOL WHICH IS BIND TO THE SERVER PORTS"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2295
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HANDLES MULTIPLE SOCKETS USING UDP PROTOCOL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1455
      Index           =   2
      Left            =   6720
      TabIndex        =   5
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER S/W"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   1
      Left            =   9480
      TabIndex        =   4
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENT S/W"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "SERVER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   2
      Left            =   9720
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NODES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "LAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   6
      Left            =   9840
      Top             =   1920
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   5
      Left            =   4200
      Top             =   2880
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   4
      Left            =   2400
      Top             =   2880
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   3
      Left            =   4200
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   2
      Left            =   2400
      Top             =   960
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   1
      Left            =   600
      Top             =   2880
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00004080&
      BackStyle       =   1  'Opaque
      Height          =   855
      Index           =   0
      Left            =   600
      Top             =   960
      Width           =   855
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   4680
      X2              =   4680
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2880
      X2              =   2880
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   960
      X2              =   960
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   960
      X2              =   9720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   6
      Left            =   9720
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   5
      Left            =   4080
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   4
      Left            =   4080
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   3
      Left            =   2280
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   2
      Left            =   2280
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   1
      Left            =   480
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1335
      Index           =   0
      Left            =   480
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "frminstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim i As Integer
For i = 0 To Shape1.UBound
Shape1(i).BorderColor = vbRed
Next i

For i = 0 To Line1.UBound
Line1(i).BorderColor = vbWhite
Next i
End Sub

Private Sub Timer2_Timer()
Dim i As Integer
For i = 0 To Shape1.UBound
Shape1(i).BorderColor = vbWhite
Next i

For i = 0 To Line1.UBound
Line1(i).BorderColor = vbRed
Next i
End Sub
