VERSION 5.00
Begin VB.Form frmalert 
   BackColor       =   &H00000000&
   Caption         =   "CYBER ALERT"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "frmalert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   4
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TIME EXTENSION"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EXTEND BY 1 HOURS"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "I DON'T WANT TO EXTEND"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   4335
      End
   End
   Begin VB.Label Label1 
      Caption         =   "ONLY 10 MINUTES IS REMAINING..."
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
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "frmalert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdok_Click()
If Option2.Value = True Then
frmclientnode.Winsock1.SendData ("[extension]")
Me.Visible = False
Else
Me.Visible = False
End If
End Sub

Private Sub Form_Deactivate()
On Error Resume Next
Me.SetFocus
End Sub


