VERSION 5.00
Begin VB.Form frmaboutme 
   BackColor       =   &H00000000&
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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   8040
      Picture         =   "frmaboutme.frx":0000
      Top             =   5760
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmaboutme.frx":030A
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3375
      Index           =   6
      Left            =   240
      TabIndex        =   1
      Top             =   5760
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHOR: Somdutt Ganguly"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   8040
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   2145
      Left            =   8760
      Picture         =   "frmaboutme.frx":053C
      Top             =   5760
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmaboutme.frx":4FA6
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   6375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   11535
   End
End
Attribute VB_Name = "frmaboutme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
