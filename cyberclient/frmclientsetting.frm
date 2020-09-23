VERSION 5.00
Begin VB.Form frmclientsetting 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "SERVER SETTING"
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7380
   ControlBox      =   0   'False
   Icon            =   "frmclientsetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data serverinformad 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "serverinforma"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      Begin VB.TextBox txtserverport 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DataField       =   "nodeno"
         DataSource      =   "serverinformad"
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
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtsearverip 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DataField       =   "serverip"
         DataSource      =   "serverinformad"
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
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   2655
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
         Left            =   1080
         TabIndex        =   1
         Top             =   2400
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "REMOTE PORT:"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "SERVER IP:"
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
End
Attribute VB_Name = "frmclientsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsubmit_Click()
If txtsearverip.Text = "" Or txtserverport.Text = "" Then
MsgBox "fill in the details first"
Exit Sub
End If
On Error GoTo errhand
serverinformad.Refresh
serverinformad.Recordset.MoveFirst
serverinformad.Recordset.Edit
serverinformad.Recordset.Fields(0) = Trim(txtsearverip.Text)
serverinformad.Recordset.Fields(1) = Trim(txtserverport.Text)
serverinformad.Recordset.Update
MsgBox "records updated"
frmclientsetting.Visible = False

Exit Sub
errhand:
MsgBox Err.Description

End Sub

Private Sub Form_Deactivate()
On Error Resume Next
frmclientsetting.SetFocus
End Sub


Private Sub Form_Load()
On Error Resume Next
Me.serverinformad.DatabaseName = App.Path & "\cybercafeclient1.mdb"
End Sub
