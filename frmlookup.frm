VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmlookup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdtotalamt 
      Caption         =   "TOTAL AMOUNT"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8280
      TabIndex        =   9
      Top             =   7440
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   8
      Top             =   8400
      Width           =   735
   End
   Begin MSComCtl2.DTPicker date11 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   0
      CalendarForeColor=   16777215
      Format          =   24576001
      CurrentDate     =   37260
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3480
      TabIndex        =   5
      Top             =   7920
      Width           =   735
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Left            =   2160
      TabIndex        =   3
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton cmdshowall 
      Caption         =   "SHOW ALL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   720
      TabIndex        =   2
      Top             =   7440
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=cyber cafe1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=cyber cafe1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from cyberpeople"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid datalookup 
      Bindings        =   "frmlookup.frx":0000
      Height          =   6015
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   0
      ForeColor       =   65280
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "BY DATE:"
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
      Left            =   720
      TabIndex        =   7
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Label lbllabels 
      BackColor       =   &H00808080&
      Caption         =   "BY NAME:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LOOKUP OF PREVIOUS VISITORS"
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
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "frmlookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdgo_Click()
On Error GoTo errhand
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
Adodc1.RecordSource = "select * from cyberpeople where name like '" & txtname.Text & "%'"
Adodc1.Refresh
If Adodc1.Recordset.Fields(0) = "" Or Adodc1.Recordset.Fields(0) = "" Then
GoTo errhand
End If

Exit Sub
errhand:
MsgBox "record not found"
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
End Sub



Private Sub cmdshowall_Click()
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
End Sub

Private Sub cmdtotalamt_Click()
On Error GoTo errhand
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
Adodc1.RecordSource = "select sum(amount) from cyberpeople"
Adodc1.Refresh
If Adodc1.Recordset.Fields(0) = "" Then
GoTo errhand
End If

Exit Sub
errhand:
MsgBox "record not found"
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
On Error GoTo errhand
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
Adodc1.RecordSource = "select * from cyberpeople where timein like '" & date11.Value & "%'"
Adodc1.Refresh
If Adodc1.Recordset.Fields(0) = "" Or Adodc1.Recordset.Fields(0) = "" Then
GoTo errhand
End If

Exit Sub
errhand:
MsgBox "record not found"
Adodc1.RecordSource = "select * from cyberpeople"
Adodc1.Refresh
End Sub


