VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "CAFE CYBER"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
cafecyber.Show
cafecyber.Top = 0
cafecyber.Left = 0
cafecyber.Height = 9000
cafecyber.Width = 12000
frmdatabase.Show
frmdatabase.Top = cafecyber.Height
frmdatabase.Left = 0
frmdatabase.Height = 9000
frmdatabase.Width = 12000
frmbackupdata.Show
frmbackupdata.Top = frmdatabase.Height + frmdatabase.Height
frmbackupdata.Left = 0
frmbackupdata.Height = 9000
frmbackupdata.Width = 12000
frmlookup.Show
frmlookup.Top = frmdatabase.Height + frmdatabase.Height + frmdatabase.Height
frmlookup.Left = 0
frmlookup.Height = 9000
frmlookup.Width = 12000
frminstruction.Show
frminstruction.Top = frmdatabase.Height + frmdatabase.Height + frmdatabase.Height + frmdatabase.Height
frminstruction.Left = 0
frminstruction.Height = 9000
frminstruction.Width = 12000
'about me form
frmaboutme.Show
frmaboutme.Top = frmdatabase.Height + frmdatabase.Height + frmdatabase.Height + frmdatabase.Height + frmdatabase.Height
frmaboutme.Left = 0
frmaboutme.Height = 9000
frmaboutme.Width = 12000
End Sub

'unloading stuff
Private Sub MDIForm_Unload(Cancel As Integer)
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
End Sub
