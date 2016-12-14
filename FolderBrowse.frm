VERSION 5.00
Begin VB.Form BrowseFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse for Folder"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "FolderBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.DirListBox Dir 
      Height          =   2790
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.DriveListBox Drive 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label lbl 
      Caption         =   "Label1"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "BrowseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim File_name As String

Public Sub Command1_Click()
Main.Enabled = True
isObjFile = 2

Main.FileLbl.Caption = lbl.Caption
Main.EnumerateBtn.Enabled = True
Unload Me
End Sub

Public Sub Command2_Click()
Main.Enabled = True
isObjFile = 0
Main.FileLbl.Caption = ""
Main.EnumerateBtn.Enabled = False
Unload Me
End Sub

Public Sub Dir_Change()
lbl.Caption = ""
lbl.Caption = Dir.Path
End Sub

Public Sub Drive_Change()
On Error Resume Next
lbl.Caption = ""
lbl.Caption = Drive.Drive & "\"
Dir.Path = Drive.Drive
End Sub

Public Sub Form_Load()
SetWindowPos hwnd, conHwndTopmost, Screen.TwipsPerPixelX * 23, Screen.TwipsPerPixelY * 15, 310, 330, conSwpNoActivate Or conSwpShowWindow
Main.Enabled = False
lbl.Caption = ""
lbl.Caption = Dir.Path
End Sub

Public Sub Form_Unload(Cancel As Integer)
Main.Enabled = True
End Sub



