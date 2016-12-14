VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form NewStream 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Stream Creator ;-)"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "NewStream.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CreateButton 
      Caption         =   "C&reate"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame F1 
      Caption         =   "Source Data "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   9015
      Begin MSComctlLib.ProgressBar P 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4800
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.TextBox NewStreamName 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1080
         MaxLength       =   48
         TabIndex        =   1
         Top             =   3120
         Width           =   6855
      End
      Begin VB.CommandButton BrowseButton 
         Caption         =   "&Browse"
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
         Left            =   7560
         TabIndex        =   0
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label SzLbl 
         Caption         =   "Size: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2160
         Width           =   5775
      End
      Begin VB.Label Label9 
         Caption         =   "Maximum Length: 48 Chars"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   16
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Review your options and click on Create when done."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   4320
         Width           =   7335
      End
      Begin VB.Label Label7 
         Caption         =   "\ / : * ? "" <> |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "A filename cannot contain:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Type a name for the new stream. or Click create to continue with the current name."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   2760
         Width           =   7335
      End
      Begin VB.Label Label3 
         Caption         =   "Browse for the file that you want to attach as a stream."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label Label5 
         Caption         =   "Step 2:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Step 3:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Step 1:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label OrigFilelbl 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   8535
      End
      Begin VB.Label StreamFilelbl 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   8535
      End
   End
   Begin VB.CommandButton Cancel 
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
      Left            =   7920
      TabIndex        =   4
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "NewStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stage As Long
Dim fileyes As Long
Dim fso As New FileSystemObject

Dim forb() As Byte 'forbidden characters..

Private Sub Cancel_Click()
Unload Me
Main.Show
End Sub

Private Sub CreateButton_Click()
formstat = 0
Dim boolval As Boolean
If Trim(NewStreamName.Text) = "" Then
    MsgBox "Type a custom name for the stream in the space provided.", vbCritical, "No Name"
    NewStreamName.SetFocus
    Exit Sub
End If

If fso.FileExists(StreamFilelbl.Caption) = False Then
    MsgBox "The selected file is invalid. Please browse again. The file may have been moved or renamed.", vbCritical, "File Not found"
    Exit Sub
End If

If fso.FileExists(OrigFilelbl.Caption) = False And fso.FolderExists(OrigFilelbl.Caption) = False Then
    MsgBox "The Original File used for the Stream Enumeration could not be located.", vbCritical, "File Not found"
    Unload Me
    Exit Sub
End If
boolval = Attacher(OrigFilelbl.Caption, StreamFilelbl.Caption, NewStreamName.Text)
If boolval = False Then
    MsgBox "The stream could not be attached.", vbCritical, "Fail!!"
    Exit Sub
End If
Main.Enabled = True
formstat = 1
Main.EnumerateBtn_Click
MsgBox "New stream successfully added.", vbExclamation, "Done!!"
Unload Me
End Sub

Private Sub Form_Load()
ReDim forb(9)
forb(0) = 92
forb(1) = 47
forb(2) = 58
forb(3) = 42
forb(4) = 63
forb(5) = 34
forb(6) = 60
forb(7) = 62
forb(8) = 124
Main.Enabled = False
OrigFilelbl.Caption = Main.StreamList.SelectedItem.Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.Enabled = True
Main.Show
End Sub

Private Sub BrowseButton_Click()
With Main
.FileDlg.Flags = FileOpenConstants.cdlOFNHideReadOnly
.FileDlg.FileName = ""
.FileDlg.ShowOpen
If .FileDlg.FileName = "" Then
    StreamFilelbl.Caption = ""
    SzLbl.Caption = "Size: Select a file."
    Exit Sub
End If
StreamFilelbl.Caption = .FileDlg.FileName
NewStreamName.Enabled = True
NewStreamName.SetFocus
SzLbl.Caption = "Size:  " & CStr(FileLen(StreamFilelbl.Caption)) & "   Bytes"
NewStreamName.Text = .FileDlg.FileTitle
NewStreamName.SelLength = Len(NewStreamName.Text)
End With
End Sub

Private Sub NewStreamName_Change()
For i = 0 To 8
If InStr(1, NewStreamName.Text, Chr(forb(i)), vbTextCompare) >= 1 Then
    MsgBox "Invalid character. Please retype.", vbCritical, "???"
    NewStreamName.Text = ""
    NewStreamName.SetFocus
Exit Sub
End If
Next
End Sub


