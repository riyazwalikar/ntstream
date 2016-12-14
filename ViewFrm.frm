VERSION 5.00
Begin VB.Form ViewFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stream Viewer"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   Icon            =   "ViewFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SaveButton 
      Caption         =   "&Save"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton ExportButton 
      Caption         =   "E&xport"
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
      Left            =   9000
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton DeleteButton 
      Caption         =   "&Delete"
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
      Left            =   7440
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton EditButton 
      Caption         =   "&Edit"
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
      Left            =   5880
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Warning!! This Viewer is meant for text files only. If you have opened a non-text file then please edit at your own risk."
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
      Left            =   120
      TabIndex        =   5
      Top             =   5520
      Width           =   10095
   End
   Begin VB.Label lblsz 
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
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   2895
   End
End
Attribute VB_Name = "ViewFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim editclicked As Integer

Private Sub DeleteButton_Click()
If MsgBox("Do you really want to delete this stream?", vbInformation + vbYesNo, "Sure?") = vbYes Then
    DeleteFile Me.Caption
    Unload Me
End If
'MsgBox "Delete functionality disabled in the Trial Version.", vbInformation, "Disabled"
End Sub

Private Sub EditButton_Click()
Text1.Locked = False
Text1.SetFocus
editclicked = 1
SaveButton.Visible = True
'MsgBox "Stream editing is disabled in the Trial Version.", vbInformation, "Disabled"
End Sub

Private Sub ExportButton_Click()
Dim buffer() As Byte
SubStreams.FileDlg.Flags = FileOpenConstants.cdlOFNHideReadOnly
SubStreams.FileDlg.FileName = ""
SubStreams.FileDlg.ShowSave
If SubStreams.FileDlg.FileName = "" Then Exit Sub
streamname = Main.StreamList.SelectedItem.Text & SubStreams.SubStreamList.SelectedItem.Text
FileNumber = FreeFile
Open streamname For Binary As FileNumber
Size = LOF(FileNumber)
ReDim buffer(Size)
Get FileNumber, 1, buffer
Close FileNumber
Open SubStreams.FileDlg.FileName For Binary As FileNumber
Put FileNumber, , buffer
datapos = datapos + UBound(buffer) + 1
If datapos = Size + 1 Then
Close FileNumber
End If
'MsgBox "The functionality to extract streams as files to the hard drive is disabled.", vbInformation, "Disabled"
End Sub

Private Sub Form_Load()
editclicked = 0
Me.Caption = Main.StreamList.SelectedItem.Text & SubStreams.SubStreamList.SelectedItem.Text
lblsz.Caption = "Stream Size: " & CStr(Format(Len(StreamBuffer) / 1024, "###")) & ".0 KB"
Text1.Text = StreamBuffer
End Sub

Private Sub Form_Unload(Cancel As Integer)
StreamBuffer = ""
If editclicked = 1 Then
    If MsgBox("Do you want to save changes?", vbYesNo + vbInformation, "Save?") = vbYes Then
    SaveEdit
    End If
End If
Unload SubStreams
SubStreams.Show
End Sub

Private Function SaveEdit()
Dim bigbuffer As String
Dim streamname As String
streamname = Main.StreamList.SelectedItem.Text & SubStreams.SubStreamList.SelectedItem.Text
bigbuffer = Text1.Text
DeleteFile streamname
FileNumber = FreeFile
Open streamname For Binary As FileNumber
Put FileNumber, , bigbuffer
Close FileNumber
End Function

Private Sub SaveButton_Click()
editclicked = 0
Text1.Locked = True
Call SaveEdit
SaveButton.Visible = False
End Sub
