VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SubStreams 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Streams"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "Streams.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileDlg 
      Left            =   240
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files *.*|*.*"
   End
   Begin MSComctlLib.ListView SubStreamList 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu SubMenu 
      Caption         =   "SubMenu"
      Visible         =   0   'False
      Begin VB.Menu ViewStream 
         Caption         =   "View Stream"
      End
      Begin VB.Menu ExtractStream 
         Caption         =   "Extract Stream"
      End
      Begin VB.Menu DeleteStream 
         Caption         =   "Delete Stream"
      End
   End
End
Attribute VB_Name = "SubStreams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim streamname As String

Private Sub Form_Load()
SetWindowPos hwnd, conHwndTopmost, Screen.TwipsPerPixelX * 16, Screen.TwipsPerPixelY * 15, 600, 230, conSwpNoActivate Or conSwpShowWindow
Main.Enabled = False
Dim ColHead As ColumnHeader
SubStreamList.ColumnHeaders.Clear
Set ColHead = SubStreamList.ColumnHeaders.Add(, , "Stream Name", 3500)
Set ColHead = SubStreamList.ColumnHeaders.Add(, , "Stream Type", 1600)
Set ColHead = SubStreamList.ColumnHeaders.Add(, , "Stream Size", 1400)
Set ColHead = SubStreamList.ColumnHeaders.Add(, , "Stream Attributes", 1800)

StreamEnlist (Main.Objlbl.Caption)
End Sub

Private Sub Form_Unload(Cancel As Integer)
formstat = 1
Main.Enabled = True
Main.SetFocus
Main.EnumerateBtn_Click
End Sub

Private Sub SubStreamList_DblClick()
Call ViewStream_Click
End Sub

Private Sub SubStreamList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Obj As ListItem
Set Obj = SubStreamList.SelectedItem
If Obj.Text <> "" Then
    If Button = 2 Then
        PopupMenu SubMenu, , x + 600, y + 200
    End If
End If
End Sub

Private Sub ViewStream_Click()
Dim hFile As Long, Size As Long, buffer As String, BytesRead As Long
If Main.OnlyADSList.Visible = True And Main.StreamList.Visible = False Then
streamname = Main.OnlyADSList.SelectedItem.Text & SubStreamList.SelectedItem.Text
ElseIf Main.OnlyADSList.Visible = False And Main.StreamList.Visible = True Then
streamname = Main.StreamList.SelectedItem.Text & SubStreamList.SelectedItem.Text
End If
hFile = CreateFileW(StrPtr(streamname), FileAccess.FRead, FileShare.FRead, 0&, FileMode.OpenExisting, 0&, 0&)
Size = GetFileSize(hFile, 0&)
buffer = String$(Size, 0)
ReadFile hFile, ByVal buffer, Size, BytesRead, 0
CloseHandle hFile
StreamBuffer = buffer
If Len(StreamBuffer) > 5048205 Then
    If MsgBox("The streamsize is more then 5 MB. You can use the Extract option to extract the stream to a file on your computer and then read using an external program. " & vbCrLf & "Do you wish to continue with the display. The program may appear to have hanged.", vbExclamation + vbYesNo, "Too Big!!") = vbYes Then
        ViewFrm.Show
        Me.Hide
        DoEvents
    End If
Exit Sub
End If
ViewFrm.Show
Me.Hide
End Sub

Private Sub DeleteStream_Click()
If MsgBox("Do you really want to delete this stream?", vbInformation + vbYesNo, "Sure?") = vbYes Then
If Main.StreamList.Visible = True Then
streamname = Main.StreamList.SelectedItem.Text & SubStreamList.SelectedItem.Text
Else
streamname = Main.OnlyADSList.SelectedItem.Text & SubStreamList.SelectedItem.Text
End If
   Dim ret As Long
   ret = DeleteFile(streamname)
   'MsgBox (GetWin32ErrorDescription(5))
   'Unload Me
   Me.Show
End If
'MsgBox "Delete functionality disabled in the Trial Version.", vbInformation, "Disabled"
End Sub

Private Sub ExtractStream_Click()
Dim buffer() As Byte
FileDlg.Flags = FileOpenConstants.cdlOFNHideReadOnly
FileDlg.FileName = ""
FileDlg.ShowSave
If FileDlg.FileName = "" Then Exit Sub
If Main.StreamList.Visible = True Then
streamname = Main.StreamList.SelectedItem.Text & SubStreamList.SelectedItem.Text
Else
streamname = Main.OnlyADSList.SelectedItem.Text & SubStreamList.SelectedItem.Text
End If
FileNumber = FreeFile
Open streamname For Binary As FileNumber
Size = LOF(FileNumber)
ReDim buffer(Size)
Get FileNumber, 1, buffer
Close FileNumber
Open FileDlg.FileName For Binary As FileNumber
Dim datapos As Long
Put FileNumber, , buffer
datapos = datapos + UBound(buffer) + 1
If datapos = Size + 1 Then
Close FileNumber
End If
'MsgBox "The functioanlity to extract streams as files to the hard drive is disabled.", vbInformation, "Disabled"
End Sub
