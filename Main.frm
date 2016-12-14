VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NTStream"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Status "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   10935
      Begin VB.CommandButton ShowADSbutton 
         Caption         =   "S&how Only ADS"
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
         Left            =   9240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label FilesWithADSlbl 
         BackColor       =   &H80000009&
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
         Left            =   4440
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label ADSfilelbl 
         BackColor       =   &H80000009&
         Caption         =   "With ADS:"
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
         Left            =   3480
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Filenumlbl 
         BackColor       =   &H80000009&
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
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label DirWithADSlbl 
         BackColor       =   &H80000009&
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
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label ADSDirlbl 
         BackColor       =   &H80000009&
         Caption         =   "With ADS:"
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
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label DirNumlbl 
         BackColor       =   &H80000009&
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
         Left            =   1320
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label ScanFilelbl 
         BackColor       =   &H80000009&
         Caption         =   "Files:"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label ScanDirlbl 
         BackColor       =   &H80000009&
         Caption         =   "Directories:"
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
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Statlbl 
         BackColor       =   &H80000009&
         Caption         =   "NTStream Idle"
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
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label StatFilelbl 
         BackColor       =   &H80000009&
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
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   9255
      End
   End
   Begin MSComctlLib.ListView OnlyADSList 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
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
   Begin MSComDlg.CommonDialog FileDlg 
      Left            =   0
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select a file to find streams..."
      Filter          =   "All Files *.*|*.*"
   End
   Begin VB.CommandButton EnumerateBtn 
      Caption         =   "&Enumerate"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Browse 
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
      Left            =   7920
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin MSComctlLib.ListView StreamList 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8281
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
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Image LogoImg 
      Height          =   960
      Left            =   240
      Picture         =   "Main.frx":0442
      Top             =   120
      Width           =   2265
   End
   Begin VB.Label Objlbl 
      BackColor       =   &H80000009&
      Height          =   255
      Left            =   9360
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label FileLbl 
      BackColor       =   &H80000009&
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
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Menu menuobj 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu FileOpenMenu 
         Caption         =   "For File"
      End
      Begin VB.Menu FolderOpenMenu 
         Caption         =   "For Folder"
      End
      Begin VB.Menu ScanFolderMenu 
         Caption         =   "Scan This Folder"
         Visible         =   0   'False
      End
      Begin VB.Menu ScanInMenu 
         Caption         =   "Scan && Include Sub Folders"
         Visible         =   0   'False
      End
      Begin VB.Menu CreateStream 
         Caption         =   "Create New Stream"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject
Dim eggcount As Integer

Public Sub Browse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    FolderOpenMenu.Visible = True
    ScanFolderMenu.Visible = False
    CreateStream.Visible = False
    FileOpenMenu.Visible = True
    ScanInMenu.Visible = False
    PopupMenu menuobj, , x + 8000, y + 1050
End If
StreamList.Visible = True
OnlyADSList.Visible = False
End Sub

Private Sub CreateStream_Click()
NewStream.Show
Me.Hide
End Sub

Public Sub EnumerateBtn_Click()
If EnumerateBtn.Caption = "&Stop Scan" Then Exit Sub
If isObjFile = 2 Then Exit Sub
If RootNTFS(FileLbl.Caption) Then
If fso.FileExists(FileLbl.Caption) = False Then
    MsgBox "The File seems to have disappeared off the hard disk. Please check if the file exists.", vbCritical, "File Gone!!"
Exit Sub
End If
If isObjFile = 1 Then
    ScanFiles = 0
    FileWithADS = 0
    ScanFolders = 0
    DirWithADS = 0
End If

EnumerateBtn.Caption = "&Stop Scan"
StreamList.ListItems.Clear
ShowADSbutton.Visible = False
Statlbl.Caption = "Now Scanning:"
quitjob = 0
    ScanFiles = 0
    FileWithADS = 0
    ScanFolders = 0
    DirWithADS = 0
    DirNumlbl.Caption = 0
    Filenumlbl.Caption = 0
    DirWithADSlbl.Caption = 0
    FilesWithADSlbl.Caption = 0
StreamHunter (FileLbl.Caption)
Statlbl.Caption = "Scan Complete"
If formstat = 0 Then
    MsgBox "Scan completed. Double click the File names to view streams or use the right click option to create additional streams.", vbInformation + vbOKOnly, "Complete"
End If
EnumerateBtn.Caption = "&Enumerate"
quitjob = 1
formstat = 0
Else
MsgBox "Alternate Data Streams are found only on NTFS drives. Please browse for another search.", vbInformation + vbOKOnly, "Non - NTFS"
End If
End Sub

Public Sub EnumerateBtn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If quitjob = 1 Then
If isObjFile = 2 Then
    If Button = 1 Then
    If fso.FolderExists(FileLbl.Caption) = False Then
        MsgBox "The folder seems to have disappeared off the hard disk!! Please check if the folder exists.", vbCritical, "Folder Gone!!"
        Exit Sub
    End If
        ScanFolderMenu.Visible = True
        CreateStream.Visible = False
        FolderOpenMenu.Visible = False
        FileOpenMenu.Visible = False
        ScanInMenu.Visible = True
        PopupMenu menuobj, , x + 9700, y + 1050
        ScanFiles = 0
    FileWithADS = 0
    ScanFolders = 0
    DirWithADS = 0
    End If
End If
End If
EnumerateBtn.Caption = "&Enumerate"
ShowADSbutton.Visible = True
quitjob = 1
StatFilelbl.Caption = ""
End Sub

Public Sub FileOpenMenu_Click()
FileDlg.Flags = FileOpenConstants.cdlOFNHideReadOnly
FileDlg.FileName = ""
FileDlg.ShowOpen
EnumerateBtn.Enabled = False
FileLbl.Caption = ""
isObjFile = 0
If FileDlg.FileName <> "" Then
    EnumerateBtn.Enabled = True
    FileLbl.Caption = FileDlg.FileName
    isObjFile = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload BrowseFrm
Unload NewStream
Unload SubStreams
Unload ViewFrm
End Sub


Private Sub LogoImg_DblClick()
eggcount = eggcount + 1
If eggcount = 3 Then
    Dim egg As String
    egg = "Programmed by: Riyaz Ahemed" & vbCrLf & "Name & Version: NTStream v1.2"
    egg = egg & vbCrLf & "Description: Tool that allows users to create, search, edit and delete NTFS Streams."
    MsgBox egg, vbInformation, "Info"
    eggcount = 0
End If
End Sub

Private Sub OnlyADSList_DblClick()
Dim Obj As ListItem
Set Obj = OnlyADSList.SelectedItem
If Obj.SubItems(1) = "" Then
    If MsgBox("No stream information to display. Do you want to create a new stream?", vbYesNo + vbExclamation, "No ADS") = vbYes Then
    NewStream.Show
    Me.Hide
    End If
    Exit Sub
End If
Objlbl.Caption = Obj.Text
SubStreams.Show
SubStreams.SetFocus
End Sub

Private Sub ScanFolderMenu_Click()
If RootNTFS(FileLbl.Caption) Then
quitjob = 0
If isObjFile = 1 Then Exit Sub
ScanFiles = 0
FileWithADS = 0
ScanFolders = 0
DirWithADS = 0
ShowADSbutton.Visible = False
StreamList.ListItems.Clear
EnumerateBtn.Caption = "&Stop Scan"
Statlbl.Caption = "Now Scanning:"
With Me
.ADSDirlbl.Visible = True
.ADSfilelbl.Visible = True
.ScanDirlbl.Visible = True
.ScanFilelbl.Visible = True
.DirNumlbl.Visible = True
.DirWithADSlbl.Visible = True
.Filenumlbl.Visible = True
.FilesWithADSlbl.Visible = True
End With
    ScanFiles = 0
    FileWithADS = 0
    ScanFolders = 0
    DirWithADS = 0
    DirNumlbl.Caption = 0
    Filenumlbl.Caption = 0
    DirWithADSlbl.Caption = 0
    FilesWithADSlbl.Caption = 0
StreamHunter (FileLbl.Caption)
Statlbl.Caption = "Scan Complete"
MsgBox "Scan completed. Double click the File names to view streams or use the right click option to create additional streams.", vbInformation + vbOKOnly, "Complete"
EnumerateBtn.Caption = "&Enumerate"
quitjob = 1
ShowADSbutton.Visible = True
Else
MsgBox "Alternate Data Streams are found only on NTFS drives. Please browse for another search.", vbInformation + vbOKOnly, "Non - NTFS"
End If
End Sub

Private Sub ScanInMenu_Click()
If RootNTFS(FileLbl.Caption) Then
quitjob = 0
If isObjFile = 1 Then Exit Sub
ScanFiles = 0
FileWithADS = 0
ScanFolders = 0
DirWithADS = 0
ShowADSbutton.Visible = False
StreamList.ListItems.Clear
EnumerateBtn.Caption = "&Stop Scan"
Statlbl.Caption = "Now Scanning:"
With Me
.ADSDirlbl.Visible = True
.ADSfilelbl.Visible = True
.ScanDirlbl.Visible = True
.ScanFilelbl.Visible = True
.DirNumlbl.Visible = True
.DirWithADSlbl.Visible = True
.Filenumlbl.Visible = True
.FilesWithADSlbl.Visible = True
End With
ScanFiles = 0
    FileWithADS = 0
    ScanFolders = 0
    DirWithADS = 0
    DirNumlbl.Caption = 0
    Filenumlbl.Caption = 0
    DirWithADSlbl.Caption = 0
    FilesWithADSlbl.Caption = 0
EnumDir (FileLbl.Caption)
Statlbl.Caption = "Scan Complete"
MsgBox "Scan completed. Double click the File names to view streams or use the right click option to create additional streams.", vbInformation + vbOKOnly, "Complete"
EnumerateBtn.Caption = "&Enumerate"
quitjob = 1
ShowADSbutton.Visible = True
Else
MsgBox "Alternate Data Streams are found only on NTFS drives. Please browse for another search.", vbInformation + vbOKOnly, "Non - NTFS"
End If
End Sub

Public Sub FolderOpenMenu_Click()
Dim Dir As String
Dir = BrowseFolder(Me)
If Len(Dir) = 0 Then
    EnumerateBtn.Enabled = False
    Exit Sub
End If

If Left(Dir, 2) = "\\" Then
        MsgBox "Network Folders not supported yet...", vbCritical, "Invalid Selection"
        FileLbl.Caption = ""
        EnumerateBtn.Enabled = False
        Exit Sub
End If

FileLbl.Caption = Dir
EnumerateBtn.Enabled = True
isObjFile = 2
'BrowseFrm.Show
End Sub

Public Sub Form_Load()
'CheckTrialCount
Dim ColHead As ColumnHeader
StreamList.ColumnHeaders.Clear
Set ColHead = StreamList.ColumnHeaders.Add(1, , "File/Folder Name", 7000)
Set ColHead = StreamList.ColumnHeaders.Add(2, , "Associated Streams", 3000)
Set ColHead = OnlyADSList.ColumnHeaders.Add(1, , "File/Folder Name", 7000)
Set ColHead = OnlyADSList.ColumnHeaders.Add(2, , "Associated Streams", 3000)
quitjob = 1
End Sub

Private Sub ShowADSbutton_Click()
If ShowADSbutton.Caption = "S&how Only ADS" Then
OnlyADSList.ListItems.Clear
Dim i As Long
StreamList.Visible = False
For i = 1 To StreamList.ListItems.Count
    If Len(StreamList.ListItems.Item(i).SubItems(1)) > 1 Then
      Set ListObj = OnlyADSList.ListItems.Add(, , StreamList.ListItems.Item(i).Text)
      ListObj.SubItems(1) = StreamList.ListItems.Item(i).SubItems(1)
    End If
Next
OnlyADSList.Visible = True
ShowADSbutton.Caption = "Show &All"
Else
StreamList.Visible = True
OnlyADSList.Visible = False
ShowADSbutton.Caption = "S&how Only ADS"
End If
End Sub

Private Sub StreamList_DblClick()
Dim Obj As ListItem
Set Obj = StreamList.SelectedItem
If Obj.SubItems(1) = "" Then
    If MsgBox("No stream information to display. Do you want to create a new stream?", vbYesNo + vbExclamation, "No ADS") = vbYes Then
    NewStream.Show
    Me.Hide
    End If
    Exit Sub
End If
Objlbl.Caption = Obj.Text
SubStreams.Show
SubStreams.SetFocus
End Sub

Private Sub StreamList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Obj As ListItem
Set Obj = StreamList.SelectedItem
If Obj.Text <> "" Then
    If Button = 2 Then
    Me.CreateStream.Visible = True
    Me.ScanFolderMenu.Visible = False
    Me.ScanInMenu.Visible = False
    Me.FolderOpenMenu.Visible = False
    Me.FileOpenMenu.Visible = False
    PopupMenu menuobj, , x + 200, y + 1800
End If
End If
End Sub

Private Function RootNTFS(rootfile As String) As Boolean
rootfile = Left(rootfile, 3)
Dim VolName As String, VolSz As Long, VolSerial As Long, MaxComp As Long
Dim FileSysFlags As Long, FileSysName As String, FileSysNameSz As Long
VolName = Space(255)
FileSysName = Space(255)
i = GetVolumeInformation(rootfile, VolName, Len(VolName), VolSerial, 0, 0, FileSysName, Len(FileSysName))
If LCase(Left(Trim(FileSysName), 4)) = "ntfs" Then
RootNTFS = True
End If
End Function

Private Sub CheckTrialCount()
On Error GoTo errfound
Dim Appdata As String, TCountFile As String
Appdata = Environ$("Appdata")
Appdata = Appdata + "\NTStream"
TCountFile = Appdata + "\TZC2oXuq5n8t.big"
If fso.FolderExists(Appdata) = False Then
fso.CreateFolder (Appdata)
Set fs = fso.CreateTextFile(TCountFile, True)
fs.WriteLine ("57")
fs.Close
MsgBox "This is a trial version of NT Stream. Some functionality may be disabled." & vbCrLf & "Number of valid trial runs remaining: 9", vbInformation, "Trial Version"
Set fs = Nothing
Else
    If fso.FileExists(TCountFile) Then
    Set fs = fso.OpenTextFile(TCountFile, ForReading)
        If FileLen(TCountFile) > 0 Then
            Data = fs.ReadLine
            Data = CLng(Data) - 1
            If CLng(Data) = 48 Then
                MsgBox "This Trial version of NT Stream has expired.", vbInformation, "Expired"
                fs.Close
                TrialTag.Caption = "Trial Version." & vbCrLf & "Valid runs remaining: " & CStr(Chr(Data))
                End
                Exit Sub
            End If
            fs.Close
            Set fs = fso.OpenTextFile(TCountFile, ForWriting)
            fs.WriteLine (Data)
            fs.Close
            MsgBox "This is a trial version of NT Stream. Some functionality may be disabled." & vbCrLf & "Number of valid trial runs remaining: " & Chr(CLng(Data)), vbInformation, "Trial Version"
            TrialTag.Caption = "Trial Version." & vbCrLf & "Valid runs remaining: " & CStr(Chr(Data))
            Exit Sub
        Else
        fs.Close
        Set fs = fso.OpenTextFile(TCountFile, ForWriting)
        fs.WriteLine ("57")
        fs.Close
        MsgBox "This is a trial version of NT Stream. Some functionality may be disabled." & vbCrLf & "Number of valid trial runs remaining: 9", vbInformation, "Trial Version"
        End If
    Else
    Set fs = fso.CreateTextFile(TCountFile, True)
    fs.WriteLine ("57")
    fs.Close
    MsgBox "This is a trial version of NT Stream. Some functionality may be disabled." & vbCrLf & "Number of valid trial runs remaining: 9", vbInformation, "Trial Version"
    End If
End If
TrialTag.Caption = "Trial Version." & vbCrLf & "Valid runs remaining: " & CStr(Chr(Data))
Exit Sub
errfound:
MsgBox "An error occurred. Please contact your system administrator. The specific error code is: " & CStr(Err.Number), vbCritical, "Error!!"
End
End Sub
