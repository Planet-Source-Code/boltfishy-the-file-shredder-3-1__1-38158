VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search for"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAddTemp3 
      Caption         =   "Add Temp Files [.~]"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CheckBox chkAddTemp2 
      Caption         =   "Add Temp Files [.$$$]"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox chkAddBak 
      Caption         =   "Add Bak Files [.bak]"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Add All Files [*.*]"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ListBox tempDir 
      Height          =   270
      Left            =   2280
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox chkAdd0 
      Caption         =   "Add Zero Length Files"
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox chkAddTemp 
      Caption         =   "Add Temp Files [.tmp]"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   3960
      Width           =   5040
      _ExtentX        =   8890
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14111
            MinWidth        =   14111
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox lstDir 
      Height          =   1530
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   15
      X2              =   5579
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   5549
      Y1              =   3360
      Y2              =   3360
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------

Private Sub Command1_Click() 'search

ScanDir

If chkAddTemp.Value = 1 Then 'add temp *.tmp
FindTemp
End If

If chkAddTemp2.Value = 1 Then 'add temp *.$$$
FindTemp2
End If

If chkAddTemp3.Value = 1 Then 'add temp *.~
FindTemp3
End If

If chkAdd0.Value = 1 Then 'add 0 length files
FindZero
End If

If chkAddBak.Value = 1 Then 'add bak .bak files
FindBak
End If

If chkAll.Value = 1 Then 'add all files
FindAll
End If

Unload Me 'when done, close this screen

End Sub

Private Sub Command2_Click() 'close
Unload Me
End Sub

Public Sub FindTemp() 'find temp .tmp files

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.tmp"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If (GetAttr(tempDir.List(cnt) & "\" & FileName) And vbNormal) = vbNormal Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub FindTemp2() 'find temp .$$$ files

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.$$$"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If (GetAttr(tempDir.List(cnt) & "\" & FileName) And vbNormal) = vbNormal Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub FindTemp3() 'find temp .$$$ files

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.~"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If (GetAttr(tempDir.List(cnt) & "\" & FileName) And vbNormal) = vbNormal Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub FindZero() 'find files of 0 length

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.*"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If FileLen(tempDir.List(cnt) & "\" & FileName) = 0 Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub FindAll() 'find all files

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.*"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If (GetAttr(tempDir.List(cnt) & "\" & FileName) And vbNormal) = vbNormal Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub FindBak() 'find bak .bak files

cnt = 0
Do While cnt < tempDir.ListCount
StatusBar1.Panels(1).Text = "Searching " & tempDir.List(cnt) & "..."
FilePath = tempDir.List(cnt) + "\*.bak"
FileName = Dir(FilePath, vbNormal)

Do While FileName <> ""
If FileName <> "." And FileName <> ".." Then
If (GetAttr(tempDir.List(cnt) & "\" & FileName) And vbNormal) = vbNormal Then
frmMain.lstFiles.AddItem tempDir.List(cnt) & "\" & FileName
End If
End If

FileName = Dir
Loop
cnt = cnt + 1
Loop

End Sub

Public Sub ScanDir() 'scan directory

StatusBar1.Panels(1).Text = "Scanning..." & Drive1.Drive

Path = lstDir.Path + "\"
FileName = Dir(Path, vbDirectory)

Do While FileName <> ""

If FileName <> "." And Name <> ".." Then
If (GetAttr(Path & FileName) And vbDirectory) = vbDirectory Then

tempDir.AddItem Path & FileName

End If
End If

FileName = Dir

Loop
cnt = 0

Do While cnt < tempDir.ListCount

Path = tempDir.List(cnt) + "\"
FileName = Dir(Path, vbDirectory)

Do While FileName <> ""

If FileName <> "." And FileName <> ".." Then
If (GetAttr(Path & FileName) And vbDirectory) = vbDirectory Then

tempDir.AddItem Path & FileName

End If
End If

FileName = Dir

Loop
cnt = cnt + 1
Loop

StatusBar1.Panels(1).Text = "Finished scanning"

Exit Sub

End Sub


Private Sub Drive1_Change()

On Error GoTo ErrSub
lstDir.Path = Drive1.Drive
Me.Refresh
ErrSub:
Drive1.Drive = lstDir.Path

End Sub
