VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2865
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1977.474
   ScaleMode       =   0  'User
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   0
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "Portions of code (C) 2000 Michael Bos."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   901.49
      X2              =   6126.374
      Y1              =   496.957
      Y2              =   496.957
   End
   Begin VB.Label lblOpenWeb 
      Caption         =   "http://www26.brinkster.com/boltfish/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Roxfort Solutions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblDetails 
      Caption         =   "The File Shredder was written by Mischa Balen. Copyright (C)  2002 Roxfort Solutions, INC. All rights reserved."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Image imgIcon 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   5337.57
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   5323.484
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "The File Shredder 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "The File Shredder 3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   240
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   901.49
      X2              =   6112.288
      Y1              =   496.957
      Y2              =   496.957
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------
    
    Private Sub cmdOK_Click() 'OK
    Unload Me 'close
    End Sub

    Private Sub Form_Load()
    Me.Caption = "About v" & App.Major & "." & App.Minor
    End Sub
