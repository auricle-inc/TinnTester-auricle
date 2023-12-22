VERSION 5.00
Begin VB.Form FormReg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "myapp.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Exit Program"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   7560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Left            =   6120
      MaskColor       =   &H8000000F&
      Picture         =   "myapp.frx":211FA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5160
      Width           =   345
   End
   Begin VB.TextBox txtInstCode 
      Height          =   2535
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "myapp.frx":2150C
      Top             =   4680
      Width           =   3495
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   375
      Left            =   6120
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "Tinnitus Tester Registration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   3
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   $"myapp.frx":2151F
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   5415
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "2) Press the ""Generate"" button to generate a unique        installation code for this computer."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label lblInstructions 
      BackStyle       =   0  'Transparent
      Caption         =   "1) Enter your user/institution name in the box below."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label lblInstCode 
      BackStyle       =   0  'Transparent
      Caption         =   "Installation Code:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "FormReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
Attribute ActiveLockEventSink.VB_VarHelpID = -1


Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtInstCode.Text

End Sub

Private Sub cmdFinish_Click()
End
End Sub

Private Sub cmdGen_Click()
    txtInstCode = ActiveLock.InstallationCode(txtUser)
End Sub

' Re-compute Installation Code when "Registered User" changes
Private Sub txtUser_Change()

End Sub

