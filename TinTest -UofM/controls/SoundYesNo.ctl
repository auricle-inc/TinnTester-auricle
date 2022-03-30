VERSION 5.00
Begin VB.UserControl soundYesNo 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   4935
   ScaleWidth      =   7200
   Begin VB.PictureBox pictureClick_F 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3585
      Left            =   720
      Picture         =   "SoundYesNo.ctx":0000
      ScaleHeight     =   3585
      ScaleWidth      =   8385
      TabIndex        =   6
      Top             =   1680
      Visible         =   0   'False
      Width           =   8385
   End
   Begin VB.PictureBox Pictureleft_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   1320
      Picture         =   "SoundYesNo.ctx":9485C
      ScaleHeight     =   4110
      ScaleMode       =   0  'User
      ScaleWidth      =   8265
      TabIndex        =   5
      Top             =   2400
      Width           =   8325
   End
   Begin VB.PictureBox Pictureright_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   1560
      Picture         =   "SoundYesNo.ctx":1290B8
      ScaleHeight     =   4110
      ScaleWidth      =   8325
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   8325
   End
   Begin VB.Timer timerClick 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Pictureleft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   3120
      Picture         =   "SoundYesNo.ctx":1BD914
      ScaleHeight     =   4110
      ScaleMode       =   0  'User
      ScaleWidth      =   8265
      TabIndex        =   0
      Top             =   0
      Width           =   8325
      Begin VB.PictureBox pictureClick 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   3585
         Left            =   2160
         Picture         =   "SoundYesNo.ctx":252170
         ScaleHeight     =   3585
         ScaleWidth      =   8385
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   8385
      End
   End
   Begin VB.PictureBox Pictureright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   2160
      Picture         =   "SoundYesNo.ctx":2E69CC
      ScaleHeight     =   4110
      ScaleWidth      =   8325
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   8325
   End
   Begin VB.PictureBox Picturemiddle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "SoundYesNo.ctx":37B228
      ScaleHeight     =   4110
      ScaleWidth      =   8325
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8325
   End
End
Attribute VB_Name = "soundYesNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dir As Integer '0=left, 1=right

Private Sub update()
If English Then
    Pictureleft.visible = (dir = 0)
    Pictureright.visible = (dir = 1)
    Picturemiddle.visible = (dir = 2)
    Pictureleft_F.visible = False
    Pictureright_F.visible = False
    'Picturemiddle_F.visible = False
Else
    Pictureleft_F.visible = (dir = 0)
    Pictureright_F.visible = (dir = 1)
    'Picturemiddle_F.visible = (dir = 2)
    Pictureleft.visible = False
    Pictureright.visible = False
    Picturemiddle.visible = False
End If
    'pictureClick.visible = False
    Form1.chkChange.Value = 1
End Sub


Public Function getvalue() As Integer
    getvalue = dir
End Function


Public Sub setvalue(ByVal Value As Integer)
    If Value >= 0 And Value <= 2 Then dir = Value
End Sub



Private Sub timerClick_Timer()
        timerClick.Enabled = False
        update
        Form1.chkClick.Value = 1
        Form1.chkChange.Value = 1
        
End Sub

Public Sub UserControl_Initialize()

    dir = 0
    If English Then
        Pictureleft.visible = (dir = 0)
        Pictureright.visible = (dir = 1)
        Picturemiddle.visible = (dir = 2)
        PictureClick.visible = False
        Pictureleft_F.visible = False
        Pictureright_F.visible = False
        'Picturemiddle_F.visible = False
        PictureClick_F.visible = False
    Else
        Pictureleft_F.visible = (dir = 0)
        Pictureright_F.visible = (dir = 1)
        'Picturemiddle_F.visible = (dir = 2)
        PictureClick_F.visible = False
        Pictureleft.visible = False
        Pictureright.visible = False
        'Picturemiddle.visible = False
        PictureClick.visible = False
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = KeyCodeConstants.vbKeyRight Then
        If dir < 1 Then
            dir = dir + 1
            update
        End If
    ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Then
        If dir > 0 Then
            dir = dir - 1
            update
        End If
    ElseIf KeyCode = KeyCodeConstants.vbKeySpace And (Not boolDblClick) Then
        If English Then
            Pictureleft.visible = False
            Pictureright.visible = False
            PictureClick.visible = True
            Pictureleft_F.visible = False
            Pictureright_F.visible = False
            PictureClick_F.visible = False
        Else
            Pictureleft_F.visible = False
            Pictureright_F.visible = False
            PictureClick_F.visible = True
            Pictureleft.visible = False
            Pictureright.visible = False
            PictureClick.visible = False
        End If
        boolDblClick = True
        Form1.timerDblClick.Enabled = True
        'pictureClick.visible = True
        'MsgBox ("Normal left = " & Pictureright.Left & " normal top = " & Pictureright.Top)
        'MsgBox ("left = " & pictureClick.Left & "  top = " & pictureClick.Top)
        timerClick.Enabled = True
    End If
    
End Sub

Private Sub UserControl_Resize()
    Pictureleft.Left = (UserControl.Width - Pictureleft.Width) / 2
    Pictureleft.Top = (UserControl.Height - Pictureleft.Height) / 2
    
    Picturemiddle.Left = (UserControl.Width - Picturemiddle.Width) / 2
    Picturemiddle.Top = (UserControl.Height - Picturemiddle.Height) / 2
    
    Pictureright.Left = (UserControl.Width - Pictureright.Width) / 2
    Pictureright.Top = (UserControl.Height - Pictureright.Height) / 2
    
    PictureClick.Left = (UserControl.Width - PictureClick.Width) / 2
    PictureClick.Top = (UserControl.Height - PictureClick.Height) / 2
    
    Pictureleft_F.Left = (UserControl.Width - Pictureleft_F.Width) / 2
    Pictureleft_F.Top = (UserControl.Height - Pictureleft_F.Height) / 2
    
'    Picturemiddle_F.Left = (UserControl.Width - Picturemiddle.Width) / 2
'    Picturemiddle_F.Top = (UserControl.Height - Picturemiddle.Height) / 2
    
    Pictureright_F.Left = (UserControl.Width - Pictureright_F.Width) / 2
    Pictureright_F.Top = (UserControl.Height - Pictureright_F.Height) / 2
    
    PictureClick_F.Left = (UserControl.Width - PictureClick_F.Width) / 2
    PictureClick_F.Top = (UserControl.Height - PictureClick_F.Height) / 2
End Sub
