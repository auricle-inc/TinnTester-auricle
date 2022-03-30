VERSION 5.00
Begin VB.UserControl whicheardial 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   ScaleHeight     =   4095
   ScaleWidth      =   8430
   Begin VB.PictureBox Pictureright_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":0000
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
   End
   Begin VB.PictureBox Pictureleft_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":96654
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   6
      Top             =   0
      Width           =   8430
   End
   Begin VB.PictureBox PictureClick_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":12CCA8
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
   End
   Begin VB.PictureBox Pictureboth_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":1C32FC
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
   End
   Begin VB.PictureBox PictureClick 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":259950
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
      Begin VB.Timer Timerclick 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   360
         Top             =   240
      End
   End
   Begin VB.PictureBox Pictureleft 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":2EFFA4
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   2
      Top             =   0
      Width           =   8430
   End
   Begin VB.PictureBox Pictureright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":3865F8
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
   End
   Begin VB.PictureBox Pictureboth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "whicheardial.ctx":41CC4C
      ScaleHeight     =   4110
      ScaleWidth      =   8430
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   8430
   End
End
Attribute VB_Name = "whicheardial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dir As Integer '1=mid, 0=left, 2=right

Private Sub update()
If English Then
    Pictureleft.visible = (dir = 0)
    Pictureright.visible = (dir = 2)
    Pictureboth.visible = (dir = 1)
    PictureClick.visible = False
    Pictureleft_F.visible = False
    Pictureright_F.visible = False
    Pictureboth_F.visible = False
    PictureClick_F.visible = False
Else
    Pictureleft_F.visible = (dir = 0)
    Pictureright_F.visible = (dir = 2)
    Pictureboth_F.visible = (dir = 1)
    PictureClick_F.visible = False
    Pictureleft.visible = False
    Pictureright.visible = False
    Pictureboth.visible = False
    PictureClick.visible = False
End If

End Sub


Public Function getvalue() As Integer
    getvalue = dir
End Function


Public Sub setvalue(ByVal Value As Integer)
    If Value >= 0 And Value <= 3 Then dir = Value
End Sub



Private Sub timerClick_Timer()
timerClick.Enabled = False
update
dir = dir + 100
End Sub

Public Sub UserControl_Initialize()
    dir = 0
If English Then
    Pictureleft.visible = (dir = 0)
    Pictureright.visible = (dir = 2)
    Pictureboth.visible = (dir = 1)
    PictureClick.visible = False
    Pictureleft_F.visible = False
    Pictureright_F.visible = False
    Pictureboth_F.visible = False
    PictureClick_F.visible = False
Else
    Pictureleft_F.visible = (dir = 0)
    Pictureright_F.visible = (dir = 2)
    Pictureboth_F.visible = (dir = 1)
    PictureClick_F.visible = False
    Pictureleft.visible = False
    Pictureright.visible = False
    Pictureboth.visible = False
    PictureClick.visible = False
End If

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = KeyCodeConstants.vbKeyRight Then
        If dir < 2 Then
            dir = dir + 1
            update
        End If
    ElseIf KeyCode = KeyCodeConstants.vbKeyLeft Then
        If dir > 0 Then
            dir = dir - 1
            update
        End If
    ElseIf KeyCode = 32 And (Not boolDblClick) Then 'user pressed space
        boolDblClick = True
        Form1.timerDblClick.Enabled = True
        If English Then
            PictureClick.visible = True
            PictureClick_F.visible = False
        Else
            PictureClick.visible = False
            Pictureleft_F.visible = False
            Pictureright_F.visible = False
            Pictureboth_F.visible = False
            PictureClick_F.visible = True
        End If
        timerClick.Enabled = True
    End If
End Sub

Private Sub UserControl_Resize()
    Pictureleft.Left = (UserControl.Width - Pictureleft.Width) / 2
    Pictureleft.Top = (UserControl.Height - Pictureleft.Height) / 2
    
    Pictureboth.Left = (UserControl.Width - Pictureboth.Width) / 2
    Pictureboth.Top = (UserControl.Height - Pictureboth.Height) / 2
    
    Pictureright.Left = (UserControl.Width - Pictureright.Width) / 2
    Pictureright.Top = (UserControl.Height - Pictureright.Height) / 2
    
    PictureClick.Left = (UserControl.Width - PictureClick.Width) / 2
    PictureClick.Top = (UserControl.Height - PictureClick.Height) / 2
End Sub
