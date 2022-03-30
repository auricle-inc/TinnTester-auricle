VERSION 5.00
Begin VB.UserControl Choice123 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7200
   ScaleHeight     =   4110
   ScaleWidth      =   7200
   Begin VB.PictureBox PictureClick_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "Choice123.ctx":0000
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
   End
   Begin VB.PictureBox Pictureright_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "Choice123.ctx":AD684
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
   End
   Begin VB.PictureBox Picturemiddle_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "Choice123.ctx":15AD08
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
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
      Picture         =   "Choice123.ctx":20838C
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   4
      Top             =   0
      Width           =   9720
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
      Picture         =   "Choice123.ctx":2B5A10
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
      Begin VB.Timer TimerC 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   240
         Top             =   120
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
      Left            =   480
      Picture         =   "Choice123.ctx":363094
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   840
      Width           =   9720
   End
   Begin VB.PictureBox Pictureright 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   360
      Picture         =   "Choice123.ctx":410718
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   9720
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
      Picture         =   "Choice123.ctx":4BDD9C
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
   End
End
Attribute VB_Name = "Choice123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dir As Integer '1=mid, 0=left, 2=right

Private Sub update()
If English Then
    Pictureleft.visible = (dir = 0)
    Pictureright.visible = (dir = 2)
    Picturemiddle.visible = (dir = 1)
    PictureClick.visible = False
    Pictureleft_F.visible = False
    Pictureright_F.visible = False
    Picturemiddle_F.visible = False
    PictureClick_F.visible = False
Else
    Pictureleft_F.visible = (dir = 0)
    Pictureright_F.visible = (dir = 2)
    Picturemiddle_F.visible = (dir = 1)
    PictureClick_F.visible = False
    Pictureleft.visible = False
    Pictureright.visible = False
    Picturemiddle.visible = False
    PictureClick.visible = False
End If
    Form1.chkChange.Value = 1
        
End Sub


Public Function getvalue() As Integer
    getvalue = dir
End Function


Public Sub setvalue(ByVal Value As Integer)
    If Value >= 0 And Value <= 3 Then dir = Value
End Sub


Private Sub Timerc_Timer()
    TimerC.Enabled = False
    update
    Form1.chkClick.Value = 1
End Sub

Public Sub UserControl_Initialize()
    dir = 0
    If English Then
        Pictureleft.visible = (dir = 0)
        Pictureright.visible = (dir = 2)
        Picturemiddle.visible = (dir = 1)
        PictureClick.visible = False
        Pictureleft_F.visible = False
        Pictureright_F.visible = False
        Picturemiddle_F.visible = False
        PictureClick_F.visible = False
    Else
    Pictureleft_F.visible = (dir = 0)
    Pictureright_F.visible = (dir = 2)
    Picturemiddle_F.visible = (dir = 1)
    PictureClick_F.visible = False
    Pictureleft.visible = False
    Pictureright.visible = False
    Picturemiddle.visible = False
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
    ElseIf (KeyCode = 32) And (Not boolDblClick) Then 'user hit spacebar, off not selected
        boolDblClick = True
        Form1.timerDblClick.Enabled = True
        If English Then
            PictureClick.visible = True
            Pictureleft.visible = False
            Pictureright.visible = False
            Picturemiddle.visible = False
            PictureClick_F.visible = False
            Pictureleft_F.visible = False
            Pictureright_F.visible = False
            Picturemiddle_F.visible = False
        Else
            PictureClick.visible = False
            Pictureleft.visible = False
            Pictureright.visible = False
            Picturemiddle.visible = False
            PictureClick_F.visible = True
            Pictureleft_F.visible = False
            Pictureright_F.visible = False
            Picturemiddle_F.visible = False
        End If
            
        TimerC.Enabled = True
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
    
End Sub
