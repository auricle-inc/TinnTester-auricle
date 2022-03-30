VERSION 5.00
Begin VB.UserControl soundtypedial 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11700
   ScaleHeight     =   4095
   ScaleWidth      =   11700
   Begin VB.PictureBox Picturesteady_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":0000
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
   End
   Begin VB.PictureBox Picturepulsing_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":AD684
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
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
      Picture         =   "soundtypedial.ctx":15AD08
      ScaleHeight     =   4110
      ScaleWidth      =   9720
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   9720
   End
   Begin VB.PictureBox Pictureoff_F 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":20838C
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
      Picture         =   "soundtypedial.ctx":2B5A10
      ScaleHeight     =   4110
      ScaleWidth      =   11985
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   11985
      Begin VB.Timer Timerc 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   240
         Top             =   240
      End
   End
   Begin VB.PictureBox Picturepulsing 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":38B70C
      ScaleHeight     =   4110
      ScaleWidth      =   11985
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11985
   End
   Begin VB.PictureBox Picturesteady 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":461408
      ScaleHeight     =   4110
      ScaleWidth      =   11985
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   11985
   End
   Begin VB.PictureBox Pictureoff 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4110
      Left            =   0
      Picture         =   "soundtypedial.ctx":537104
      ScaleHeight     =   4110
      ScaleWidth      =   11985
      TabIndex        =   1
      Top             =   0
      Width           =   11985
   End
End
Attribute VB_Name = "soundtypedial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public dir As Integer '1=mid, 0=left, 2=right

Private Sub update()
If English Then
    Picturesteady.visible = (dir = 0)
    Picturepulsing.visible = (dir = 2)
    Pictureoff.visible = (dir = 1)
    PictureClick.visible = False
    Picturesteady_F.visible = False
    Picturepulsing_F.visible = False
    Pictureoff_F.visible = False
    PictureClick_F.visible = False
Else
    Picturesteady_F.visible = (dir = 0)
    Picturepulsing_F.visible = (dir = 2)
    Pictureoff_F.visible = (dir = 1)
    PictureClick_F.visible = False
    Picturesteady.visible = False
    Picturepulsing.visible = False
    Pictureoff.visible = False
    PictureClick.visible = False
End If
    Form1.chkChange = 1
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
Form1.chkClick = 1
End Sub

Public Sub UserControl_Initialize()
    dir = 1
    If English Then
        Picturesteady.visible = False
        Picturepulsing.visible = False
        Pictureoff.visible = True
        PictureClick.visible = False
        Picturesteady_F.visible = False
        Picturepulsing_F.visible = False
        Pictureoff_F.visible = False
        PictureClick_F.visible = False
    Else
        Picturesteady.visible = False
        Picturepulsing.visible = False
        Pictureoff.visible = False
        PictureClick.visible = False
        Picturesteady_F.visible = False
        Picturepulsing_F.visible = False
        Pictureoff_F.visible = True
        PictureClick_F.visible = False
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
    ElseIf (KeyCode = 32 And dir <> 1 And (Not boolDblClick)) Then
        boolDblClick = True
        Form1.timerDblClick.Enabled = True
        Picturesteady.visible = False
        Picturepulsing.visible = False
        Pictureoff.visible = False
        Picturesteady_F.visible = False
        Picturepulsing_F.visible = False
        Pictureoff_F.visible = False
        If English Then
            PictureClick.visible = True
            PictureClick_F.visible = False
        Else
            PictureClick.visible = False
            PictureClick_F.visible = True
        End If
        TimerC.Enabled = True
    End If
End Sub

Private Sub UserControl_Resize()
    Picturesteady.Left = (UserControl.Width - Picturesteady.Width) / 2
    Picturesteady.Top = (UserControl.Height - Picturesteady.Height) / 2
    
    Pictureoff.Left = (UserControl.Width - Pictureoff.Width) / 2
    Pictureoff.Top = (UserControl.Height - Pictureoff.Height) / 2
    
    Picturepulsing.Left = (UserControl.Width - Picturepulsing.Width) / 2
    Picturepulsing.Top = (UserControl.Height - Picturepulsing.Height) / 2
    
    PictureClick.Left = (UserControl.Width - PictureClick.Width) / 2
    PictureClick.Top = (UserControl.Height - PictureClick.Height) / 2
End Sub
