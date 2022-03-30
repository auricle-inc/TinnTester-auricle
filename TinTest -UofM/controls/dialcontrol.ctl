VERSION 5.00
Begin VB.UserControl dialcontrol 
   BackColor       =   &H00F0F0E6&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   5925
   Begin VB.Timer TimerClick 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   5520
      Top             =   240
   End
   Begin VB.PictureBox PictureClick 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   120
      Picture         =   "dialcontrol.ctx":0000
      ScaleHeight     =   3405
      ScaleWidth      =   3405
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Timer click 
      Interval        =   350
      Left            =   5040
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   4920
      TabIndex        =   11
      Top             =   1380
      Width           =   405
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   105
      Top             =   3705
   End
   Begin VB.Frame backgroundframe 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0E6&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.Frame volumeframe 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3405
         Left            =   3525
         TabIndex        =   2
         Top             =   120
         Width           =   1035
         Begin VB.PictureBox volumesliderparent 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2685
            Left            =   345
            ScaleHeight     =   2655
            ScaleWidth      =   285
            TabIndex        =   3
            Top             =   345
            Width           =   315
            Begin VB.PictureBox volumeslider 
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Height          =   15
               Left            =   0
               ScaleHeight     =   15
               ScaleWidth      =   285
               TabIndex        =   6
               Top             =   -10000
               Width           =   285
            End
         End
         Begin VB.Label volumelabel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   420
            TabIndex        =   5
            Top             =   60
            Width           =   150
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "VOLUME"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   4
            Top             =   3090
            Width           =   1005
         End
      End
      Begin VB.PictureBox dialpicture 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3405
         Left            =   120
         Picture         =   "dialcontrol.ctx":25EC8
         ScaleHeight     =   3405
         ScaleWidth      =   3405
         TabIndex        =   1
         Top             =   120
         Width           =   3405
         Begin VB.PictureBox yellowarrowleft 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   945
            Picture         =   "dialcontrol.ctx":4BD90
            ScaleHeight     =   525
            ScaleWidth      =   780
            TabIndex        =   9
            Top             =   375
            Width           =   780
         End
         Begin VB.PictureBox yellowarrowright 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   540
            Left            =   1695
            Picture         =   "dialcontrol.ctx":4D326
            ScaleHeight     =   540
            ScaleWidth      =   780
            TabIndex        =   10
            Top             =   360
            Width           =   780
         End
         Begin VB.PictureBox blackarrowleft 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   585
            Left            =   915
            Picture         =   "dialcontrol.ctx":4E958
            ScaleHeight     =   585
            ScaleWidth      =   795
            TabIndex        =   7
            Top             =   330
            Width           =   795
         End
         Begin VB.PictureBox blackarrowright 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   600
            Left            =   1710
            Picture         =   "dialcontrol.ctx":501FA
            ScaleHeight     =   600
            ScaleWidth      =   780
            TabIndex        =   8
            Top             =   285
            Width           =   780
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   7
            Visible         =   0   'False
            X1              =   1740
            X2              =   1740
            Y1              =   1605
            Y2              =   1765
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   6
            Visible         =   0   'False
            X1              =   1785
            X2              =   2145
            Y1              =   1800
            Y2              =   1680
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   5
            Visible         =   0   'False
            X1              =   1785
            X2              =   2145
            Y1              =   1800
            Y2              =   1920
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   4
            Visible         =   0   'False
            X1              =   1785
            X2              =   2145
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   3
            Visible         =   0   'False
            X1              =   1320
            X2              =   1680
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   2
            Visible         =   0   'False
            X1              =   1305
            X2              =   1665
            Y1              =   1680
            Y2              =   1800
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   1
            Visible         =   0   'False
            X1              =   1740
            X2              =   1740
            Y1              =   1845
            Y2              =   2005
         End
         Begin VB.Line LineClick 
            BorderWidth     =   2
            Index           =   0
            Visible         =   0   'False
            X1              =   1305
            X2              =   1665
            Y1              =   1920
            Y2              =   1800
         End
      End
      Begin VB.Shape mainframe 
         BorderWidth     =   3
         Height          =   3465
         Left            =   120
         Top             =   120
         Visible         =   0   'False
         Width           =   4500
      End
   End
End
Attribute VB_Name = "dialcontrol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private maxspeed As Integer
Private maxvolume
Private lasttime As Long
Private fvolume As Integer
Private Direction As Integer
Private MaxPast As Integer 'stores how many times the user tries to turn the volume up past 0

Public Function getvolume() As Integer
    If MaxPast < 5 Then
        getvolume = fvolume
    Else 'volume cannot go any louder
        getvolume = 999
    End If
End Function


Public Sub setvolume(ByVal volume As Integer)
    movevolume volume - fvolume
End Sub

Public Sub movevolume(ByVal delta As Integer)
    fvolume = fvolume + delta
    If (delta = -1) And (MaxPast > 0) Then MaxPast = 0 'user turned down volume so reset maxpast
    If fvolume < 0 Then
        fvolume = 0
    ElseIf fvolume > maxvolume Then
        fvolume = maxvolume
        MaxPast = MaxPast + 1
    End If
    volumeslider.Height = Int((volumesliderparent.Height - 2) * fvolume / maxvolume)
    volumeslider.Top = volumesliderparent.Height - 2 - volumeslider.Height
    volumelabel.Caption = CStr(fvolume)
    Form1.chkChange = 1
End Sub

Public Sub setvolumevisible(ByVal visible As Boolean)
    
    If Not visible Then
        mainframe.Width = dialpicture.Width + 4 * 15
    Else
        mainframe.Width = dialpicture.Width + 4 * 15 + volumeframe.Width
    End If
    volumeframe.visible = visible
    backgroundframe.Width = mainframe.Width + 180
End Sub

Private Sub click_Timer()
Dim c9 As Integer
'    If lineClick(0).visible = False Then
'        For c9 = 0 To 7 Step 1
'            lineClick(c9).visible = True
'        Next c9
'    Else
        For c9 = 0 To 7 Step 1
            lineClick(c9).visible = False
        Next c9
        Click.Enabled = False
'    End If

End Sub

Private Sub Command1_Click()
    setvolumevisible Not volumeframe.visible
End Sub

Private Sub Timer1_Timer()
  
    
    blackarrowleft.visible = fvolume > 0
    blackarrowright.visible = fvolume < maxvolume
        
    yellowarrowleft.visible = (Direction = -1) And (lasttime = 0 Or ((Sin(GetTickCount / 30) > 0) And blackarrowleft.visible))
    yellowarrowright.visible = (Direction = 1) And (lasttime = 0 Or ((Sin(GetTickCount / 30) > 0) And blackarrowright.visible))
        
        
    yellowarrowright.visible = blackarrowright.visible And yellowarrowright.visible
    yellowarrowleft.visible = blackarrowleft.visible And yellowarrowleft.visible
        
       
    If lasttime > 0 And GetTickCount - lasttime > 100 Then
        Timer1.Enabled = False
        yellowarrowleft.visible = False
        yellowarrowright.visible = False
        lasttime = 0
    End If


  
End Sub

Private Sub timerClick_Timer()
    PictureClick.visible = False
    timerClick.Enabled = False
    Form1.chkChange = 1
    Form1.chkClick.Value = 1
End Sub

Public Sub UserControl_Initialize()

    maxspeed = 100
    MaxPast = 0
    fvolume = 0
    maxvolume = 120
    
    'hide volume bar at side of control
    mainframe.Width = dialpicture.Width + 4 * 15
    volumeframe.visible = visible
    backgroundframe.Width = mainframe.Width + 180

    
End Sub

Private Sub changevolume(ByVal up As Boolean)
    If up Then
        Direction = 1
    Else
        Direction = -1
    End If
    movevolume Direction
    
   
   lasttime = GetTickCount
    Timer1.Interval = 1
    Timer1.Enabled = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyRight Then
        changevolume True
    ElseIf KeyCode = vbKeyLeft Then
        changevolume False
    ElseIf KeyCode = vbKeySpace And (Not boolDblClick) Then

         'LineClick(0).visible = True
         'LineClick(1).visible = True
         'LineClick(2).visible = True
         'LineClick(3).visible = True
         'LineClick(4).visible = True
         'LineClick(5).visible = True
         'LineClick(6).visible = True
         'LineClick(7).visible = True
         PictureClick.visible = True
         boolDblClick = True
         Form1.timerDblClick.Enabled = True
         timerClick.Enabled = True
         
    End If
End Sub

Public Sub Show_Arrows()
'called to show both arrows when moving between trials.  Implimented Nov 2013
'to fix bug where 2nd yellow arrow would stay invisible if user maxed out loudness
    blackarrowleft.visible = True
    blackarrowright.visible = True
End Sub
