VERSION 5.00

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest_back.frm
' This code block includes the declaration of three ActiveX controls: PA5x.ocx, mscomctl.ocx, and comdlg32.ocx.
' These controls are referenced by their respective Object IDs and file names.
' The first control (PA5x.ocx) has an Object ID of "{73B7CFAB-7D2C-487A-81EC-E6A15FB9E84A}#1.0#0".
' The second control (mscomctl.ocx) has an Object ID of "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0".
' The third control (comdlg32.ocx) has an Object ID of "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0".
Object = "{73B7CFAB-7D2C-487A-81EC-E6A15FB9E84A}#1.0#0"; "PA5x.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"

' This code represents the design of a form named "Form1" in the Tinnitus Tester v2.0 application.
' It includes various controls such as command buttons, labels, and scrollbars.
' The form has a specific layout and appearance, with different colors and captions for each control.
' Some controls are initially hidden (Visible = False) and disabled (Enabled = False).
' The form is designed to be maximized (WindowState = 2) when displayed.
' The code snippet provided represents a portion of the form's design, including the Loudness Rating frame and the Left Ear frame.
' Each control within the frames has specific properties such as caption, size, position, and font settings.
' The form is part of the TinTest.frm file located at "C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm".

Begin VB.Form Form1 
   BackColor       =   &H00F0F0E6&
   Caption         =   "Tinnitus Tester v2.0"
   ClientHeight    =   14850
   ClientLeft      =   570
   ClientTop       =   345
   ClientWidth     =   18105
   Icon            =   "TinTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1207
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrintReport 
      Caption         =   "Create Report"
      Height          =   495
      Left            =   8280
      TabIndex        =   179
      Top             =   14520
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PrintFakeReport"
      Height          =   375
      Left            =   120
      TabIndex        =   178
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   120
      TabIndex        =   177
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame framLoudness 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Loudness Rating"
      ClipControls    =   0   'False
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4455
      Left            =   960
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   16815
      Begin VB.HScrollBar hscrScale 
         Height          =   1935
         Left            =   600
         Max             =   100
         Min             =   1
         TabIndex        =   18
         Top             =   600
         Value           =   1
         Width           =   15375
      End
      Begin VB.Label lbl95 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Extremely Strong"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13800
         TabIndex        =   24
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label lbl70 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Very Strong"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10680
         TabIndex        =   23
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lbl30 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Moderate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   4560
         TabIndex        =   21
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbl5 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Extremely Weak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   480
         TabIndex        =   20
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label lblScale 
         BackColor       =   &H00F0F0E6&
         Caption         =   "0         10          20        30         40        50          60        70        80         90       100"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   19
         Top             =   2880
         Width           =   15495
      End
      Begin VB.Line Line9 
         X1              =   14280
         X2              =   14280
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line8 
         X1              =   12720
         X2              =   12720
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line7 
         X1              =   9840
         X2              =   9840
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line6 
         X1              =   6840
         X2              =   6840
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line5 
         X1              =   2280
         X2              =   2280
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line4 
         X1              =   3840
         X2              =   3840
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line3 
         X1              =   11280
         X2              =   11280
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   5280
         Y1              =   2520
         Y2              =   2760
      End
      Begin VB.Line Line1 
         X1              =   8280
         X2              =   8280
         Y1              =   2760
         Y2              =   2520
      End
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   840
         Top             =   2520
         Width           =   14895
      End
      Begin VB.Label lbl50 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Strong"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7440
         TabIndex        =   22
         Top             =   3360
         Width           =   1815
      End
   End
   Begin VB.Frame frmMono 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Left Ear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   9855
      Index           =   1
      Left            =   1800
      TabIndex        =   113
      Top             =   2760
      Visible         =   0   'False
      Width           =   5895
      Begin VB.VScrollBar VScroll1 
         Height          =   8895
         Index           =   1
         LargeChange     =   2
         Left            =   2040
         Max             =   101
         Min             =   1
         TabIndex        =   114
         Top             =   600
         Value           =   101
         Width           =   1455
      End
      Begin VB.Line Line10 
         Index           =   1
         X1              =   1920
         X2              =   1920
         Y1              =   840
         Y2              =   9240
      End
      Begin VB.Line Line11 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line12 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Line Line13 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   1730
         Y2              =   1730
      End
      Begin VB.Line Line14 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   2570
         Y2              =   2570
      End
      Begin VB.Line Line15 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   3410
         Y2              =   3410
      End
      Begin VB.Line Line16 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   4250
         Y2              =   4250
      End
      Begin VB.Line Line17 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line18 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line19 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   6700
         Y2              =   6700
      End
      Begin VB.Line Line20 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   7520
         Y2              =   7520
      End
      Begin VB.Line Line21 
         Index           =   1
         X1              =   2040
         X2              =   1920
         Y1              =   8340
         Y2              =   8340
      End
      Begin VB.Label lblFive 
         BackColor       =   &H00F0F0E6&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   130
         Top             =   680
         Width           =   375
      End
      Begin VB.Label lb4 
         BackColor       =   &H00F0F0E6&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   129
         Top             =   1510
         Width           =   375
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00F0F0E6&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   128
         Top             =   2350
         Width           =   255
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00F0F0E6&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   1440
         TabIndex        =   127
         Top             =   3180
         Width           =   255
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00F0F0E6&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   126
         Top             =   4050
         Width           =   375
      End
      Begin VB.Label lbl0 
         BackColor       =   &H00F0F0E6&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   125
         Top             =   4850
         Width           =   255
      End
      Begin VB.Label lblN1 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   124
         Top             =   5700
         Width           =   615
      End
      Begin VB.Label lblN2 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   123
         Top             =   6510
         Width           =   615
      End
      Begin VB.Label lblN3 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1200
         TabIndex        =   122
         Top             =   7290
         Width           =   615
      End
      Begin VB.Label lblN4 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   121
         Top             =   8120
         Width           =   615
      End
      Begin VB.Label lblN5 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   120
         Top             =   9000
         Width           =   615
      End
      Begin VB.Label lblMuchLouder 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "TINNITUS MUCH LOUDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   119
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblLouder 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "LOUDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   1
         Left            =   3600
         TabIndex        =   118
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblNoChange 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "NO CHANGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   1
         Left            =   3600
         TabIndex        =   117
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label lblSofter 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "SOFTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   116
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label lblGone 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "TINNITUS GONE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   1
         Left            =   3600
         TabIndex        =   115
         Top             =   8880
         Width           =   2175
      End
   End
   Begin VB.Frame frmMono 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Right Ear"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   9855
      Index           =   0
      Left            =   7800
      TabIndex        =   95
      Top             =   3840
      Visible         =   0   'False
      Width           =   5895
      Begin VB.VScrollBar VScroll1 
         Height          =   8895
         Index           =   0
         LargeChange     =   2
         Left            =   2040
         Max             =   101
         Min             =   1
         TabIndex        =   96
         Top             =   600
         Value           =   91
         Width           =   1455
      End
      Begin VB.Label lblGone 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "TINNITUS GONE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Index           =   0
         Left            =   3600
         TabIndex        =   112
         Top             =   9000
         Width           =   2175
      End
      Begin VB.Label lblSofter 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "SOFTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   3600
         TabIndex        =   111
         Top             =   6960
         Width           =   2175
      End
      Begin VB.Label lblNoChange 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "Pas de changement"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   0
         Left            =   3600
         TabIndex        =   110
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label lblLouder 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "LOUDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   109
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblMuchLouder 
         Alignment       =   2  'Center
         BackColor       =   &H00F0F0E6&
         Caption         =   "TINNITUS MUCH LOUDER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   0
         Left            =   3600
         TabIndex        =   108
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblN5 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   107
         Top             =   9000
         Width           =   615
      End
      Begin VB.Label lblN4 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   106
         Top             =   8110
         Width           =   615
      End
      Begin VB.Label lblN3 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1200
         TabIndex        =   105
         Top             =   7280
         Width           =   615
      End
      Begin VB.Label lblN2 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   104
         Top             =   6480
         Width           =   615
      End
      Begin VB.Label lblN1 
         BackColor       =   &H00F0F0E6&
         Caption         =   "- 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   103
         Top             =   5680
         Width           =   615
      End
      Begin VB.Label lbl0 
         BackColor       =   &H00F0F0E6&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1500
         TabIndex        =   102
         Top             =   4850
         Width           =   400
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00F0F0E6&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1500
         TabIndex        =   101
         Top             =   4010
         Width           =   400
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00F0F0E6&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   1500
         TabIndex        =   100
         Top             =   3190
         Width           =   400
      End
      Begin VB.Label lbl3 
         BackColor       =   &H00F0F0E6&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1500
         TabIndex        =   99
         Top             =   2340
         Width           =   400
      End
      Begin VB.Label lb4 
         BackColor       =   &H00F0F0E6&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1500
         TabIndex        =   98
         Top             =   1500
         Width           =   400
      End
      Begin VB.Label lblFive 
         BackColor       =   &H00F0F0E6&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1500
         TabIndex        =   97
         Top             =   650
         Width           =   400
      End
      Begin VB.Line Line21 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   8350
         Y2              =   8350
      End
      Begin VB.Line Line20 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   7520
         Y2              =   7520
      End
      Begin VB.Line Line19 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   6700
         Y2              =   6700
      End
      Begin VB.Line Line18 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line17 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line16 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   4230
         Y2              =   4230
      End
      Begin VB.Line Line15 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   3410
         Y2              =   3410
      End
      Begin VB.Line Line14 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   2570
         Y2              =   2570
      End
      Begin VB.Line Line13 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   1730
         Y2              =   1730
      End
      Begin VB.Line Line12 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   9240
         Y2              =   9240
      End
      Begin VB.Line Line11 
         Index           =   0
         X1              =   2040
         X2              =   1920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line10 
         Index           =   0
         X1              =   1920
         X2              =   1920
         Y1              =   840
         Y2              =   9240
      End
   End
   Begin Project1.PitchControl PitchControl1 
      Height          =   4695
      Left            =   4440
      TabIndex        =   176
      Top             =   6000
      Visible         =   0   'False
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8281
   End
   Begin Project1.Choice123 Choice1231 
      Height          =   4815
      Left            =   3480
      TabIndex        =   175
      Top             =   6240
      Visible         =   0   'False
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8493
   End
   Begin VB.CommandButton cmdTinTrain 
      Caption         =   "Run TinTrain"
      Height          =   495
      Left            =   360
      TabIndex        =   174
      Top             =   14280
      Width           =   1335
   End
   Begin VB.TextBox txtPA5ThreshValue 
      Enabled         =   0   'False
      Height          =   285
      Left            =   5640
      TabIndex        =   173
      Top             =   13800
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundThreshold 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   5640
      TabIndex        =   172
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundThreshold 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   171
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundThreshold 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   170
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtIntensity2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   169
      Top             =   13320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Project1.soundtypedial soundtypedial1 
      Height          =   4335
      Left            =   3840
      TabIndex        =   152
      Top             =   5880
      Visible         =   0   'False
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7646
   End
   Begin VB.TextBox txtRIRightT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   17160
      TabIndex        =   166
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   16680
      TabIndex        =   165
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   16080
      TabIndex        =   164
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   15600
      TabIndex        =   163
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundLevelMatch 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   162
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdEar 
      Caption         =   "Dial Control"
      Height          =   375
      Index           =   4
      Left            =   13680
      TabIndex        =   159
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEar 
      Caption         =   "Bandwidth"
      Height          =   375
      Index           =   3
      Left            =   12480
      TabIndex        =   158
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEar 
      Caption         =   "Yes / No"
      Height          =   375
      Index           =   2
      Left            =   14880
      TabIndex        =   157
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEar 
      Caption         =   "Sound Type"
      Height          =   375
      Index           =   1
      Left            =   13680
      TabIndex        =   156
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEar 
      Caption         =   "Which Ear"
      Height          =   375
      Index           =   0
      Left            =   12480
      TabIndex        =   155
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer timerClick 
      Enabled         =   0   'False
      Interval        =   350
      Left            =   4800
      Top             =   720
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   120
      TabIndex        =   154
      Top             =   2400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer TimerStep8 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   720
   End
   Begin Project1.soundYesNo soundYesNo1 
      Height          =   3975
      Left            =   2880
      TabIndex        =   153
      Top             =   7800
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7011
   End
   Begin VB.Timer TimerStep6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   720
   End
   Begin VB.Timer TimerStep3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   240
   End
   Begin Project1.soundbandwidthdial soundbandwidthdial1 
      Height          =   4575
      Left            =   2760
      TabIndex        =   151
      Top             =   6600
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   8070
   End
   Begin Project1.dialcontrol dialcontrol1 
      Height          =   3750
      Left            =   6840
      TabIndex        =   150
      Top             =   6000
      Visible         =   0   'False
      Width           =   3750
      _ExtentX        =   8916
      _ExtentY        =   8493
   End
   Begin VB.Timer TimerStep2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4320
      Top             =   240
   End
   Begin VB.Timer TimerStep1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3840
      Top             =   240
   End
   Begin Project1.whicheardial whicheardial1 
      Height          =   4095
      Left            =   4440
      TabIndex        =   149
      Top             =   6600
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7223
   End
   Begin VB.PictureBox imgCheck 
      BackColor       =   &H00F0F0E6&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   12840
      Picture         =   "TinTest.frx":030A
      ScaleHeight     =   3135
      ScaleWidth      =   3375
      TabIndex        =   148
      Top             =   9720
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer timerCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   840
   End
   Begin VB.Timer timerDblClick 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2160
      Top             =   840
   End
   Begin VB.Timer TimerVolume 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2760
      Top             =   240
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   120
      TabIndex        =   144
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtRIRightT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   17160
      TabIndex        =   143
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   17160
      TabIndex        =   142
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   17160
      TabIndex        =   141
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   16680
      TabIndex        =   140
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   16680
      TabIndex        =   139
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   16680
      TabIndex        =   138
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   16080
      TabIndex        =   137
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   16080
      TabIndex        =   136
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRIRightT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   16080
      TabIndex        =   135
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   15600
      TabIndex        =   134
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   15600
      TabIndex        =   133
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtRILeftT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   15600
      TabIndex        =   132
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundLevelMatch 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   6240
      TabIndex        =   94
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundLevelMatch 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   93
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundLevelMatch 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6240
      TabIndex        =   92
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtSoundThreshold 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   91
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   4920
      TabIndex        =   89
      Top             =   15120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   4920
      TabIndex        =   88
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   4920
      TabIndex        =   87
      Top             =   14640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   4920
      TabIndex        =   86
      Top             =   14400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   85
      Top             =   14160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4920
      TabIndex        =   84
      Top             =   13920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   83
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   82
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   81
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   80
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   4440
      TabIndex        =   79
      Top             =   15120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   4440
      TabIndex        =   78
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   4440
      TabIndex        =   77
      Top             =   14640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   4440
      TabIndex        =   76
      Top             =   14400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   75
      Top             =   14160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   74
      Top             =   13920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   73
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   72
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   71
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   70
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   3960
      TabIndex        =   69
      Top             =   15120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   68
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3960
      TabIndex        =   67
      Top             =   14640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   3960
      TabIndex        =   66
      Top             =   14400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   3960
      TabIndex        =   65
      Top             =   14160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3960
      TabIndex        =   64
      Top             =   13920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   63
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   62
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   61
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   60
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT3 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4920
      TabIndex        =   59
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   58
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPitchMatchT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   57
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   3000
      TabIndex        =   56
      Top             =   15120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   55
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   54
      Top             =   14640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   53
      Top             =   14400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   3000
      TabIndex        =   52
      Top             =   14160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3000
      TabIndex        =   51
      Top             =   13920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   50
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   49
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   48
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   47
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   2520
      TabIndex        =   46
      Top             =   15120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2520
      TabIndex        =   45
      Top             =   14880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   8
      Left            =   2520
      TabIndex        =   44
      Top             =   14640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   2520
      TabIndex        =   43
      Top             =   14400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   42
      Top             =   14160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   41
      Top             =   13920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   40
      Top             =   13680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   39
      Top             =   13440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   38
      Top             =   13200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   37
      Top             =   12960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT2 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   36
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtLoudnessT1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   35
      Top             =   12720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtDebug2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   12000
      TabIndex        =   34
      Text            =   "Text2"
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtDebug1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   12000
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.TextBox txtTimer 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   32
      Text            =   "0"
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   240
   End
   Begin VB.TextBox txtLoudness 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   31
      Top             =   14400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtTemporal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   30
      Top             =   14040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtIntensity 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   29
      Top             =   12960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtLocalize 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   28
      Top             =   12600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtBandwidth 
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   27
      Top             =   13680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin PA5XLib.PA5x PA5x2 
      Left            =   1440
      Top             =   720
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin PA5XLib.PA5x PA5x1 
      Left            =   1440
      Top             =   240
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.CheckBox chkChange 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1080
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtValue 
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Text            =   "1"
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox chkClick 
      Enabled         =   0   'False
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Enabled         =   0   'False
      Height          =   495
      Left            =   16320
      TabIndex        =   4
      Top             =   14280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   14280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   11520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   4440
      TabIndex        =   146
      Top             =   6360
      Visible         =   0   'False
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Max             =   30
   End
   Begin VB.Frame frmBegin 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Welcome to the Tinnitus Tester"
      Height          =   3615
      Left            =   3600
      TabIndex        =   5
      Top             =   3240
      Width           =   10575
      Begin VB.ComboBox cboResume 
         Height          =   315
         Left            =   6000
         TabIndex        =   26
         Text            =   "Resume From..."
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.DirListBox dirResume 
         Height          =   1440
         Left            =   3240
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtInitials 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   720
         Width           =   4215
      End
      Begin VB.OptionButton optResume 
         BackColor       =   &H00F0F0E6&
         Caption         =   "Resume Subject:               "
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optNew 
         BackColor       =   &H00F0F0E6&
         Caption         =   "New Subject - Enter Initials:"
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "File Open"
      Filter          =   "TinTest Data (Maindata00.csv)|Maindata00.csv"
   End
   Begin VB.Label lblSon 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E6&
      Caption         =   "(son aigu-son grave)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   44.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   180
      Top             =   1680
      Visible         =   0   'False
      Width           =   16815
   End
   Begin VB.Label lblSoft 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E6&
      Caption         =   "Softer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   168
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblLoud 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E6&
      Caption         =   "Louder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13440
      TabIndex        =   167
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblInstruct3 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Please Enter Subject Initials:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   840
      TabIndex        =   161
      Top             =   5280
      Visible         =   0   'False
      Width           =   16695
   End
   Begin VB.Label lblInstruct2 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Please Enter Subject Initials:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   840
      TabIndex        =   160
      Top             =   3120
      Visible         =   0   'False
      Width           =   16695
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      Index           =   2
      Visible         =   0   'False
      X1              =   24
      X2              =   24
      Y1              =   472
      Y2              =   448
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   600
      X2              =   520
      Y1              =   496
      Y2              =   520
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   29
      X2              =   29
      Y1              =   416
      Y2              =   428
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   32
      X2              =   48
      Y1              =   432
      Y2              =   424
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   8
      X2              =   24
      Y1              =   440
      Y2              =   432
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   136
      X2              =   288
      Y1              =   424
      Y2              =   440
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   8
      X2              =   24
      Y1              =   432
      Y2              =   432
   End
   Begin VB.Line lineClick 
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   8
      X2              =   24
      Y1              =   424
      Y2              =   432
   End
   Begin VB.Label lblVolume 
      Caption         =   "Increasing Volume"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   147
      Top             =   11160
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Image imgVolume 
      Height          =   3240
      Index           =   1
      Left            =   6840
      Picture         =   "TinTest.frx":0F47
      Top             =   8040
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Image imgVolume 
      Height          =   3240
      Index           =   0
      Left            =   6840
      Picture         =   "TinTest.frx":1D52
      Top             =   8040
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VB.Label lblEar 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Adjust Left Ear First..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1800
      TabIndex        =   145
      Top             =   4680
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label lblLeftRight 
      Caption         =   " Left     Right    Left    Right"
      Height          =   255
      Left            =   13440
      TabIndex        =   131
      Top             =   12840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblNextSound 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Starting next sound - begin turning dial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   960
      TabIndex        =   90
      Top             =   6600
      Visible         =   0   'False
      Width           =   16575
   End
   Begin VB.Label lblChoice5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Pulsing Sound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   16
      Top             =   9360
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label lblChoice4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A Steady Sound"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   15
      Top             =   9360
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Shape shpChoice5 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   8400
      Top             =   11160
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Shape shpChoice4 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   1440
      Top             =   8280
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label lblChoice3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Right Ear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      TabIndex        =   13
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblChoice1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Left Ear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   12
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblChoice2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Both Ears"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7440
      TabIndex        =   11
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Shape shpChoice3 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   8400
      Shape           =   1  'Square
      Top             =   10080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape shpChoice2 
      BackColor       =   &H0080FF80&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   6960
      Shape           =   1  'Square
      Top             =   8040
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Shape shpChoice1 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   3255
      Left            =   1680
      Shape           =   1  'Square
      Top             =   8040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E6&
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   16815
   End
   Begin VB.Label lblMainInstructions 
      Alignment       =   2  'Center
      BackColor       =   &H00F0F0E6&
      Caption         =   "Please Enter Subject ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   16575
   End
End
' Form1
' This form represents the main form of the TinTest application.
' It contains various attributes and properties that define the behavior and appearance of the form.
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-----------------------------------------------------------------------------------------------
' Description: This code block contains the necessary declarations and constants for playing sound using the sndPlaySound API call.
'              It also includes public variables used for volume adjustment, dial offset, and other sound-related settings.
'              The constants define various flags for controlling the behavior of the sound playback.
'-----------------------------------------------------------------------------------------------
'First we have to declare the API call
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Then we have to declare the constants that go along with the sndPlaySound function
Const SND_ASYNC = &H1       'ASYNC allows us to play waves with the ability to interrupt
Const SND_LOOP = &H8        'LOOP causes to sound to be continuously replayed
Const SND_NODEFAULT = &H2   'NODEFAULT causes no sound to be played if the wav can't be found
Const SND_SYNC = &H0        'SYNC plays a wave file without returning control to the calling program until it's finished
Const SND_NOSTOP = &H10     'NOSTOP ensures that we don't stop another wave from playing
Const SND_MEMORY = &H4      'MEMORY plays a wave file stored in memory
Public VolAdj As Boolean    'step flag - indicates whether a volume icon should be displayed during adjustment
Public DialOffset As Integer 'this will hold how much to offset the dial from the top.  Needed due to TinTrain & TinTest both using same dial
Public intMaxVolume As Integer
Public SoundTypeDialTop As Integer
Public SoundBandwidthDialTop As Integer
Public WhichEarTop As Integer
Public YesNoTop As Integer


' This event handler is triggered when the "cboResume" combo box is clicked.
' If the selected text in the combo box is "Resume From...", the "cmdNext" button is not enabled.
' Otherwise, the "cmdNext" button is enabled.
Private Sub cboResume_Click()
    If cboResume.Text = "Resume From..." Then 'do not enable next button
    Else
        cmdNext.Enabled = True
    End If
End Sub

' This subroutine is triggered when a button with an index is clicked.
' It toggles the visibility of different controls based on the button index.
Private Sub cmdEar_Click(Index As Integer)
    If Index = 0 Then
        If whicheardial1.Visible Then
            whicheardial1.Visible = False
        Else
            whicheardial1.Visible = True
        End If
    ElseIf Index = 1 Then
        If soundtypedial1.Visible Then
            soundtypedial1.Visible = False
        Else
            soundtypedial1.Visible = True
        End If
    ElseIf Index = 2 Then
        If soundYesNo1.Visible Then
            soundYesNo1.Visible = False
        Else
            soundYesNo1.Visible = True
        End If
    ElseIf Index = 3 Then
        If soundbandwidthdial1.Visible Then
            soundbandwidthdial1.Visible = False
        Else
            soundbandwidthdial1.Visible = True
        End If
    ElseIf Index = 4 Then
        If dialcontrol1.Visible Then
            dialcontrol1.Visible = False
        Else
            dialcontrol1.Visible = True
        End If
    End If
End Sub



' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'---------------------------------------------------------------------------
' Procedure: cmdNext_Click
' Description: This procedure is the event handler for the "Next" button click event. It is responsible for executing the steps of the tinnitus tester based on the user's selection. It also handles the resume functionality, where the user can choose to resume the test from a specific step. The procedure performs various file operations, data manipulation, and calls other subroutines for each step of the tinnitus test.
'---------------------------------------------------------------------------
Private Sub cmdNext_Click()
    ' Variable Declarations
    Dim intfilenumber, t1, intResponse, ReadCounter, ReadCounter2, ReadCounter3, ReadCounter4 As Integer
    Dim TempString, tempstring2, InputArray(200) As String
    Dim ResumeCounter As Integer
    Dim lmdata As Single ' tinnitus loudness in dbspl- to be passed into tinnitus report call
    
    ' Set attenuation level for PA5x1
    PA5x1.SetAtten (50)
    
    ' Check if user is using 2 PA5s and set level for PA5x2
    If usePA52 Then
        PA5x2.SetAtten (50)
    End If
    
    ' Set dial offset for tintester
    DialOffset = 100
    dialcontrol1.Top = (Form1.ScaleHeight / 2) - (dialcontrol1.Height / 2) + DialOffset
    
    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
    ' This code block starts the tinnitus tester from scratch. It performs the following steps:
    ' 1. Checks if the user has entered initials. If not, displays a message and exits.
    ' 2. Creates a working directory and checks if the file already exists. If it does, displays a message and exits.
    ' 3. Creates the working directory and writes the "Recorded On" timestamp to a file.
    ' 4. Calls each step of the tinnitus test, hiding the previous step before calling the next one.
    ' Start tinnitus tester from scratch
    If optNew.Value = True Then
        ' Check if user has entered initials
        If txtInitials.Text = "" Then
            MsgBox "Please Enter Users Intials Before Continuing"
            Exit Sub
        End If
        
        ' Create working directory and check if file already exists
        WorkingDir = "c:\TinData\" & Year(Now) & "_" & Month(Now) & "_" & Day(Now) & "_" & txtInitials
        WorkingFile = "\MainData00.csv"
        If Dir(WorkingDir & WorkingFile) <> "" Then
            MsgBox "File already exists. Please enter a new filename."
            Exit Sub
        End If
        
        ' Create working directory and write recorded on timestamp to file
        MkDir (WorkingDir)
        intfilenumber = FreeFile
        Open (WorkingDir & WorkingFile) For Output As #intfilenumber
        Write #intfilenumber, "Recorded On:", Now
        Close #intfilenumber

        ' This code calls each step of the tinnitus test in a specific order.
        ' It hides all previous steps before calling the next step.
        ' The steps include localizing the tinnitus, determining sound intensity, bandwidth, temporal characteristics,
        ' loudness rating, loudness matching, pitch matching, threshold determination, and residual inhibition.
        ' Each step is called using the "Call" keyword followed by the corresponding subroutine name.
        ' After each step is called, the "hide_all" subroutine is called to hide all elements on the form.
        ' This code is located in the file "TinTest.frm" within the "TinnTester" project.
        Call hide_all
        Call Step1_Localize
        Call hide_all
        Call Step2_SoundIntensity
        Call hide_all
        Call Step3_Bandwidth
        Call hide_all
        Call Step4_Temporal
        Call hide_all
        Call Step5_LoudnessRating
        Call hide_all
        Call Step6_LoudnessMatching
        Call hide_all
        Call Step7_PitchMatching
        Call hide_all
        Call Step8_Threshold
        Call hide_all
        Call Step9_ResidualInhibition
        
    ' Resume functionality
    Else
        ' Warn the user that data will be overwritten
        intResponse = MsgBox("WARNING! All data from and including " & cboResume.Text & " will be overwritten! Press OK to continue, or Cancel to abort.", vbOKCancel)
        
        If intResponse = 1 Then ' User pressed OK
            ' Do nothing, continue with experiment
        Else ' User hit cancel
            Exit Sub
        End If
        
        ' Determine where to start the file name
        lblMainInstructions.Alignment = 0
        c1 = 0
        ResumeCounter = 0
        TempString = dirResume.Path & "\MainData0" & c1 & ".csv"
        Do While dir(TempString) = ("MainData0" & c1 & ".csv")
            c1 = c1 + 1
            TempString = dirResume.Path & "\MainData0" & c1 & ".csv"
        Loop
        
        ' Set working directory and file
        WorkingDir = dirResume.Path
        WorkingFile = "\MainData00.csv"
        
        ' This code reads existing records from a file, overwrites the file with new records, and erases previous records.
        ' It uses a loop to read the existing records from the file and store them in an array called InputArray.
        ' Then, it opens the file for output and uses multiple loops to write the new records to the file.
        ' The number of records to be written in each loop is determined by the values of ReadCounter, ReadCounter2, ReadCounter3, and ReadCounter4.
        ' Finally, it closes the file.
        ' Note: The file path is assumed to be "\MainData00.csv" in the same directory as the form file.
        c1 = 1
        intfilenumber = FreeFile
        Open (dirResume.Path & "\MainData00.csv") For Input As #intfilenumber
        Do While Not EOF(intfilenumber)
            Input #intfilenumber, InputArray(c1)
            c1 = c1 + 1
        Loop
        Close #intfilenumber
        intfilenumber = FreeFile
        Open (dirResume.Path & "\MainData00.csv") For Output As #intfilenumber
        c1 = 1
        Do While c1 <= ReadCounter
            Write #intfilenumber, InputArray(c1), InputArray(c1 + 1)
            c1 = c1 + 2
        Loop
        Do While c1 <= ReadCounter2
            Write #intfilenumber, InputArray(c1), InputArray(c1 + 1), InputArray(c1 + 2), InputArray(c1 + 3)
            c1 = c1 + 4
        Loop
        Do While c1 <= ReadCounter3
            Write #intfilenumber, InputArray(c1), InputArray(c1 + 1), InputArray(c1 + 2), InputArray(c1 + 3), InputArray(c1 + 4)
            c1 = c1 + 5
        Loop
        Do While c1 <= ReadCounter4
            Write #intfilenumber, InputArray(c1), InputArray(c1 + 1)
            c1 = c1 + 2
        Loop
        Close #intfilenumber
        
        ' Call each step of the tinnitus test based on the resume counter
        Do While ResumeCounter <= 9
            Select Case ResumeCounter
                Case Is = 1
                    Call Step1_Localize
                Case Is = 2
                    Call Step2_SoundIntensity
                Case Is = 3
                    Call Step3_Bandwidth
                Case Is = 4
                    Call Step4_Temporal
                Case Is = 5
                    Call Step5_LoudnessRating
                Case Is = 6
                    Call Step6_LoudnessMatching
                Case Is = 7
                    Call Step7_PitchMatching
                Case Is = 8
                    Call Step8_Threshold
                Case Is = 9
                    Call Step9_ResidualInhibition
            End Select
            ResumeCounter = ResumeCounter + 1
        Loop
    End If
    
    ' Display completion message and write tinnitus report
    If English Then
        lblMainInstructions.Caption = "Thank you. The program is now complete. Please wait for the experimenter to enter."
    Else
        lblMainInstructions.Caption = "Merci. Le programme est maintenant terminé. Veuillez avertir l'évaluateur."
    End If
    Call WriteReportSPL
    lblMainInstructions.Visible = True
    
    ' Output tinnitus report
    If PReport Then
        lmdata = CInt(txtPA5ThreshValue.Text) - ((CInt(txtLoudnessT1(1)) + CInt(txtLoudnessT2(1))) / 2)
        If English Then
            Call OutputReport(CInt(txtLoudness.Text), CInt(lmdata), RI5k)
        Else
            Call OutputReport_F(CInt(txtLoudness.Text), CInt(lmdata), RI5k)
        End If
    End If
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'*************************************************************************************
'** Subroutine Name: WriteReportSPL
'**
'** Purpose: This subroutine is responsible for writing a report in CSV format with various SPL (Sound Pressure Level) measurements.
'**
'** Parameters:
'**    - None
'**
'** Returns:
'**    - None
'**
'** Notes:
'**    - This subroutine reads calibration data from a file and uses it to calculate SPL values for different frequencies and tinnitus properties.
'**    - The calculated SPL values are written to a CSV file along with other information such as tinnitus location, temporal property, and loudness rating.
'**    - The file path for the output file is determined by the variable WorkingDir and the file name is "MainDataSPL.csv".
'**    - If the calibration data file does not exist, a message box is displayed indicating that the SPL values will be incorrect.
'**    - The output file is opened in write mode and various information is written to it using the Write statement.
'**    - The calculated SPL values are obtained by subtracting the corresponding values from the calibration data.
'**    - The output file is closed after all the information is written.
'**
'** Example Usage:
'**    WriteReportSPL
'**
'*************************************************************************************
Private Sub WriteReportSPL()
Dim intfilenumber As Integer
Dim TempString1 As String
Dim TinBW As Integer 'tinnitus bandwidth
Dim x1 As Integer
Dim CFreq(0 To 10) As String
Dim cm1 As String
Dim CMF As Integer  'custom masker frequency
CFreq(0) = "0.5 kHz"
CFreq(1) = "1 kHz"
CFreq(2) = "2 kHz"
CFreq(3) = "3 kHz"
CFreq(4) = "4 kHz"
CFreq(5) = "5 kHz"
CFreq(6) = "6 kHz"
CFreq(7) = "7 kHz"
CFreq(8) = "8 kHz"
CFreq(9) = "10 kHz"
CFreq(10) = "12 kHz"

cm1 = CustomMaskerPath() 'returns the file to use for custom masker
If Len(cm1) = 2 Then '1 to 9
    CMF = CInt(Right(cm1, 1))
ElseIf Len(cm1) = 3 Then
    CMF = CInt(Right(cm1, 2)) '10 or 11
Else
    CMF = 1
    MsgBox "ERROR DETERMINING CMF"
End If

        ' This code reads in calibration values from a CSV file.
        ' If the file exists, it reads the values into the CalibData array.
        ' If the file does not exist, it sets all values in the CalibData array to 0 and displays a message.
        ' The CalibData array stores calibration values for pure tone, ringing, hissing, and white noise.
        ' The values are used to calculate sound pressure level (SPL) values.
        ' Note: If the calibration data file is missing, the recorded PA5 values will be correct, but the SPL values will be incorrect.
        ' The file path for the calibration data file is "C:\TinData\CalibrationData.csv".
        'first,we must read in calibration values
        If (dir("C:\TinData\CalibrationData.csv")) = "CalibrationData.csv" Then 'the datafile exists, read it in
            'row 1 = pure tone, row 2 = ringing row 3 = hissing
            intfilenumber = FreeFile
            Open ("C:\TinData\CalibrationData.csv") For Input As #intfilenumber
            For c1 = 1 To 3 Step 1
                For c2 = 1 To 11 Step 1
                    Input #intfilenumber, tempS
                    CalibData(c1, c2) = tempS
                Next c2
            Next c1
            Input #intfilenumber, tempS
            CalibData(4, 1) = tempS 'white noise value
            Close #intfilenumber
        Else
            For c1 = 1 To 3 Step 1
                For c2 = 1 To 11 Step 1
                    CalibData(c1, c2) = 0
                Next c2
            Next c1
            MsgBox "No Calibration Data File Exists.  Recorded PA5 values will be correct, but SPL values will be incorrect"
        End If
        WorkingFile = "\MainDataSPL.csv"
        'If dir(WorkingDir & WorkingFile) <> "" Then 'file exists, create a new file
        '    MsgBox "File already exists.  Please enter a new filename."
        '    Exit Sub
        'End If

        ' This code writes data to a file specified by the WorkingDir and WorkingFile variables.
        ' The data includes information about the recorded date, tinnitus location, temporal property, tinnitus bandwidth,
        ' comfortable listening levels, loudness rating, tinnitus loudness match, tinnitus likeness match, RI results,
        ' and RI sound levels for left and right ears.
        ' The data is written in a specific format to the file using the Write statement.
        ' The file is then closed using the Close statement.
        intfilenumber = FreeFile ' This is safer than assigning a number
        Open (WorkingDir & WorkingFile) For Output As #intfilenumber
            'first, write the date file was recorded on
            Write #intfilenumber, "Recorded On:", Now
            'next, output location
            Select Case CInt(txtLocalize.Text)
            Case Is = 1
                TempString1 = "Localized in Left Ear"
            Case Is = 2
                TempString1 = "Localized in Both Ears"
            Case Is >= 3
                TempString1 = "Localized in Right Ear"
            End Select
            Write #intfilenumber, "Tinnitus Location:", TempString1
            'next temporal property:
            Select Case CInt(txtTemporal.Text)
                Case Is = 1 'Steady Sound
                    TempString1 = "Steady Tinnitus"
                Case Is >= 2 'pulsing Sound
                    TempString1 = "Pusling Tinnitus"
            End Select
            Write #intfilenumber, "Temporal Property:", TempString1
            'next, bandwidth
            Select Case CInt(txtBandwidth.Text)
            Case Is = 1
                TempString1 = "Hissing"
                TinBW = 3
            Case Is = 2
                TempString1 = "Ringing"
                TinBW = 2
            Case Is >= 3
                TempString1 = "Tonal"
                TinBW = 1
            End Select
            Write #intfilenumber, "Tinnitus Bandwidth:", TempString1
            'SPL for comfortable listening levels:
            Write #intfilenumber, "500Hz Comfortable Level:", (CalibData(1, 1) - CInt(txtIntensity.Text)), "dB SPL"
            Write #intfilenumber, "5kHz Comfortable Level:", (CalibData(1, 6) - CInt(txtIntensity.Text)), "dB SPL"
            Write #intfilenumber, "Loudness Rating:", txtLoudness.Text
            'output SPL levels for tinnitus loudness match
            Write #intfilenumber, "TINNITUS LOUDNESS MATCH"
            Write #intfilenumber, "Center Freq", "Trial 1", "Trial 2", "Average"
            For x1 = 0 To 10 Step 1
                Write #intfilenumber, CFreq(x1), (CalibData(TinBW, x1 + 1) - CInt(txtLoudnessT1(x1).Text)), (CalibData(TinBW, x1 + 1) - CInt(txtLoudnessT2(x1).Text)), ((CalibData(TinBW, x1 + 1) - CInt(txtLoudnessT1(x1).Text)) + (CalibData(TinBW, x1 + 1) - CInt(txtLoudnessT2(x1).Text))) / 2, "dB SPL"
            Next x1
            Write #intfilenumber, "TINNITUS LIKENESS MATCH"
            Write #intfilenumber, "Center Freq", "Trial 1", "Trial 2", "Trial 3", "Average"
            For x1 = 0 To 10 Step 1
                Write #intfilenumber, CFreq(x1), CInt(txtPitchMatchT1(x1).Text), CInt(txtPitchMatchT2(x1).Text), CInt(txtPitchMatchT3(x1).Text), (CInt(txtPitchMatchT1(x1).Text) + CInt(txtPitchMatchT2(x1).Text) + CInt(txtPitchMatchT3(x1).Text)) / 3
            Next x1
            Write #intfilenumber, "RI RESULTS"
            Write #intfilenumber, "1kHz Threshold:", (CalibData(1, 2) - CInt(txtPA5ThreshValue.Text)), "dB SPL"
            'Write #intfilenumber, "500 Hz Masker Level:", (CalibData(3, 1) - (CInt(txtSoundLevelMatch(0).Text)))
            Write #intfilenumber, "5000 Hz Masker Level:", (CalibData(3, 6) - (CInt(txtSoundLevelMatch(1).Text)))
            'write #intfilenumber, "White Noise Masker Level:", (CalibData(4, 1) - (CInt(txtSoundLevelMatch(2).Text)))
            'Write #intfilenumber, "Custom Masker " & cm1 & " Level:", (CalibData(TinBW, CMF) - CInt(txtSoundLevelMatch(3).Text))
            Write #intfilenumber, "RI SOUND", "LEFT EAR T1", "RIGHT EAR T1", "LEFT EAR T2", "RIGHT EAR T2"
            'Write #intfilenumber, "500Hz NBN", ((CInt(txtRILeftT1(0).Text) - 50 - 1) / 10), ((CInt(txtRIRightT1(0).Text) - 50 - 1) / 10), ((CInt(txtRILeftT2(0).Text) - 50 - 1) / 10), ((CInt(txtRIRightT2(0).Text) - 50 - 1) / 10)
            Write #intfilenumber, "5000Hz NBN", ((CInt(txtRILeftT1(0).Text) - 50 - 1) / 10), ((CInt(txtRIRightT1(0).Text) - 50 - 1) / 10), ((CInt(txtRILeftT1(1).Text) - 50 - 1) / 10), ((CInt(txtRIRightT1(1).Text) - 50 - 1) / 10)
            'Write #intfilenumber, "White Noise", ((CInt(txtRILeftT1(2).Text) - 50 - 1) / 10), ((CInt(txtRIRightT1(2).Text) - 50 - 1) / 10), ((CInt(txtRILeftT2(2).Text) - 50 - 1) / 10), ((CInt(txtRIRightT2(2).Text) - 50 - 1) / 10)
            'Write #intfilenumber, "Custom Masker: " & cm1, ((CInt(txtRILeftT1(3).Text) - 50 - 1) / 10), ((CInt(txtRIRightT1(3).Text) - 50 - 1) / 10), ((CInt(txtRILeftT2(3).Text) - 50 - 1) / 10), ((CInt(txtRIRightT2(3).Text) - 50 - 1) / 10)
            
        Close #intfilenumber
        
        '
        

End Sub

'******************************************************************************************************
'** Description: This code reads data from a file and populates various text boxes in a form.
'**
'** Inputs:
'**   - comFile.FileName: The path of the file to be read.
'**   - txtLocalize: Text box to display the localization data.
'**   - txtIntensity: Text box to display the sound intensity data.
'**   - txtBandwidth: Text box to display the bandwidth data.
'**   - txtTemporal: Text box to display the temporal data.
'**   - txtLoudness: Text box to display the loudness data.
'**   - txtLoudnessT1: Array of text boxes to display loudness matching data for condition T1.
'**   - txtLoudnessT2: Array of text boxes to display loudness matching data for condition T2.
'**   - txtPitchMatchT1: Array of text boxes to display pitch matching data for condition T1.
'**   - txtPitchMatchT2: Array of text boxes to display pitch matching data for condition T2.
'**   - txtPitchMatchT3: Array of text boxes to display pitch matching data for condition T3.
'**
'** Outputs:
'**   - None
'**
'** Notes:
'**   - The code opens the specified file and reads the data line by line.
'**   - The data is then assigned to the appropriate text boxes based on the line number.
'**   - The code assumes a specific structure of the file and assigns the data accordingly.
'**   - Error handling is implemented to handle any issues with file reading.
'******************************************************************************************************
Private Sub cmdPrintReport_Click()
    Dim intfilenumber, c1, c2 As Integer
    Dim TempString As String
    Dim lmdata As Single
    cmdPrintReport.Caption = "Running..."
    cmdPrintReport.Enabled = False
    cmdNext.Enabled = False
    cmdTinTrain.Enabled = False
    frmBegin.Enabled = False
    DoEvents
    txtLocalize.Text = ""
    txtIntensity.Text = ""
    txtBandwidth.Text = ""
    txtTemporal.Text = ""
    txtLoudness.Text = ""
    For c1 = 0 To 10
        txtLoudnessT1(c1).Text = ""
        txtLoudnessT2(c1).Text = ""
        txtPitchMatchT1(c1).Text = ""
        txtPitchMatchT2(c1).Text = ""
        txtPitchMatchT3(c1).Text = ""
    Next c1
    On Error GoTo comErrorHandler
    intfilenumber = FreeFile
    comFile.InitDir = "C:\TinData"
    comFile.ShowOpen 'open file input box
    'code either continues if user clicks ok or skips to error if they hit cancel
    

        ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
        '*************************************************************************************
        '** This code reads data from a file and populates various text boxes with the data. **
        '** The file is opened for input and each record is read until the end of the file.   **
        '** Depending on the record number, the corresponding text box is updated with the   **
        '** value from the record.                                                           **
        '**                                                                                 **
        '** Input:                                                                          **
        '** - comFile: The file to be opened for input.                                      **
        '** - intfilenumber: The file number used for opening the file.                      **
        '**                                                                                 **
        '** Output:                                                                         **
        '** - Various text boxes are updated with the data from the file.                    **
        '**                                                                                 **
        '** Note:                                                                           **
        '** - The code assumes that the text boxes (txtLocalize, txtIntensity, etc.)        **
        '**   already exist in the form.                                                     **
        '** - The code uses a Select Case statement to determine which text box to update    **
        '**   based on the record number.                                                    **
        '** - The code assumes that the file contains the necessary number of records        **
        '**   corresponding to the text boxes. If the file is missing any records, the       **
        '**   corresponding text boxes will not be updated.                                  **
        '*************************************************************************************
        c1 = 0
        Open comFile.FileName For Input As #intfilenumber
        Do While Not EOF(intfilenumber)
            Input #intfilenumber, TempString
            c1 = c1 + 1
            
            Select Case c1
                ' Update text boxes based on record number
                Case Is = 3 'localize data
                    txtLocalize.Text = TempString
                Case Is = 5 'sound intensity data
                    txtIntensity.Text = TempString
                Case Is = 7 'bandwidth data
                    txtIntensity2.Text = TempString
                Case Is = 9 'bandwidth data
                    txtBandwidth.Text = TempString
                Case Is = 11 'temporal data
                    txtTemporal.Text = TempString
                Case Is = 13 'loudness data
                    txtLoudness.Text = TempString
                Case Is = 19  'loudness matching data
                    txtLoudnessT1(0).Text = TempString
                Case Is = 20
                    txtLoudnessT2(0).Text = TempString
                Case Is = 23  'loudness matching data
                    txtLoudnessT1(1).Text = TempString
                Case Is = 24
                    txtLoudnessT2(1).Text = TempString
                Case Is = 27  'loudness matching data
                    txtLoudnessT1(2).Text = TempString
                Case Is = 28
                    txtLoudnessT2(2).Text = TempString
                Case Is = 31  'loudness matching data
                    txtLoudnessT1(3).Text = TempString
                Case Is = 32
                    txtLoudnessT2(3).Text = TempString
                Case Is = 35  'loudness matching data
                    txtLoudnessT1(4).Text = TempString
                Case Is = 36
                    txtLoudnessT2(4).Text = TempString
                Case Is = 39  'loudness matching data
                    txtLoudnessT1(5).Text = TempString
                Case Is = 40
                    txtLoudnessT2(5).Text = TempString
                Case Is = 43  'loudness matching data
                    txtLoudnessT1(6).Text = TempString
                Case Is = 44
                    txtLoudnessT2(6).Text = TempString
                Case Is = 47  'loudness matching data
                    txtLoudnessT1(7).Text = TempString
                Case Is = 48
                    txtLoudnessT2(7).Text = TempString
                Case Is = 51  'loudness matching data
                    txtLoudnessT1(8).Text = TempString
                Case Is = 52
                    txtLoudnessT2(8).Text = TempString
                Case Is = 55  'loudness matching data
                    txtLoudnessT1(9).Text = TempString
                Case Is = 56
                    txtLoudnessT2(9).Text = TempString
                Case Is = 59  'loudness matching data
                    txtLoudnessT1(10).Text = TempString
                Case Is = 60
                    txtLoudnessT2(10).Text = TempString
                Case Is = 68
                    txtPitchMatchT1(0).Text = TempString
                Case Is = 69
                    txtPitchMatchT2(0).Text = TempString
                Case Is = 70
                    txtPitchMatchT3(0).Text = TempString
                Case Is = 73
                    txtPitchMatchT1(1).Text = TempString
                Case Is = 74
                    txtPitchMatchT2(1).Text = TempString
                Case Is = 75
                    txtPitchMatchT3(1).Text = TempString
                Case Is = 78
                    txtPitchMatchT1(2).Text = TempString
                Case Is = 79
                    txtPitchMatchT2(2).Text = TempString
                Case Is = 80
                    txtPitchMatchT3(2).Text = TempString
                Case Is = 83
                    txtPitchMatchT1(3).Text = TempString
                Case Is = 84
                    txtPitchMatchT2(3).Text = TempString
                Case Is = 85
                    txtPitchMatchT3(3).Text = TempString
                Case Is = 88
                    txtPitchMatchT1(4).Text = TempString
                Case Is = 89
                    txtPitchMatchT2(4).Text = TempString
                Case Is = 90
                    txtPitchMatchT3(4).Text = TempString
                Case Is = 93
                    txtPitchMatchT1(5).Text = TempString
                Case Is = 94
                    txtPitchMatchT2(5).Text = TempString
                Case Is = 95
                    txtPitchMatchT3(5).Text = TempString
                Case Is = 98
                    txtPitchMatchT1(6).Text = TempString
                Case Is = 99
                    txtPitchMatchT2(6).Text = TempString
                Case Is = 100
                    txtPitchMatchT3(6).Text = TempString
                Case Is = 103
                    txtPitchMatchT1(7).Text = TempString
                Case Is = 104
                    txtPitchMatchT2(7).Text = TempString
                Case Is = 105
                    txtPitchMatchT3(7).Text = TempString
                Case Is = 108
                    txtPitchMatchT1(8).Text = TempString
                Case Is = 109
                    txtPitchMatchT2(8).Text = TempString
                Case Is = 110
                    txtPitchMatchT3(8).Text = TempString
                Case Is = 113
                    txtPitchMatchT1(9).Text = TempString
                Case Is = 114
                    txtPitchMatchT2(9).Text = TempString
                Case Is = 115
                    txtPitchMatchT3(9).Text = TempString
                Case Is = 118
                    txtPitchMatchT1(10).Text = TempString
                Case Is = 119
                    txtPitchMatchT2(10).Text = TempString
                Case Is = 120
                    txtPitchMatchT3(10).Text = TempString
                Case Is = 123
                    txtSoundThreshold(0).Text = TempString
                Case Is = 125
                    txtSoundThreshold(1).Text = TempString
                Case Is = 127
                    txtSoundThreshold(2).Text = TempString
                Case Is = 129
                    txtSoundThreshold(3).Text = TempString
                Case Is = 131
                    txtPA5ThreshValue.Text = TempString
                Case Is = 133
                    txtSoundLevelMatch(0).Text = TempString
                    txtSoundLevelMatch(1).Text = TempString
                    txtSoundLevelMatch(2).Text = TempString
                    txtSoundLevelMatch(3).Text = TempString
                'Case Is = 135
                '    txtSoundLevelMatch(1).Text = TempString
                'Case Is = 137
                '    txtSoundLevelMatch(2).Text = TempString
                'Case Is = 139
                '    txtSoundLevelMatch(3).Text = TempString
                'RI DATA
                Case Is = 141
                    txtRILeftT1(0).Text = TempString
                Case Is = 142
                    txtRIRightT1(0).Text = TempString
                Case Is = 144
                    txtRILeftT1(1).Text = TempString
                Case Is = 145
                    txtRIRightT1(1).Text = TempString
                Case Is = 147
                    txtRILeftT1(2).Text = TempString
                Case Is = 148
                    txtRIRightT1(2).Text = TempString
                Case Is = 150
                    txtRILeftT1(3).Text = TempString
                Case Is = 151
                    txtRIRightT1(3).Text = TempString
                Case Is = 153
                    txtRILeftT2(0).Text = TempString
                Case Is = 154
                    txtRIRightT2(0).Text = TempString
                Case Is = 156
                    txtRILeftT2(1).Text = TempString
                Case Is = 157
                    txtRIRightT2(1).Text = TempString
                Case Is = 159
                    txtRILeftT2(2).Text = TempString
                Case Is = 160
                    txtRIRightT2(2).Text = TempString
                Case Is = 162
                    txtRILeftT2(3).Text = TempString
                Case Is = 163
                    txtRIRightT2(3).Text = TempString
            End Select
        Loop
        Close #intfilenumber
        
        '******************************************************************************************************
        '** Description: This code calculates various values based on user input and fills in the user values for a report.
        '**
        '** Inputs:
        '**   - txtPA5ThreshValue: The threshold value for loudness match at 1kHz.
        '**   - txtLoudnessT1: Array of loudness values for condition T1.
        '**   - txtLoudnessT2: Array of loudness values for condition T2.
        '**   - txtLocalize: The localization value for tinnitus (1 = left ear, 2 = both ears, 3 = right ear).
        '**   - txtRILeftT1: Array of RI values for left ear in condition T1.
        '**   - txtRIRightT1: Array of RI values for right ear in condition T1.
        '**   - txtRILeftT2: Array of RI values for left ear in condition T2.
        '**   - txtRIRightT2: Array of RI values for right ear in condition T2.
        '**   - txtBandwidth: The bandwidth value (1 = hissing, 2 = ringing, >=3 = tonal).
        '**   - English: Boolean value indicating whether the report should be in English or another language.
        '**
        '** Outputs:
        '**   - lmdata: Loudness match data at 1kHz.
        '**   - RI5k: RI data at 5kHz.
        '**   - UserBW: User-defined bandwidth value for the report.
        '**
        '** Notes:
        '**   - The code calculates the loudness match data at 1kHz based on the threshold value and the average of loudness values.
        '**   - The code calculates the RI data at 5kHz based on the localization value and the average of RI values.
        '**   - The code fills in the user-defined bandwidth value for the report based on the bandwidth input.
        '******************************************************************************************************
                '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                'calculate loudness match data at 1khz
                lmdata = CInt(txtPA5ThreshValue.Text) - ((CInt(txtLoudnessT1(1)) + CInt(txtLoudnessT2(1))) / 2)
                'calculate ri data @ 5khz
                'MsgBox (CInt(txtLoudnessT1(1)) & " " & CInt(txtLoudnessT2(1)))
                If CInt(txtLocalize.Text) = 2 Then   'tinnitus is in both ears
                    RI5k = (CInt(txtRILeftT1(1)) + CInt(txtRIRightT1(1)) + CInt(txtRILeftT2(1)) + CInt(txtRIRightT2(1))) / 4
                ElseIf CInt(txtLocalize.Text) = 3 Then 'tinnitus is in right ear only
                    RI5k = (CInt(txtRIRightT1(1)) + CInt(txtRIRightT2(1))) / 2
                Else 'tinnitus is in left ear only, txtLocalize.Text = 1
                    RI5k = (CInt(txtRILeftT1(1)) + CInt(txtRILeftT2(1))) / 2
                End If
                ' fill in user values for report
                Select Case CInt(txtBandwidth.Text)
                    Case Is = 1
                        If English Then
                            UserBW = "Hissing"
                        Else
                            UserBW = "Stifflement"
                        End If
                    Case Is = 2
                        If English Then
                            UserBW = "Ringing"
                        Else
                            UserBW = "Sonnerie"
                        End If
                    Case Is >= 3
                        If English Then
                            UserBW = "Tonal"
                        Else
                            UserBW = "Tonal"
                        End If
                  End Select

        ' This Select Case statement assigns a value to the variable UserTL based on the value of txtLocalize.Text.
        ' If txtLocalize.Text is 1, UserTL is assigned "Left Ear" if English is True, otherwise "Oreille Gauche".
        ' If txtLocalize.Text is 2, UserTL is assigned "Both Ears" if English is True, otherwise "Deux Oreilles".
        ' If txtLocalize.Text is greater than or equal to 3, UserTL is assigned "Right Ear" if English is True, otherwise "Oreille Droite".
        Select Case CInt(txtLocalize.Text)
            Case Is = 1
                If English Then
                    UserTL = "Left Ear"
                Else
                    UserTL = "Oreille Gauche"
                End If
            Case Is = 2
                If English Then
                    UserTL = "Both Ears"
                Else
                    UserTL = "Deux Oreilles"
                End If
            Case Is >= 3
                If English Then
                    UserTL = "Right Ear"
                Else
                    UserTL = "Oreille Droite"
                End If
        End Select
    
                ' This Select Case statement determines the value of the variable txtTemporal.Text and assigns a corresponding value to the variable UserSorP.
                ' If the value of txtTemporal.Text is 1, the variable UserSorP is assigned the value "Steady" (or "Continu" if the language is not English).
                ' If the value of txtTemporal.Text is greater than or equal to 2, the variable UserSorP is assigned the value "Pulsing" (or "Pulsatif" if the language is not English).
                Select Case CInt(txtTemporal.Text)
                    Case Is = 1 'Steady Sound
                        If English Then
                            UserSorP = "Steady"
                        Else
                            UserSorP = "Continu"
                        End If
                    Case Is >= 2 'pulsing Sound
                        If English Then
                            UserSorP = "Pulsing"
                        Else
                            UserSorP = "Pulsatif"
                        End If
                End Select
                

        'tell output report where to put file:
        WorkingDir = Left(comFile.FileName, (Len(comFile.FileName) - Len(comFile.FileTitle) - 1)) 'removes filename from path and \
        'output report:
        If English Then
            Call OutputReport(CInt(txtLoudness.Text), CInt(lmdata), RI5k)
        Else
            Call OutputReport_F(CInt(txtLoudness.Text), CInt(lmdata), RI5k)
        End If
        MsgBox ("Report output complete")
        cmdTinTrain.Enabled = True
        cmdPrintReport.Caption = "Create Report"
        cmdPrintReport.Enabled = True
        cmdNext.Enabled = True
        frmBegin.Enabled = True
        Form1.SetFocus
comErrorHandler:
    'do nothing - user hit cancel

End Sub

'---------------------------------------------------------------------------
' Procedure:   cmdTinTrain_Click
' Description: This procedure is the click event handler for the "cmdTinTrain" button.
'              It initializes the necessary settings and displays instructions for the tinnitus testing program.
'---------------------------------------------------------------------------
Private Sub cmdTinTrain_Click()
    ' Set attenuation level for PA5x1
    PA5x1.SetAtten (50)
    
    ' Set dial offset
    DialOffset = 222
    
    ' Position dial control in the center of the form
    dialcontrol1.Top = (Form1.ScaleHeight / 2) - (dialcontrol1.Height / 2) + DialOffset
    
    ' Hide all labels and instructions
    Call hide_all
    
    ' Set labels and instructions based on language selection
        lblTitle.Caption = "Welcome"
        lblMainInstructions.Caption = "This program will introduce you to the computerized testing program for tinnitus."
        lblInstruct2.Caption = "By following instructions and working slowly through this introduction, you will learn how to respond to automated cues from the computer."
        lblInstruct3.Caption = "When you are ready to begin, please press the dial."
        lblInstruct2.Top = 256
        lblInstruct2.Height = 150
        lblInstruct3.Top = 420
    
    ' Set alignment of main instructions label to right justify
    lblMainInstructions.Alignment = 0
    
    ' Show labels and instructions
    lblMainInstructions.Visible = True
    lblInstruct2.Visible = True
    lblInstruct3.Visible = True
    lblTitle.Visible = True
    
    ' Reset checkbox value and set focus to the form
    chkClick.Value = 0
    Form1.SetFocus
    
    ' Wait until the user clicks the knob
    Do While (chkClick.Value = 0)
        DoEvents
    Loop
    
    ' Hide unnecessary labels and reset checkbox value
    lblTitle.Visible = False
    lblMainInstructions.Visible = False
    lblInstruct2.Visible = False
    lblInstruct3.Visible = False
    chkClick.Value = 0
    
    ' Reset positions and sizes of labels and instructions
    lblInstruct2.Top = 208
    lblInstruct2.Height = 113
    lblInstruct3.Top = 328
    
    ' Call subroutines for each step of the tinnitus training program
    Call TinTrain1_DialChoice
    Call TinTrain2_Loudness
    Call TinTrain3_pitch
    Call TinTrain4_HSlide
    Call TinTrain5_VSlide
    Call TinTrain6_Complete
    
    ' Show the "cmdTinTrain" button
    cmdTinTrain.Visible = True
End Sub

Private Sub Command1_Click()
    'This uses the constants ASYNC and NODEFAULT to play the sound
    'sndPlaySound App.Path & "\sound.wav", SND_ASYNC Or SND_NODEFAULT
    sndPlaySound "c:\davet\650am.wav", SND_ASYNC Or SND_NODEFAULT
End Sub

Private Sub Command2_Click()
    'sndPlaySound "C:\davet\silence.wav", SND_ASYNC Or SND_NODEFAULT
    Call WriteReportSPL
End Sub

'---------------------------------------------------------------------------
' hide_all
'
' Description: Hides all the controls and elements on the form.
'
'---------------------------------------------------------------------------
Public Sub hide_all()
    cmdNext.visible = False
    cmdTinTrain.visible = False
    frmBegin.visible = False
    lblMainInstructions.visible = False
    lblTitle.visible = False
    txtInitials.visible = False
    chkClick.Value = 0
    chkChange.Value = 0
    shpChoice1.visible = False
    shpChoice2.visible = False
    shpChoice3.visible = False
    shpChoice4.visible = False
    shpChoice5.visible = False
    lblChoice1.visible = False
    lblChoice2.visible = False
    lblChoice3.visible = False
    lblChoice4.visible = False
    lblChoice5.visible = False
    framLoudness.visible = False
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    txtValue.Text = 1
    cmdPrintReport.visible = False
End Sub

'---------------------------------------------------------------------------
' Procedure: TinTrain1_DialChoice
' Description: This procedure handles the functionality of selecting options using a dial.
'              It displays instructions in English or French based on the language setting.
'              The user can turn the dial to change the selection and press it down to proceed.
'---------------------------------------------------------------------------
Public Sub TinTrain1_DialChoice()
    lblMainInstructions.Caption = "By turning the dial, you can select a different option below."
    lblInstruct2.Caption = "Try turning the dial and changing the selection.  Gentle movements are all you need because the dial is very sensitive."
    lblInstruct3.Caption = "When you are comfortable with this, press the dial down to move to the next step."

    lblInstruct2.Top = 256 'these need to be bumped down a bit due to two lines in maininstruct
    lblInstruct2.Height = 150 '113
    lblInstruct3.Top = 420  '376

lblMainInstructions.visible = True
lblInstruct2.visible = True
lblInstruct3.visible = True
chkClick.Value = 0
Choice1231.UserControl_Initialize
Choice1231.visible = True
Choice1231.SetFocus
Do While chkClick.Value = 0
    DoEvents
Loop
lblMainInstructions.visible = False
lblInstruct2.visible = False
lblInstruct3.visible = False
Choice1231.visible = False
lblInstruct2.Top = 208
lblInstruct2.Height = 113
lblInstruct3.Top = 328
End Sub

'*******************************************************************************
'**  FUNCTION NAME: TinTrain2_Loudness
'**
'**  DESCRIPTION: This subroutine adjusts the loudness of a sound using a dial control.
'**               It displays instructions in English or French based on the language setting.
'**               The user can turn the dial clockwise to increase the loudness and
'**               counter-clockwise to decrease the loudness. Once the user is comfortable,
'**               they can press the dial to move to the next step.
'**
'**  PARAMETERS:
'**      None
'**
'**  RETURNS:
'**      None
'**
'**  EXAMPLES:
'**      Call TinTrain2_Loudness
'**
'**  NOTES:
'**      - This subroutine assumes the existence of the following controls on the form:
'**          - lblMainInstructions: Label control to display main instructions
'**          - lblInstruct2: Label control to display additional instructions
'**          - lblInstruct3: Label control to display further instructions
'**          - txtValue: TextBox control to display and set the volume value
'**          - dialcontrol1: Dial control to adjust the volume
'**          - PA5x1: Object representing the first PA5 device
'**          - PA5x2: Object representing the second PA5 device (optional)
'**          - lblSoft: Label control to display "Softer"
'**          - lblLoud: Label control to display "Louder"
'**          - chkClick: CheckBox control to detect when the user clicks the knob
'**          - TimerStep2: Timer control to handle timing events
'**          - sndPlaySound: Function to play sound files
'**
'*******************************************************************************
Public Sub TinTrain2_Loudness()

    lblMainInstructions.Caption = "Turning the dial will adjust the LOUDNESS of a sound."
    lblInstruct2.Caption = "Turn the dial clockwise until you hear a sound. Try turning the sound up and down to get a feel for changing the LOUDNESS with the dial."
    lblInstruct3.Caption = "When you are comfortable with this, press the dial to move to the next step."

    lblInstruct2.Top = 215 'these need to be bumped down a bit due to two lines in maininstruct
    lblInstruct2.Height = 150 '113
    lblInstruct3.Top = 380  '376

lblMainInstructions.visible = True
lblInstruct2.visible = True
lblInstruct3.visible = True
VolAdj = True
intMaxVolume = 0
txtValue.Text = 30
dialcontrol1.UserControl_Initialize
dialcontrol1.setvolume (CInt(txtValue.Text))
        'set inital PA5 value to 90 here (pa5value = 120 - cint(txtvalue.text))
PA5x1.SetAtten (120 - CInt(txtValue.Text))
lblSoft.Left = (Form1.ScaleWidth / 2) - (lblSoft.Width / 2) - 200
lblSoft.Top = (Form1.ScaleHeight / 2) - (lblSoft.Height / 2) + 200
lblLoud.Left = (Form1.ScaleWidth / 2) - (lblLoud.Width / 2) + 200
lblLoud.Top = (Form1.ScaleHeight / 2) - (lblLoud.Height / 2) + 200

    lblSoft.Caption = "Softer"
    lblLoud.Caption = "Louder"

lblSoft.visible = True
lblLoud.visible = True
'play 500hz sound on loop
sndPlaySound "C:\TinData\tintest_wav\s1.wav", SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
dialcontrol1.visible = True
dialcontrol1.SetFocus
TimerStep2.Enabled = True
chkClick.Value = 0

Do While chkClick.Value = 0  'wait until the user clicks the knob
    DoEvents
Loop
sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
TimerStep2.Enabled = False
lblMainInstructions.visible = False
lblInstruct2.visible = False
lblInstruct3.visible = False
dialcontrol1.visible = False
lblSoft.visible = False
lblLoud.visible = False
lblInstruct2.Top = 208
lblInstruct2.Height = 113
lblInstruct3.Top = 328
End Sub

'---------------------------------------------------------------------------
' TinTrain3_pitch
'---------------------------------------------------------------------------
' Description:
'   This subroutine is responsible for training the user to differentiate sounds based on pitch.
'   It sets the instructions and captions based on the language preference.
'   It initializes the PitchControl user control and handles the user interaction with the control.
'   It plays different sounds based on the user's selection and stops playing when the user clicks the knob.
'
' Parameters:
'   None
'
' Returns:
'   None
'
'---------------------------------------------------------------------------
Public Sub TinTrain3_pitch()
    lblMainInstructions.Caption = "The sounds you hear below vary in PITCH."
    lblInstruct2.Caption = "Turn the dial to hear the different sounds. Try turning the dial left and right to get a feel for the differences in PITCH of each sound."
    lblInstruct3.Caption = "When you have listened to each sound a few times, press the dial to move to the next step."
    lblInstruct3.Top = 380
    lblInstruct2.Top = 215 'these need to be bumped down a bit due to two lines in maininstruct
    lblInstruct2.Height = 150 '113
    
    
    PA5x1.SetAtten (35) 'set PA5 at reasonable level
    If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
        PA5x2.SetAtten (35)
    End If
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    PitchControl1.UserControl_Initialize
    PitchControl1.visible = True
    PitchControl1.SetFocus
    chkClick.Value = 0
    chkChange.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            txtValue.Text = CInt(PitchControl1.getvalue) + 1 'call the usercontrol and updated textbox
            Select Case CInt(txtValue.Text)
                Case Is = 1
                    'play lowpitch sound - 1khz
                    'CHANGED ON JAN 26-Request of LER to 250Hz
                    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop playing old sound
                    sndPlaySound "C:\TinData\TinTestTrainer\250_1s_48k.wav", SND_ASYNC Or SND_NODEFAULT 'start playing new sound
                Case Is = 2
                    'play medium pitch sound - 2khz
                    'CHANGED ON JAN 26-Request of LER to 500Hz
                    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop playing old sound
                    sndPlaySound "C:\TinData\TinTestTrainer\500_1s_48k.wav", SND_ASYNC Or SND_NODEFAULT 'start playing new sound
                Case Is >= 3
                    'play high pitch sound - 5khz
                    'CHANGED ON JAN 26-Request of LER to 2kHz
                    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop playing old sound
                    sndPlaySound "C:\TinData\TinTestTrainer\2k_1s_48k.wav", SND_ASYNC Or SND_NODEFAULT 'start playing new sound
                    txtValue.Text = 3
            End Select
        End If
        DoEvents
    Loop
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT

    lblMainInstructions.visible = False
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    PitchControl1.visible = False
    
    lblInstruct2.Top = 208
    lblInstruct2.Height = 113
    lblInstruct3.Top = 328
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'
' TinTrain4_HSlide
'
' This subroutine controls the behavior of a horizontal slider in the TinTest form.
' It displays instructions in either English or French, allows the user to adjust the slider,
' and changes the color of certain labels and frames when the user interacts with the slider.
'
' Parameters:
'   None
'
' Returns:
'   None
'
' Example usage:
'   TinTrain4_HSlide
'
Public Sub TinTrain4_HSlide()
    lblMainInstructions.Caption = "By turning the dial, you can control the slider below."
    lblInstruct2.Caption = "Move the slider all the way back and forth a few times to get a feel for it."
    lblInstruct3.Caption = "When you're comfortable, press the dial to move on."
    framLoudness.Caption = "Horizontal Slider"

    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    lbl5.visible = False    'hide labels
    lbl30.visible = False
    lbl50.visible = False
    lbl70.visible = False
    lbl95.visible = False
    framLoudness.visible = True
    Form1.SetFocus
    txtValue.Text = 1
    chkChange.Value = 0
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            If CInt(txtValue.Text) < 100 Then
                If CInt(txtValue.Text) > 1 Then
                    hscrScale.Value = CInt(txtValue.Text)
                Else
                    hscrScale.Value = 1
                    txtValue.Text = 1
                End If
            Else
                hscrScale.Value = 100
                txtValue.Text = 100
            End If
        End If
        DoEvents
    Loop
    '******test changing colour for button press
    framLoudness.BackColor = &HC0FFC0
    lbl5.BackColor = &HC0FFC0
    lbl30.BackColor = &HC0FFC0    'original color:  &H00F0F0E6&
    lbl50.BackColor = &HC0FFC0
    lbl70.BackColor = &HC0FFC0
    lbl95.BackColor = &HC0FFC0
    lblScale.BackColor = &HC0FFC0
    timerClick.Enabled = True
    Do While timerClick.Enabled
        DoEvents
    Loop
    lblMainInstructions.visible = False
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    framLoudness.visible = False
    If English Then
        framLoudness.Caption = "Loudness Rating"
    Else
        framLoudness.Caption = "Intensit�"
    End If
    lbl5.visible = True
    lbl30.visible = True
    lbl50.visible = True
    lbl70.visible = True
    lbl95.visible = True
    
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'------------------------------------------------------------------------------
' TinTrain5_VSlide
'------------------------------------------------------------------------------
' Description: This subroutine is responsible for handling the functionality of the vertical slider in the TinTest form.
'              It sets the captions of various labels based on the language selected, hides and shows certain labels,
'              sets the initial values of checkboxes and textboxes, adjusts the position of the form and slider,
'              waits for the user to click the knob, handles user adjustments, changes the color of certain labels,
'              enables a timer, and finally hides and shows certain labels again.
'
' Parameters: None
'
' Returns:    None
'
' Notes:      - This subroutine assumes the existence of various form controls such as labels, checkboxes, textboxes, and timers.
'             - The behavior of this subroutine may vary depending on the values of certain form controls and variables.
'             - The language selection is determined by the value of the "English" variable.
'------------------------------------------------------------------------------
Public Sub TinTrain5_VSlide()
    lblMainInstructions.Caption = "The dial can also move the slider up and down."
    lblInstruct2.Caption = "Turn the dial to move the slider up and down a few times to get a feel for it."
    lblInstruct3.Caption = "When you're comfortable, press the dial to move on."

    lblSofter(0).visible = False 'hide labels
    lblGone(0).visible = False
    lblNoChange(0).visible = False
    lblLouder(0).visible = False
    lblMuchLouder(0).visible = False
    Form1.SetFocus
    txtValue.Text = 1
    chkChange.Value = 0
    chkClick.Value = 0
    frmMono(0).Caption = "Vertical Slider"
    frmMono(0).ForeColor = &H0&        'change text to Black
        lblMainInstructions.Top = 20
        lblInstruct2.Top = 80
        lblInstruct3.Top = 200
        frmMono(0).Left = (Form1.ScaleWidth / 2) - (frmMono(0).Width / 2)
        frmMono(0).Top = (Form1.ScaleHeight / 2) - (frmMono(0).Height / 2) + 120 'put it lower than center so it doesn't cover up text
    VScroll1(0).Value = 51
    txtValue.Text = 51
    frmMono(0).visible = True
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            If (101 - CInt(txtValue.Text)) < 100 Then  'if scroll bar greater than 1 (-5)
                If (101 - CInt(txtValue.Text)) > 0 Then 'if scroll bar is less than 101 (+5)
                    VScroll1(0).Value = 102 - CInt(txtValue.Text)
                Else
                    VScroll1(0).Value = 1
                    txtValue.Text = 101
                End If
            Else 'set scroll bar to -5
                VScroll1(0).Value = 101
                txtValue.Text = 1
            End If
        End If
        DoEvents
    Loop

    '******change colour for button press
    lbl0(0).BackColor = &HC0FFC0
    lbl1(0).BackColor = &HC0FFC0
    lbl2(0).BackColor = &HC0FFC0
    lbl3(0).BackColor = &HC0FFC0
    lb4(0).BackColor = &HC0FFC0
    lblFive(0).BackColor = &HC0FFC0
    lblN1(0).BackColor = &HC0FFC0
    lblN2(0).BackColor = &HC0FFC0
    lblN3(0).BackColor = &HC0FFC0
    lblN4(0).BackColor = &HC0FFC0
    lblN5(0).BackColor = &HC0FFC0
    lblSofter(0).BackColor = &HC0FFC0
    lblGone(0).BackColor = &HC0FFC0
    lblNoChange(0).BackColor = &HC0FFC0
    lblLouder(0).BackColor = &HC0FFC0
    lblMuchLouder(0).BackColor = &HC0FFC0
    frmMono(0).BackColor = &HC0FFC0

    timerClick.Enabled = True
    Do While timerClick.Enabled
        DoEvents
    Loop
    lblMainInstructions.visible = False
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    frmMono(0).visible = False
    lblSofter(0).visible = True
    lblGone(0).visible = True
    lblNoChange(0).visible = True
    lblLouder(0).visible = True
    lblMuchLouder(0).visible = True
    
End Sub

'*******************************************************************************
' FUNCTION NAME: TinTrain6_Complete
' DESCRIPTION:   This subroutine is called when the training program is complete.
'                It updates the captions of the main instructions and the second instruction labels based on the language setting.
'                It sets the visibility of the main instructions and the second instruction labels.
'                It initializes the TinTrainComplete variable to False.
'                It sets the focus to Form1.
'                It waits until the user clicks on a button.
'                It updates the caption and alignment of the main instructions label.
'                It sets the visibility of the second instruction label, the begin form, the next button, and the initials textbox.
' PARAMETERS:    None
' RETURNS:       None
'*******************************************************************************
Public Sub TinTrain6_Complete()
        lblMainInstructions.Caption = "Congratulations, the training program is now complete!"
        lblInstruct2.Caption = "If you have any questions about the program you've just finished, please ask the experimenter."
    
    lblMainInstructions.Visible = True
    lblInstruct2.Visible = True
    'chkClick.Value = 0
    TinTrainComplete = False
    Form1.SetFocus
    Do While (TinTrainComplete = False)  'wait until the user clicks c
        DoEvents
    Loop
    lblMainInstructions.Caption = "Please Enter Subject Initials:"
    lblMainInstructions.Alignment = 2
    lblInstruct2.Visible = False
    frmBegin.Visible = True
    cmdNext.Visible = True
    txtInitials.Visible = True
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'-----------------------------------------------------------------------------------------------------------------------
' Procedure: Step1_Localize
' Purpose:   This procedure is responsible for localizing the tinnitus sensation in the user's ear.
'            It displays instructions and options on the form, waits for user input, and records the selected ear.
'-----------------------------------------------------------------------------------------------------------------------
Public Sub Step1_Localize()
    Dim intfilenumber As Integer
    Dim TempString As String
    
    Call Form_Resize  'fix formating after training has moved it all around.

        lblTitle.Caption = "Welcome"
        lblMainInstructions.Caption = "This program will test your tinnitus sensation."
        lblMainInstructions.Alignment = 0 'right justify
        lblInstruct2.Caption = "Instructions will appear on the screen to guide you.  Ask the experimenter if you need assistance."
        lblInstruct3.Caption = "When you are ready to begin, please press the dial."

        lblTitle.Caption = "Bienvenue"
        lblMainInstructions.Caption = "Ce programme va examiner vos sensations d'acouph�ne."
        lblMainInstructions.Alignment = 0 'right justify
        lblInstruct2.Caption = "Des instructions vont apparaitrent sur l'�cran pour vous guider. Demandez � l'�valuateur si vous avez besoin d'aide."
        lblInstruct3.Caption = "Quand vous �tes pr�t � commencer, pressez la commande rotative."

    
    If Form2.opt1024 Then '1024 mode.  Move text around.
        lblMainInstructions.Top = 144
        lblInstruct2.Top = 208
        lblInstruct3.Top = 352
    End If
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    lblTitle.visible = True
    Form1.SetFocus
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    lblTitle.visible = False
    lblInstruct2.visible = False 'hide 2nd instructino line
    lblInstruct3.visible = False 'hide 3rd instructino line
    chkClick.Value = 0
        lblMainInstructions.Caption = "Which ear is your tinnitus coming from? "
        lblInstruct2.Caption = "Turn the dial to select one of the options, then press when you've made your selection."
    lblInstruct2.visible = True

    whicheardial1.UserControl_Initialize
    whicheardial1.visible = True
    whicheardial1.SetFocus
    
    txtValue.Text = 2
    chkChange.Value = 0
    TimerStep1.Enabled = True
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    TimerStep1.Enabled = False
    txtValue.Text = CInt(whicheardial1.getvalue) - 99
    Select Case CInt(txtValue.Text)
        Case Is = 1
            TempString = "Localized in Left Ear"
            t1 = 1

                UserTL = "Left Ear"
        Case Is = 2
            TempString = "Localized in Both Ears"
            t1 = 2
            UserTL = "Both Ears"
        Case Is >= 3
            TempString = "Localized in Right Ear"
            t1 = 3
                UserTL = "Right Ear"
                
    End Select
    whicheardial1.visible = False
    lblInstruct2.visible = False
    'MsgBox TempString
    'output info
    txtLocalize.Text = t1  'set textbox to store info for interprogram use
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, t1, TempString
    Close #intfilenumber
End Sub

'*******************************************************************************
'**  FUNCTION NAME: Step2_SoundIntensity
'**
'**  DESCRIPTION: This subroutine is responsible for adjusting the sound intensity
'**               using a dial control. It presents two different tones (500 Hz and
'**               5000 Hz) and allows the user to adjust the volume until it reaches
'**               a comfortable level. The comfortable volume levels are recorded
'**               and appended to a file.
'**
'**  PARAMETERS:
'**      None
'**
'**  RETURNS:
'**      None
'**
'**  EXAMPLES:
'**      Call Step2_SoundIntensity
'**
'**  NOTES:
'**      - This subroutine assumes that the necessary form controls (lblMainInstructions,
'**        lblInstruct2, lblSoft, lblLoud, txtValue, dialcontrol1, chkClick, txtIntensity,
'**        txtIntensity2) and objects (PA5x1, PA5x2) are properly initialized and available.
'**      - The sound files used for the tones (s1.wav and s6.wav) are assumed to be located
'**        at the specified file paths.
'**      - The WorkingDir and WorkingFile variables are assumed to be defined and contain
'**        the appropriate file path and file name for appending the volume level information.
'*******************************************************************************
Private Sub Step2_SoundIntensity()
    Dim intfilenumber As Integer
        lblMainInstructions.Caption = "Turning the dial will adjust the loudness of a sound.  "
        lblInstruct2.Caption = "Please turn the dial until the sound is at a comfortable level, then press to continue."
        lblSoft.Caption = "Softer"
        lblLoud.Caption = "Louder"

    If Form2.opt1024 Then '1024 mode.  Move text around.
        lblMainInstructions.Top = 144
        lblInstruct2.Top = 208
        lblInstruct3.Top = 352
    End If


    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    
    lblSoft.visible = True
    lblLoud.visible = True
    VolAdj = True
    intMaxVolume = 0
    txtValue.Text = 30
    dialcontrol1.UserControl_Initialize
    dialcontrol1.setvolume (CInt(txtValue.Text))
    lblSoft.Left = (Form1.ScaleWidth / 2) - (lblSoft.Width / 2) - 200
    lblSoft.Top = (Form1.ScaleHeight / 2) - (lblSoft.Height / 2) + 55
    lblLoud.Left = (Form1.ScaleWidth / 2) - (lblLoud.Width / 2) + 200
    lblLoud.Top = (Form1.ScaleHeight / 2) - (lblLoud.Height / 2) + 55

    'set inital PA5 value to 90 here (pa5value = 120 - cint(txtvalue.text))
        PA5x1.SetAtten (120 - CInt(txtValue.Text))
        If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
            PA5x2.SetAtten (120 - CInt(txtValue.Text))
        End If
        Form1.SetFocus
        'first we'll adjust a 500hz tone:
        sndPlaySound "C:\TinData\tintest_wav\s1.wav", SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
        dialcontrol1.Show_Arrows
        dialcontrol1.visible = True
        dialcontrol1.SetFocus
        TimerStep2.Enabled = True
        Do While chkClick.Value = 0  'wait until the user clicks the knob
            DoEvents
        Loop
        
        sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
        TimerStep2.Enabled = False
        txtIntensity.Text = 120 - CInt(txtValue.Text)
    
    'next we'll present a 5kHz tone.
        lblMainInstructions.Caption = "We will now present a second sound."
        lblInstruct2.Caption = "Please turn the dial until this sound is also at a comfortable level, then press to continue."
    VolAdj = True
    intMaxVolume = 0
    txtValue.Text = 30
    dialcontrol1.UserControl_Initialize
    dialcontrol1.setvolume (CInt(txtValue.Text))
    'set inital PA5 value to 90 here (pa5value = 120 - cint(txtvalue.text))
    PA5x1.SetAtten (120 - CInt(txtValue.Text))
    If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
        PA5x2.SetAtten (120 - CInt(txtValue.Text))
    End If
    Form1.SetFocus
    chkClick.Value = 0
    sndPlaySound "C:\TinData\tintest_wav\s6.wav", SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
    dialcontrol1.Show_Arrows
    dialcontrol1.visible = True
    dialcontrol1.SetFocus
    TimerStep2.Enabled = True
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    TimerStep2.Enabled = False
    txtIntensity2.Text = 120 - CInt(txtValue.Text)
    
    dialcontrol1.visible = False
    lblInstruct2.visible = False
    lblSoft.visible = False
    lblLoud.visible = False
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    VolAdj = False
    TimerStep2.Enabled = False
    'output info

    ' Open a file for appending and write the values of txtIntensity.Text and txtIntensity2.Text along with their corresponding descriptions.
        ' Parameters:
        '   - intfilenumber: The file number to be used for opening the file.
        '   - WorkingDir: The directory path where the file is located.
        '   - WorkingFile: The name of the file to be opened.
        '   - txtIntensity.Text: The value of the txtIntensity TextBox control.
        '   - txtIntensity2.Text: The value of the txtIntensity2 TextBox control.
        intfilenumber = FreeFile ' This is safer than assigning a number
        Open (WorkingDir & WorkingFile) For Append As #intfilenumber
            Write #intfilenumber, txtIntensity.Text, "Comfortable PA5 Value for 500 Hz"
            Write #intfilenumber, txtIntensity2.Text, "Comfortable PA5 Value for 5000 Hz"
        Close #intfilenumber
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'
' Subroutine: Step3_Bandwidth
' Description: This subroutine handles the selection of the bandwidth for the tinnitus sensation.
'              It displays instructions and options to the user, allows them to make a selection using a dial,
'              and saves the selected bandwidth to a file.
' Parameters: None
' Returns: None
Private Sub Step3_Bandwidth()
    ' Variable declarations
    Dim intfilenumber, t1 As Integer
    Dim TempString As String
    
    ' Set initial values and display instructions
    OneStep = False
        lblMainInstructions.Caption = "Which of the sounds below does your tinnitus sensation sound more like? "
        lblInstruct2.Caption = "Turn the dial to select and hear. Once you've decided, press the dial to move on."
    lblInstruct2.Top = 256 'need to move this down a bit as mainstructions span 2 lines.
    lblInstruct2.visible = True
    lblMainInstructions.visible = True
    
    ' Set initial values for controls
    txtValue.Text = 99
    chkChange.Value = 0
    PA5x1.SetAtten CInt(txtIntensity2.Text) 'set value of PA5 to intensity of 5khz tone
    
    ' Set initial values for dial and display it
    TempString = "Bandwidth is Hissing"
    txtValue.Text = 1
    soundbandwidthdial1.UserControl_Initialize
    soundbandwidthdial1.visible = True
    soundbandwidthdial1.SetFocus
    ' Wait for user to make a selection
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            OneStep = True
            txtValue.Text = CInt(soundbandwidthdial1.getvalue) + 1 'call the usercontrol and updated textbox
            
            ' Determine the selected bandwidth and update the relevant variables and controls
            Select Case CInt(txtValue.Text)
                Case Is = 1
                    TempString = "Bandwidth is Hissing"
                        UserBW = "Hissing"
                    sndPlaySound "C:\TinData\tintest_wav\w6.wav", SND_ASYNC Or SND_NODEFAULT
                Case Is = 2
                    TempString = "Bandwidth is Ringing"
                        UserBW = "Ringing"
                    sndPlaySound "C:\TinData\tintest_wav\r6.wav", SND_ASYNC Or SND_NODEFAULT
                Case Is >= 3
                    TempString = "Bandwidth is Tonal"
                        UserBW = "Tonal"
                    txtValue.Text = 3
                    sndPlaySound "C:\TinData\tintest_wav\s6.wav", SND_ASYNC Or SND_NODEFAULT
            End Select
        End If
        DoEvents
    Loop
    ' Hide dial and instructions
    soundbandwidthdial1.visible = False
    lblInstruct2.visible = False
    lblInstruct2.Top = 208 'put it back where it was
    
    ' Play silence sound and save selected bandwidth to file
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    txtBandwidth.Text = txtValue.Text
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, CInt(txtBandwidth.Text), TempString
    Close #intfilenumber
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'
' Subroutine: Step4_Temporal
' Description: This subroutine handles the fourth step of the tinnitus testing process, which involves
'              playing sounds and allowing the user to select the temporal property of their tinnitus.
'              The user can choose between a steady sound or a pulsing sound. The selected temporal
'              property is recorded in a file along with the corresponding value.
'
' Parameters:
'   None
'
' Returns:
'   None
'
' Example:
'   Call Step4_Temporal
'
Private Sub Step4_Temporal()
    Dim intfilenumber As Integer
    Dim PlayFileName1, PlayFileName2 As String
        lblMainInstructions.Caption = "Turn the dial to hear these sounds.  "
        lblInstruct2.Caption = "Which sound is most like your tinnitus? Once you've decided, press the dial to move on."
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    PA5x1.SetAtten CInt(txtIntensity2.Text) 'set value of PA5 in case resuming experiment
    If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
        PA5x2.SetAtten CInt(txtIntensity2.Text) 'set value of PA5 in case resuming experiment
    End If
    
    TempString = "Temporal Property is Steady"
    txtValue.Text = 1
    chkChange.Value = 0
    Select Case CInt(txtBandwidth.Text)
        Case Is = 1  'user selected "Hissing" in previous step
            PlayFileName1 = "C:\TinData\tintest_wav\w6.wav"
            PlayFileName2 = "C:\TinData\tintest_wav\w6_pulse_3s.wav"
        Case Is = 2 'user selected "Ringing" in previous step
            PlayFileName1 = "C:\TinData\tintest_wav\r6.wav"
            PlayFileName2 = "C:\TinData\tintest_wav\r6_pulse_3s.wav"
        Case Is = 3 'user selected "Tonal" in previous step
            PlayFileName1 = "C:\TinData\tintest_wav\s6.wav"
            PlayFileName2 = "C:\TinData\tintest_wav\s6_pulse_3s.wav"
    End Select

    soundtypedial1.UserControl_Initialize
    soundtypedial1.visible = True
    soundtypedial1.SetFocus
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            Select Case CInt(soundtypedial1.getvalue)
                Case Is = 0 'Steady Sound
                    TempString = "Temporal Property is Steady"
                    sndPlaySound PlayFileName1, SND_ASYNC Or SND_NODEFAULT
                    txtValue.Text = 1
                        UserSorP = "Steady"
                Case Is = 1 'off
                    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
                Case Is >= 2 'pulsing Sound
                        UserSorP = "Pulsing"
                    TempString = "Temporal Property is Pulsing"
                    txtValue.Text = 2
                    sndPlaySound PlayFileName2, SND_ASYNC Or SND_NODEFAULT
            End Select
        End If
        DoEvents
    Loop
    soundtypedial1.visible = False
    lblInstruct2.visible = False
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    txtTemporal.Text = txtValue.Text
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, CInt(txtTemporal.Text), TempString
    Close #intfilenumber

End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'
' Subroutine: Step5_LoudnessRating
' Description: This subroutine is responsible for rating the loudness of tinnitus.
'              It displays instructions and a scale for the user to rate the loudness.
'              The user can adjust the rating using a knob and press a button to record the rating.
'              The subroutine also saves the rating to a file.
'
' Parameters: None
' Returns: None
'
' Remarks: - The subroutine checks the language setting (English or non-English) to display the appropriate labels and instructions.
'          - The subroutine adjusts the position and visibility of various controls on the form.
'          - The user can adjust the rating by turning the knob and press the button to record the rating.
'          - The subroutine changes the color of certain controls temporarily when the button is pressed.
'          - The subroutine saves the rating to a file.
Private Sub Step5_LoudnessRating()
    Dim intfilenumber As Integer
    Dim TempTop2 As Integer
    Dim TempTop3 As Integer
        lblTitle.Caption = "Tinnitus Loudness Rating"
        lblMainInstructions.Caption = "We have now switched the sounds off, so that you can hear only your tinnitus. Listen to your tinnitus now."
        lblInstruct2.Caption = "How Loud is your tinnitus? "
        lblInstruct3.Caption = "Rate its loudness on the scale below by turning the dial, then press to record your rating."
        framLoudness.Caption = "Loudness Rating"
        lbl5 = "Extremely Weak"
        lbl30 = "Moderate"
        lbl50 = "Strong"
        lbl70 = "Very Strong"
        lbl95 = "Extremely Strong"
    lblInstruct2.Left = lblMainInstructions.Left
    lblInstruct3.Left = lblMainInstructions.Left
    TempTop2 = lblInstruct2.Top
    TempTop3 = lblInstruct3.Top
    lblInstruct2.Top = lblMainInstructions.Top + 120
    'lblInstruct3.Top = lblInstruct2.Top + (lblInstruct2.Top - lblMainInstructions.Top)
    lblInstruct3.Top = lblInstruct2.Top + 70
    lblTitle.visible = True
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    framLoudness.visible = True
    hscrScale.Value = 1
    Form1.SetFocus
    txtValue.Text = 1
    chkChange.Value = 0
    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
    ' This code snippet waits for the user to click a knob and adjusts the scale value accordingly.
    ' It continuously checks if the user has made any adjustments and updates the scale value based on the text value entered.
    ' If the text value is less than 1, it sets the scale value to 1. If the text value is greater than 100, it sets the scale value to 100.
    ' The loop continues until the user clicks the knob.
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        If chkChange.Value = 1 Then 'user has made an adjustment
            chkChange.Value = 0 'reset change flag
            If CInt(txtValue.Text) < 100 Then
                If CInt(txtValue.Text) > 1 Then
                    hscrScale.Value = CInt(txtValue.Text)
                Else
                    hscrScale.Value = 1
                    txtValue.Text = 1
                End If
            Else
                hscrScale.Value = 100
                txtValue.Text = 100
            End If
        End If
        DoEvents
    Loop
    '******test changing colour for button press
    framLoudness.BackColor = &HC0FFC0
    lbl5.BackColor = &HC0FFC0
    lbl30.BackColor = &HC0FFC0    'original color:  &H00F0F0E6&
    lbl50.BackColor = &HC0FFC0
    lbl70.BackColor = &HC0FFC0
    lbl95.BackColor = &HC0FFC0
    lblScale.BackColor = &HC0FFC0
    timerClick.Enabled = True
    Do While timerClick.Enabled
        DoEvents
    Loop
    lblTitle.visible = False
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    lblInstruct2.Top = TempTop2
    lblInstruct3.Top = TempTop3
    txtLoudness.Text = txtValue.Text
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, CInt(txtValue.Text), "Tinnitus Loudness Rating"
    Close #intfilenumber
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest_back.frm
'
' This code block is part of the TinTest_back.frm file and contains the subroutine "Step6_LoudnessMatching".
' The Step6_LoudnessMatching subroutine is responsible for conducting the loudness matching task in a tinnitus testing program.
' It displays instructions to the user, plays a series of sounds, and allows the user to adjust the volume using a dial control.
' The subroutine utilizes various variables and objects such as intfilenumber, c1, WavNum, PA5Results, SoundOrder, TSoundOrder, TempChar, MaxedOutFlag, temp, MaxedOutCounter, lblNextSound, lblTitle, lblMainInstructions, lblInstruct2, lblInstruct3, chkClick, vRes, lblNextSound, lblInstruct2, lblInstruct3, txtBandwidth, txtValue, dialcontrol1, Timer1, VolAdj, intMaxVolume, txtTimer, sndPlaySound, and more.
' It also includes conditional statements, loops, and function calls to handle different scenarios and interactions with the user.
Private Sub Step6_LoudnessMatching()
    Dim intfilenumber, c1, WavNum, PA5Results(11, 2) As Integer
    Dim SoundOrder(1 To 22) As String
    Dim TSoundOrder(1 To 11) As String
    Dim TempChar As String
    Dim MaxedOutFlag(1 To 11) As Byte
    Dim temp As Integer
    Dim MaxedOutCounter As Integer 'this is used to prevent the user from going into an infinate loop of 'yes, no' maxed out scenario
    
    If vRes = 1024 Then '
        lblNextSound.Top = 780
        lblNextSound.Alignment = 2 'center for a better look when the text is below the dial
    Else
        lblNextSound.Top = 440
        lblNextSound.Alignment = 0 'Left justify for better look when the text is above the dial
    End If

    'first we will show an intro screen that will describe the loudness mathcing task:
    lblTitle.Caption = "Loudness Matching"
    lblMainInstructions.Caption = "We are now going to measure the LOUDNESS of your tinnitus by presenting several sounds."
    lblInstruct2.Caption = "Using the dial, you will increase the LOUDNESS of each sound until it matches your tinnitus."
    lblInstruct3.Caption = "When ready, press the dial to hear the first sound."
    lblInstruct2.Top = 256 'these need to be bumped down a bit due to two lines in maininstruct
    lblInstruct3.Top = 376
    lblTitle.visible = True

    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    lblInstruct2.Top = 208  'reset positions for 2nd and 3rd line of instructions
    lblInstruct3.Top = 328

    'now we will start the actual loudness matching procedure:
    lblMainInstructions.Caption = "Using the dial, increase the LOUDNESS of this sound until it is the same LOUDNESS as your tinnitus (not softer or louder, but the same loudness as your tinnitus).  "
    lblInstruct3.Caption = "When it matches the LOUDNESS of your tinnitus, push the dial to move on."
    lblMainInstructions.visible = True
    lblInstruct3.visible = True
    Form1.SetFocus
    Timer1.Interval = 100
    VolAdj = True
    intMaxVolume = 0
    
    ' This Select Case statement is used to determine the value of the variable TempChar based on the user's selection of bandwidth.
    ' The value of txtBandwidth.Text is converted to an Integer using the CInt function.
    ' If the value is 1, TempChar is assigned the value "w" (representing "Hissing").
    ' If the value is 2, TempChar is assigned the value "r" (representing "Ringing").
    ' If the value is 3, TempChar is assigned the value "s" (representing "Tonal").
    ' If none of the above cases match, an error message is displayed and TempChar is assigned the default value "s" (representing "Tonal").
    'first we need to populate a randomized array with the sound order.  2 passes are used of all 11 stim.  Stim are
    'based on selection in step 3 (bandwidth)
    Select Case CInt(txtBandwidth.Text)
        Case Is = 1  'user selected "Hissing" in previous step
            TempChar = "w"
        Case Is = 2 'user selected "Ringing" in previous step
            TempChar = "r"
        Case Is = 3 'user selected "Tonal" in previous step
            TempChar = "s"
        Case Else
            MsgBox "Error: Case is " & temchar & ". Bandwidth data unavailable; selecting tonal as default for sound"
            TempChar = "s"
    End Select

    ' This code block initializes the arrays TSoundOrder, SoundOrder, and MaxedOutFlag.
    ' It assigns file paths to the elements of TSoundOrder array based on the value of TempChar and c1.
    ' The RandomizeArray function is then called to randomize the order of elements in TSoundOrder array.
    ' The elements of TSoundOrder array are then copied to SoundOrder array.
    ' Finally, the MaxedOutFlag array is set to 0 for all elements.
    For c1 = 1 To 11
        TSoundOrder(c1) = ("C:\TinData\tintest_wav\" & TempChar & CStr(c1) & ".wav")
    Next c1
    RandomizeArray TSoundOrder
    For c1 = 1 To 11
        SoundOrder(c1) = TSoundOrder(c1)
    Next c1
    For c1 = 1 To 11
        TSoundOrder(c1) = ("C:\TinData\tintest_wav\" & TempChar & CStr(c1) & ".wav")
    Next c1
    RandomizeArray TSoundOrder
    For c1 = 1 To 11
        SoundOrder(11 + c1) = TSoundOrder(c1)
    Next c1
    For c1 = 1 To 11  'make sure all of the flags are set to 0
        MaxedOutFlag(c1) = 0
    Next c1

    txtValue.Text = 0
    lblNextSound.Caption = "Starting first sound - begin turning dial"
    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
    '
    ' This code block is part of the TinTest.frm file and contains a subroutine called "Step6".
    ' The Step6 subroutine is responsible for playing a series of sounds and allowing the user to adjust the volume using a dial control.
    ' It also checks if the user can hear the sound and determines if it is softer than their tinnitus.
    ' The code block includes various conditional statements and control structures to handle different scenarios.
    ' It utilizes several variables and objects such as c1, dialcontrol1, txtValue, PA5x1, SoundOrder, MaxedOutFlag, chkClick, lblNextSound, txtTimer, sndPlaySound, and more.
    ' The code block is part of a larger program designed for tinnitus testing.
    ' Note: Some parts of the code are commented out and may not be active.
    '
    c1 = 1
    dialcontrol1.Show_Arrows
    dialcontrol1.visible = True
    dialcontrol1.SetFocus
    Timer1.Enabled = True 'turn on timer
    Do While (c1 <= 22) 'cycle through all 22 sounds
        txtValue.Text = 30
        dialcontrol1.UserControl_Initialize
        dialcontrol1.setvolume (CInt(txtValue.Text))
        dialcontrol1.Show_Arrows
        'set inital PA5 value to 90 here (pa5value = 120 - cint(txtvalue.text))
        PA5x1.SetAtten (120 - CInt(txtValue.Text))
        'the following if statements pull out the file number.  This will be used to store the values in an array
        'and calculate a final average
        If Len(SoundOrder(c1)) = 29 Then 'string is 29 character long, and thus a single digit extnsion
            WavNum = CInt(Mid(SoundOrder(c1), 25, 1))
        ElseIf Len(SoundOrder(c1)) = 30 Then 'string is 30 characters long, and thus double digit extension
            WavNum = CInt(Mid(SoundOrder(c1), 25, 2))
        Else
            MsgBox "Error in wav filename string length in Step6 Subroutine"
        End If
        If MaxedOutFlag(WavNum) = 0 Or MaxedOutFlag(WavNum) = 2 Then 'play the sound since the user can hear it
            chkClick.Value = 0
            MaxedOutFlag(WavNum) = 0 'reset this to 0 so they get a 2nd chance to hear and rate the sound for "Softer than tinnitus" (-102) scneario
            MaxedOutCounter = 0
            lblNextSound.visible = True
            'TimerStep6.Enabled = True  'this will update PA5s with subroutine call from dialcontrol1
            Do While chkClick.Value = 0  'wait until the user clicks the knob
                txtTimer.Text = 0 'reset the textbox that holds the timer variable
                sndPlaySound SoundOrder(c1), SND_ASYNC Or SND_NODEFAULT
                Do While (CInt(txtTimer.Text) < 40) 'loop for 4000ms
                    If chkChange.Value = 1 Then 'user has made an adjustment
                        txtValue.Text = dialcontrol1.getvolume
                        PA5x1.SetAtten (120 - CInt(txtValue.Text))
                        If CInt(txtValue.Text) = 999 Then 'user has continually tried to turn the volume up (5 times) past it's loudest level
                            Call hide_all
                            intMaxVolume = 0
                            dialcontrol1.visible = False
                            temp = CanYouHearThis(SoundOrder(c1))
                            If temp = 0 Then  'user can hear the sound and it is louder than their tinnitus. Go back to asking them to adjust volume
                                intMaxVolume = 0
                                txtTimer.Text = 0
                                dialcontrol1.movevolume (-1) 'this will reduce sound by 1 db but also reset the maxpast flag
                                dialcontrol1.movevolume (1) 'reset volume to full
                                MaxedOutCounter = MaxedOutCounter + 1 'this will count how many times the user enters this step
                                If MaxedOutCounter = 2 Then 'user has done this twice already for this sound, so mark it as full volume and go onto next sound
                                    txtTimer.Text = 40  'force timer to max to end loop
                                    PA5x1.SetAtten (120) 'set PA5 to 120 as a saftey precaution as it is now playing full volume
                                    chkClick.Value = 1 'exit from loop
                                    txtValue.Text = 120 'ensure a value of 0 is recorded
                                End If
                            ElseIf temp = 1 Then 'user cannot hear the sound
                                MaxedOutFlag(WavNum) = 1 'set flag appropriatly
                                txtTimer.Text = 40  'force timer to max to end loop
                                PA5x1.SetAtten (120) 'set PA5 to 120 as a saftey precaution as it is now playing full volume
                                chkClick.Value = 1 'exit from loop
                            ElseIf temp = 2 Then 'user CAN hear the sound, but it is quiter than their tinnitus
                                MaxedOutFlag(WavNum) = 2 'set flag appropriatly
                                txtTimer.Text = 40 'force timer to max to end loop
                                chkClick.Value = 1 'exit from loop
                            End If
                            
                            dialcontrol1.visible = True
                            lblMainInstructions.Caption = "Using the dial, increase the loudness of this sound until it is the same LOUDNESS as your tinnitus (not softer or louder, but the same loudness as your tinnitus).  "
                            lblInstruct3.Caption = "When it matches the loudness of your tinnitus, push the dial to move on."
                            lblMainInstructions.visible = True
                            lblInstruct3.visible = True
                            dialcontrol1.SetFocus
                        End If
                        'enter code here for PA5 adjustment
                        'PA5 value = 120-cint(txtValue.text)
                    End If
                    If (chkClick.Value = 1 And CInt(txtTimer.Text) > 21) Then txtTimer.Text = 40 'user clicked, so end loop early
                    DoEvents
                Loop
                lblNextSound.visible = False
                DoEvents
            Loop
            'TimerStep6.Enabled = False
        ElseIf MaxedOutFlag(WavNum) = 1 Then 'it has already been determined that the subject can't hear this sound
            'do not play sound again if they can't hear it
        End If

        ' This code block adjusts and stores the results of sound loudness testing.
        ' It uses the sndPlaySound function to play a silence.wav file and stop any currently playing sound.
        ' The VolAdj variable is set to False.
        ' Depending on the MaxedOutFlag value, the results are stored in the PA5Results array and displayed in the corresponding text boxes.
        ' If MaxedOutFlag is 0, the sound has been adjusted normally and the loudness value is stored.
        ' If MaxedOutFlag is 1, the user cannot hear the sound at all and the code -101 is stored.
        ' If MaxedOutFlag is 2, the user can hear the sound, but it is softer than their tinnitus and the code -102 is stored.
        sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop sound from playing
        VolAdj = False
        If MaxedOutFlag(WavNum) = 0 Then 'sound has been adjusted nomrally so output text normally
            If c1 <= 11 Then  'store the results in an array
                PA5Results(WavNum, 1) = 120 - CInt(txtValue.Text)
                txtLoudnessT1(WavNum - 1).Text = 120 - CInt(txtValue.Text)
            Else
                PA5Results(WavNum, 2) = 120 - CInt(txtValue.Text)
                txtLoudnessT2(WavNum - 1).Text = 120 - CInt(txtValue.Text)
            End If
        ElseIf MaxedOutFlag(WavNum) = 1 Then
            If c1 <= 11 Then  'store the results in an array
                PA5Results(WavNum, 1) = -101 'outputs code -101 which means user cannot hear sound at all
                txtLoudnessT1(WavNum - 1).Text = -101
            Else
                PA5Results(WavNum, 2) = -101 'outputs code -101 which means user cannot hear sound at all
                txtLoudnessT2(WavNum - 1).Text = -101
            End If
        ElseIf MaxedOutFlag(WavNum) = 2 Then
            If c1 <= 11 Then  'store the results in an array
                PA5Results(WavNum, 1) = -102 'outputs code -102 which means user can hear sound, but it is softer than their tinnitus
                txtLoudnessT1(WavNum - 1).Text = -102
            Else
                PA5Results(WavNum, 2) = -102 'outputs code -102 which means user can hear sound, but it is softer than their tinnitus
                txtLoudnessT2(WavNum - 1).Text = -102
            End If
        End If
        
        'insert a small 2 second pause between trials, unless we are skipping this sound due to maxed out volume
        If MaxedOutFlag(WavNum) = 1 And c1 > 11 Then   'don't insert pause
        Else 'insert pause
            txtTimer.Text = 0
            Do While (CInt(txtTimer.Text) < 20) 'loop for 4000ms
                DoEvents
            Loop
        End If
        c1 = c1 + 1
        VolAdj = True
        lblNextSound.Caption = "Starting next sound - begin turning dial"
    Loop
    Timer1.Enabled = False
   
    'output info
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    VolAdj = False
    dialcontrol1.visible = False
    lblInstruct3.visible = False
    lblSon.visible = False
    lblTitle.FontSize = 60
    lblTitle.FontBold = False
    lblTitle.Height = 97
    lblNextSound.Top = 440

    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, "PA5ValueTrial1", "PA5ValueTrial2", "PA5ValueAVG", "File"
        For c1 = 0 To 10
            Write #intfilenumber, CInt(txtLoudnessT1(c1).Text), CInt(txtLoudnessT2(c1).Text), (CInt(txtLoudnessT1(c1).Text) + CInt(txtLoudnessT2(c1).Text)) / 2, (TempChar & CStr(c1 + 1))
        Next c1
    Close #intfilenumber
End Sub

'' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
''
'' Subroutine: Step7_PitchMatching
'   Description: This subroutine is responsible for rating the pitch of tinnitus.
'           It displays instructions and a scale for the user to rate the pitch.
'           The user can adjust the rating using a knob and press a button to record the rating.
'           The subroutine also saves the rating to a file.
'   Parameters: None
'   Returns: None
'   Remarks: - The subroutine checks the language setting (English or non-English) to display the appropriate labels and instructions.
'            - The subroutine adjusts the position and visibility of various controls on the form.
'            - The user can adjust the rating by turning the knob and press the button to record the rating.
'            - The subroutine changes the color of certain controls temporarily when the button is pressed.
'            - The subroutine saves the rating to a file.
'   Example: Call Step7_PitchMatching
Private Sub Step7_PitchMatching()
    Dim intfilenumber, c1, WavNum, SliderResults(11, 3) As Integer
    Dim SoundOrder(1 To 33), TSoundOrder(1 To 11), TempChar As String
    Dim bClick As Boolean
    Dim t1 As Integer
    lblNextSound.Top = 440
    lblNextSound.Alignment = 0 'Left justify for better look when the text is above the dial
    'first we will show an intro screen that will describe the pitch mathcing task:
        lblTitle.Caption = "Pitch Matching"
        lblMainInstructions.Caption = "We are now going to present several sounds differing in PITCH."
        lblInstruct2.Caption = "Using the dial, rate the similarity of each PITCH to your tinnitus by adjusting the slider."
        lblInstruct3.Caption = "When ready, press the dial to hear the first sound."
        lblInstruct2.Top = 256 'these need to be bumped down a bit due to two lines in maininstruct
        lblInstruct3.Top = 376
    
    lblTitle.visible = True
    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    lblInstruct2.visible = False
    lblInstruct3.visible = False
        lblInstruct2.Top = 208
        lblInstruct3.Top = 328
    Else
    End If

    lblMainInstructions.Caption = "You should now hear a sound playing. How similar is the PITCH of the sound to your tinnitus? Rate PITCH similarity by turning the dial, then press to record your rating."

    lblMainInstructions.visible = True
    PA5x1.SetAtten (CInt(txtIntensity.Text))
    Form1.SetFocus
    Timer1.Interval = 100
        'first we need to populate a randomized array with the sound order.  2 passes are used of all 11 stim.  Stim are
        'based on selection in step 3 (bandwidth)
        Select Case CInt(txtBandwidth.Text)
            Case Is = 1  'user selected "Hissing" in previous step
                TempChar = "w"
            Case Is = 2 'user selected "Ringing" in previous step
                TempChar = "r"
            Case Is = 3 'user selected "Tonal" in previous step
                TempChar = "s"
        End Select
    '&&&&&&&&&&&&&&&& POPULATE ARRAYS &&&&&&&&&&&&&&&&&&&&&
        ' This section of code populates the SoundOrder array with file paths to sound files.
        ' It uses the TSoundOrder array to generate the file paths based on the value of TempChar.
        ' The file paths are in the format "C:\TinData\tintest_wav\[TempChar][c1].wav", where [TempChar] is the value of TempChar and [c1] is a number from 1 to 11.
        ' The TSoundOrder array is then randomized using the RandomizeArray function.
        ' The randomized TSoundOrder array is then copied to the SoundOrder array.
        ' This process is repeated three times, with the SoundOrder array being populated with a total of 33 file paths.
        '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '&&&&&&&&&&&&&&&& POPULATE ARRAYS &&&&&&&&&&&&&&&&&&&&&
    For c1 = 1 To 11
        TSoundOrder(c1) = ("C:\TinData\tintest_wav\" & TempChar & CStr(c1) & ".wav")
    Next c1
    RandomizeArray TSoundOrder
    For c1 = 1 To 11
        SoundOrder(c1) = TSoundOrder(c1)
    Next c1
    
    For c1 = 1 To 11
        TSoundOrder(c1) = ("C:\TinData\tintest_wav\" & TempChar & CStr(c1) & ".wav")
    Next c1
    RandomizeArray TSoundOrder
    For c1 = 1 To 11
        SoundOrder(11 + c1) = TSoundOrder(c1)
    Next c1
    
    For c1 = 1 To 11
        TSoundOrder(c1) = ("C:\TinData\tintest_wav\" & TempChar & CStr(c1) & ".wav")
    Next c1
    RandomizeArray TSoundOrder
    For c1 = 1 To 11
        SoundOrder(22 + c1) = TSoundOrder(c1)
    Next c1
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

        lbl5.Caption = "Not At All"
        lbl30.Caption = "Not very Similar"
        lbl50.Caption = "Somewhat Similar"
        lbl70.Caption = "Very Similar"
        lbl95.Caption = "Identical"
        lblNextSound.Caption = "Starting First Sound..."
        framLoudness.Caption = "Pitch Rating"
    txtValue.Text = 0

    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm

    ' This code block represents a subroutine that cycles through 33 sounds and performs various operations for each sound.
    ' It starts by setting the initial values and turning on the timer. Then, it checks the length of the sound filename string
    ' and determines the WavNum (file number) accordingly. If the user can hear the sound (volume not maxed out), it sets the volume
    ' level based on the user's input in previous steps. It waits for the user to click the knob and plays the sound using the
    ' sndPlaySound function. During the sound playback, it checks for any user adjustments and updates the scale accordingly.
    ' After the sound finishes playing, it stores the result in an array and inserts a small pause before moving to the next sound.
    ' If the user cannot hear the sound (volume maxed out), it sets the corresponding textboxes to -101. Finally, it performs some
    ' UI updates and stops the timer.
    c1 = 1
    framLoudness.visible = True
    Form1.SetFocus
    bClick = False
    txtTimer.Text = 0
    Timer1.Enabled = True 'turn on timer

    Do While (c1 <= 33) 'cycle through all 33 sounds:  11 sounds played 3 times each
        txtValue.Text = 0
        
        ' Determine the WavNum (file number) based on the length of the sound filename string
        If Len(SoundOrder(c1)) = 29 Then 'string is 29 character long, and thus a single digit extension
            WavNum = CInt(Mid(SoundOrder(c1), 25, 1))
        ElseIf Len(SoundOrder(c1)) = 30 Then 'string is 30 characters long, and thus double digit extension
            WavNum = CInt(Mid(SoundOrder(c1), 25, 2))
        Else
            MsgBox "Error in wav filename string length in Step7 Subroutine"
        End If
        
        ' Check if the user can hear the sound
        If (CInt(txtLoudnessT2(WavNum - 1) <> -101)) Then 'user did not max out volume in step 6 and can at least hear the sound
            
            ' Set the volume to the proper PA5 level
            If txtLoudnessT2(WavNum - 1).Text = -102 Then 'user maxed out sound but can still hear it
                PA5x1.SetAtten (0)
                If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
                    PA5x2.SetAtten (0)
                End If
            Else
                If CInt(txtLoudnessT1(WavNum - 1).Text) < CInt(txtLoudnessT2(WavNum - 1).Text) Then 'first trial was louder, use that value
                    PA5x1.SetAtten (CInt(txtLoudnessT1(WavNum - 1).Text))
                    If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
                        PA5x2.SetAtten (CInt(txtLoudnessT1(WavNum - 1).Text))
                    End If
                Else '2nd trial was louder...user that instead
                    PA5x1.SetAtten (CInt(txtLoudnessT2(WavNum - 1).Text))
                    If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
                        PA5x2.SetAtten (CInt(txtLoudnessT2(WavNum - 1).Text))
                    End If
                End If
            End If
            
            Do While chkClick.Value = 0  'wait until the user clicks the knob
                txtTimer.Text = 0 'reset the textbox that holds the timer variable
                sndPlaySound SoundOrder(c1), SND_ASYNC Or SND_NODEFAULT
                
                Do While (CInt(txtTimer.Text) < 40) 'loop for 4000ms
                    If chkChange.Value = 1 Then 'user has made an adjustment
                        chkChange.Value = 0 'reset change flag
                        
                        ' Update the scale based on the user's adjustment
                        If CInt(txtValue.Text) < 100 Then
                            If CInt(txtValue.Text) > 1 Then
                                hscrScale.Value = CInt(txtValue.Text)
                            Else
                                hscrScale.Value = 1
                                txtValue.Text = 1
                            End If
                        Else
                            hscrScale.Value = 100
                            txtValue.Text = 100
                        End If
                    End If
                    
                    If (chkClick.Value = 1 And bClick = False) Then 'the click boolean is necessary to make sure this code runs only once
                        'since sound plays for 2 seconds, this will always allow the sound to finish playing out without
                        'any clicks or pops.  The user will recieve immediate feedback though indicating the button has
                        'been pressed:
                        If CInt(txtTimer.Text) <= 20 Then
                            txtTimer.Text = 40 - (21 - CInt(txtTimer.Text))
                        Else
                            txtTimer.Text = 40 'Sound is already stopped so we can safely exit loop without any clicks
                        End If
                        
                        framLoudness.BackColor = &HC0FFC0
                        lbl5.BackColor = &HC0FFC0
                        lbl30.BackColor = &HC0FFC0    'original color:  &H00F0F0E6&
                        lbl50.BackColor = &HC0FFC0
                        lbl70.BackColor = &HC0FFC0
                        lbl95.BackColor = &HC0FFC0
                        lblScale.BackColor = &HC0FFC0
                        
                        timerClick.Enabled = True
                        
                        Do While timerClick.Enabled
                            DoEvents
                        Loop
                        
                        bClick = True
                    End If
                    
                    DoEvents
                Loop
                
                lblNextSound.visible = False
                DoEvents
            Loop
            
            bClick = False
            sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop sound from playing
            
            lblNextSound.Caption = "Starting Next Sound"
            
            If c1 <= 11 Then  'store the results in an array
                SliderResults(WavNum, 1) = CInt(txtValue.Text)
                txtPitchMatchT1(WavNum - 1).Text = CInt(txtValue.Text)
            ElseIf c1 <= 22 Then
                SliderResults(WavNum, 2) = CInt(txtValue.Text)
                txtPitchMatchT2(WavNum - 1).Text = CInt(txtValue.Text)
            Else 'c1 >22 and <33
                SliderResults(WavNum, 3) = CInt(txtValue.Text)
                txtPitchMatchT3(WavNum - 1).Text = CInt(txtValue.Text)
            End If
            
            'insert a small 2 second pause between trials
            txtTimer.Text = 0
            Do While (CInt(txtTimer.Text) < 20) 'loop for 2000ms
                DoEvents
            Loop
            
        Else
            txtPitchMatchT1(WavNum - 1).Text = -101
            txtPitchMatchT2(WavNum - 1).Text = -101
            txtPitchMatchT3(WavNum - 1).Text = -101
        End If
        
        c1 = c1 + 1
    Loop

    Timer1.Enabled = False
    lblTitle.FontSize = 60
    lblTitle.FontBold = False
    lblTitle.Height = 97
    lblSon.visible = False
    lblNextSound.Top = 440

    'output info
        sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
        intfilenumber = FreeFile ' This is safer than assigning a number
        Open (WorkingDir & WorkingFile) For Append As #intfilenumber
            Write #intfilenumber, "SliderValueTrial1", "SliderValueTrial2", "SliderValueTrial3", "SliderValueAVG", "File"
            For c1 = 0 To 10
                Write #intfilenumber, CInt(txtPitchMatchT1(c1).Text), CInt(txtPitchMatchT2(c1).Text), CInt(txtPitchMatchT3(c1).Text), (CInt(txtPitchMatchT1(c1).Text) + CInt(txtPitchMatchT2(c1).Text) + CInt(txtPitchMatchT3(c1).Text)) / 3, (TempChar & CStr(c1 + 1))
            Next c1
        Close #intfilenumber

End Sub

Private Sub Step8_Threshold()
    Dim c1 As Integer
    Dim SoundOrderFile(1 To 4) As String
    Dim SoundOrder(1 To 4) As Integer
    Dim PA5Level(1 To 4) As Integer  'since we need to present different sounds at different start volumes
    Dim vLevels(0 To 3) As Integer
    Dim PA5ValuePure, PA5ValueHiss As Integer
    Dim PassFlag As Integer
    Dim MCustString As String
    Dim TempString As String
    Dim ShellResults As Integer
    Dim CM As String 'holds the filename for Custom Masker
    
    If vRes = 1024 Then '
        lblNextSound.Top = 400
        lblNextSound.Alignment = 0 'center for a better look when the text is below the dial
    Else
        lblNextSound.Top = 440
        lblNextSound.Alignment = 0 'Left justify for better look when the text is above the dial
    End If
    
    VolAdj = True
    intMaxVolume = 0
    TempString = ""
    For c1 = 0 To 10
        TempString = TempString & " " & CInt((CInt(txtPitchMatchT1(c1).Text) + CInt(txtPitchMatchT2(c1).Text) + CInt(txtPitchMatchT3(c1).Text)) / 3)
    Next c1
'    'we must generate a custom masker:
    txtTimer.Text = 0
    Timer1.Interval = 100

    CM = CustomMaskerPath() 'returns the file to use for custom masker
    'first we determine the order of the maskers
    SoundOrder(1) = 2
    SoundOrder(2) = 2
    SoundOrder(3) = 2
    SoundOrder(4) = 2

    SoundOrderFile(1) = "C:\TinData\tintest_wav\w6.wav"
    PA5Level(1) = CInt(txtIntensity2.Text) 'set inital sound intensity to the 5000Hz comfortable level

    lblTitle.Caption = "Threshold Measurement"
    lblMainInstructions.Caption = "Pitch matching is now complete."
    lblInstruct2.Caption = "In the next step, we are going to measure how your tinnitus is affected by masking sounds."
    lblInstruct3.Caption = "When ready, press the dial to begin."
    lblSoft.Caption = "Softer"
    lblLoud.Caption = "Louder"

    lblMainInstructions.visible = True
    lblInstruct2.visible = True
    lblInstruct3.visible = True
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop

    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
    ' This code block sets the visibility of various labels and captions based on user instructions.
    ' It also determines the user's threshold by adjusting the dial and playing sounds.
    ' The code then sorts the threshold values and calculates the median value.
    ' Finally, it sets the visibility and captions for the next step of the process.
    ' Note: The code references external objects such as lblInstruct2, lblInstruct3, lblTitle, lblSoft, lblLoud, lblMainInstructions, txtValue, PA5x1, PA5x2, dialcontrol1, chkClick, chkChange, sndPlaySound, TimerStep8, txtSoundThreshold, vLevels, BubbleSortArray, txtPA5ThreshValue, lblNextSound, PA5ValuePure.
    lblInstruct2.visible = False
    lblInstruct3.visible = False
    lblTitle.visible = True
    lblSoft.visible = True
    lblLoud.visible = True
    '******************************************
    '      determine users threshold          *
    '******************************************
    For c1 = 0 To 3 Step 1
        ' This code block sets the instructions and values for a dial control in a form.
        ' If c1 is equal to 0, it sets the instructions for the dial control based on the language preference.
        ' It also sets the initial value of txtValue to 0 and adjusts the attenuation level of PA5x1 and PA5x2 accordingly.
        ' Finally, it sets the volume of dialcontrol1 to 0.
        If c1 = 0 Then
            lblMainInstructions.Caption = "Turn the dial clockwise until you can hear a sound and stop turning it as soon as you hear it.  Then press to move on to the next step."
            txtValue.Text = 0 ' set PA5 to 120
            PA5x1.SetAtten (120 - CInt(txtValue.Text))
            dialcontrol1.setvolume (0)
        ' This code block is part of a conditional statement that executes when the value of c1 is equal to 1.
        ' It sets the caption of the lblMainInstructions label based on the value of the English variable.
        ' It also sets the text of the txtValue textbox based on the calculation using the txtSoundThreshold value.
        ' Additionally, it sets the attenuation level of the PA5x1 and PA5x2 objects based on the txtValue.
        ' Finally, it sets the volume of the dialcontrol1 object based on the txtValue.
        ElseIf c1 = 1 Then
            lblMainInstructions.Caption = "You should now hear the sound playing softly.  Turn the dial counter-clockwise until the sound just disappears, then click to move on."
            txtValue.Text = 120 - (CInt(txtSoundThreshold(0)) - 10) 'set PA5 to last threshold, plus 10dB
            PA5x1.SetAtten (120 - CInt(txtValue.Text))
            dialcontrol1.setvolume (CInt(txtValue.Text))
        ElseIf c1 = 2 Then
            lblMainInstructions.Caption = "Once again, please turn the dial clockwise until you can just barely hear a sound.  Then press to move on to the next step."
            
            txtValue.Text = 0
            PA5x1.SetAtten (120 - CInt(txtValue.Text))
            dialcontrol1.setvolume (0)
        Else 'c1 = 3
                lblMainInstructions.Caption = "Now turn the dial counter-clockwise until the sound just disappears.  Press to move on."

            txtValue.Text = 120 - (CInt(txtSoundThreshold(2)) - 10) 'set PA5 to last threshold, plus 10dB
            PA5x1.SetAtten (120 - CInt(txtValue.Text))
            dialcontrol1.setvolume (CInt(txtValue.Text))
        End If
        lblMainInstructions.visible = True
        'Form1.SetFocus
        dialcontrol1.visible = True
        dialcontrol1.SetFocus
        chkClick.Value = 0
 
        'play 1khz tone to find users threshold.
        ' Plays a sound file and waits for user interaction.
        ' The sound file path is "C:\TinData\tintest_wav\s2.wav".
        ' The sound is played asynchronously and continuously in a loop.
        ' The TimerStep8.Enabled property is set to True.
        ' The code waits in a loop until the user clicks the knob.
        ' If the user makes an adjustment, the change flag is reset and the attenuation level is updated.
        ' If the user is using two PA5 devices, the attenuation level is also updated for the second device.
        ' The loop continues until the user clicks the knob.
        ' This code utilizes the sndPlaySound function and various control properties.
        ' Note: This code assumes the existence of certain objects and controls (e.g., sndPlaySound, TimerStep8, chkClick, chkChange, PA5x1, PA5x2, txtValue).
        ' It is recommended to review the entire code file for a complete understanding of the functionality.
        sndPlaySound "C:\TinData\tintest_wav\s2.wav", SND_ASYNC Or SND_NODEFAULT Or SND_LOOP
        TimerStep8.Enabled = True
        Do While chkClick.Value = 0  'wait until the user clicks the knob
            If chkChange.Value = 1 Then 'user has made an adjustment
                chkChange.Value = 0 'reset change flag
                PA5x1.SetAtten (120 - CInt(txtValue.Text))
                If usePA52 Then 'user is using 2 pa5s so set level for 2nd pa5
                    PA5x2.SetAtten (120 - CInt(txtValue.Text))
                End If
                'PA5 value = 120-cint(txtValue.text)
            End If
            DoEvents
        Loop
        sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
        TimerStep8.Enabled = False
        txtSoundThreshold(c1).Text = 120 - CInt(txtValue.Text)
        vLevels(c1) = CInt(txtSoundThreshold(c1).Text)
    Next c1
    BubbleSortArray vLevels   'sort array
    txtPA5ThreshValue.Text = (vLevels(1) + vLevels(2)) / 2 'calc the median value for a 4 number array
    'MsgBox ("vLevel 0: " & vLevels(0) & "vLevel 1: " & vLevels(1) & "vLevel 2: " & vLevels(2) & "vLevel 3: " & vLevels(3))
    '**********************************************
    lblSoft.visible = False
    lblLoud.visible = False
        lblTitle.Caption = "Loudness Matching"
        lblMainInstructions.Caption = "Turn the dial slowly until the second sound is the same loudness as the first.  When you are satisfied they are the SAME LOUDNESS, press the dial to move on."
        lblNextSound.Caption = "Presenting two sounds"
    lblNextSound.visible = True
    If CInt(txtPA5ThreshValue.Text) > 65 Then
        PA5ValuePure = CInt(txtPA5ThreshValue.Text) - 65
    Else
        PA5ValuePure = 0
    End If
    
    ' This code block is responsible for presenting sounds and adjusting sound levels.
    ' It iterates through a loop to present multiple sounds and allows the user to adjust the sound level using a dial control.
    ' The code also includes pauses between sounds and handles user interactions such as clicking the knob.
    ' Sound levels are calculated based on the difference between a predefined value and the PA5 level.
    ' The code also updates a text box with the sound level and stores the sound level matches in an array.
    ' The code is specific to the TinTest application and the UofM version.
    ' 
    ' Note: The code references external resources such as sound files and controls (e.g., dialcontrol1, chkClick).
    '       The code also includes conditional statements based on certain flags (e.g., usePA52, English).
    '       The code assumes the existence of other variables and functions (e.g., PA5Level, PA5ValuePure, SoundOrderFile).
    '       Please refer to the complete codebase for a comprehensive understanding.
    
    c1 = 1
    PA5ValueHiss = PA5Level(1)
    txtValue.Text = (120 - PA5ValueHiss) 'txtvalue holds the sound level for the hissing
    dialcontrol1.setvolume (CInt(txtValue.Text))
    chkClick.Value = 0
    Do While c1 <= 1
        lblNextSound.visible = True
        txtTimer.Text = 0
        Timer1.Enabled = True
        Do While CInt(txtTimer.Text) < 40 'insert 4 second pause
            DoEvents
        Loop
        PA5ValueHiss = PA5Level(c1)
        txtValue.Text = (120 - PA5ValueHiss)
        dialcontrol1.Show_Arrows
        dialcontrol1.setvolume (CInt(txtValue.Text))
        Do While chkClick.Value = 0  'wait until the user clicks the knob
            txtTimer.Text = 0
            Timer1.Enabled = True
            'reset the textbox that holds the timer variable
            Do While (CInt(txtTimer.Text) <= 60)
                'set PA5 to apprpriate level
                'pa5.level = pa5valuehiss
                PA5x1.SetAtten (PA5ValuePure)
                sndPlaySound "C:\TinData\tintest_wav\s2.wav", SND_ASYNC Or SND_NODEFAULT
                
                Do While (CInt(txtTimer.Text) < 30) 'loop for 3000ms - 2000ms for sound 1000ms off
                    If chkChange.Value = 1 Then 'user has made an adjustment
                        chkChange.Value = 0 'reset change flag
                        'Change only the txtvalue box, not the PA5 here
                        'PA5 value = 120-cint(txtValue.text)
                    End If
                    DoEvents
                Loop
                If chkClick.Value = 1 Then txtTimer.Text = 61 'user clicked, so end loop and go to next sound
                lblNextSound.visible = False
                DoEvents
                txtValue.Text = dialcontrol1.getvolume
                If CInt(txtValue.Text) <> 999 Then
                    PA5x1.SetAtten (120 - CInt(txtValue.Text))
                Else
                    PA5x1.SetAtten (0)
                End If
                If CInt(txtTimer.Text) <= 60 Then
                    sndPlaySound SoundOrderFile(c1), SND_ASYNC Or SND_NODEFAULT
                    TimerStep8.Enabled = True
                    Do While (CInt(txtTimer.Text) <= 60) 'loop for 3000ms: 2000ms for sound 1000ms off
                        If chkChange.Value = 1 Then 'user has made an adjustment
                            chkChange.Value = 0 'reset change flag
                            If CInt(txtValue.Text) <> 999 Then
                                PA5x1.SetAtten (120 - CInt(txtValue.Text))
                            Else
                                PA5x1.SetAtten (0)
                            End If
                            'PA5 value = 120-cint(txtValue.text)
                        End If
                        'If chkClick.value = 1 Then txtTimer.Text = 60 'user clicked, so end loop early
                        '***had to remove the above line because it was causing clicks when the sound shut off.
                        '***now use gets instant feed back that dial has been pushed, but sound is allowed to finish playing
                        DoEvents
                    Loop
                    TimerStep8.Enabled = False
                End If
            Loop
            DoEvents
        Loop
        txtTimer.Text = 0
        Do While CInt(txtTimer.Text) < 20 'insert 2 second pause
            DoEvents
        Loop
        chkClick.Value = 0
        Timer1.Enabled = False
            lblNextSound.Caption = "Presenting two more sounds"
        If CInt(txtValue.Text) <> 999 Then
            'txtSoundLevelMatch(SoundOrder(c1) - 1) = 120 - CInt(txtValue.Text)
            'UofM only:
            txtSoundLevelMatch(0) = 120 - CInt(txtValue.Text)
            txtSoundLevelMatch(1) = 120 - CInt(txtValue.Text)
            txtSoundLevelMatch(2) = 120 - CInt(txtValue.Text)
            txtSoundLevelMatch(3) = 120 - CInt(txtValue.Text)
            
        Else
            'txtSoundLevelMatch(SoundOrder(c1) - 1) = 0
            'UofM only:
            txtSoundLevelMatch(0) = 0
            txtSoundLevelMatch(1) = 0
            txtSoundLevelMatch(2) = 0
            txtSoundLevelMatch(3) = 0
        End If
        c1 = c1 + 1
        DoEvents
    Loop
    
    ' This code block stops the sound from playing, disables Timer1, hides dialcontrol1 and lblNextSound,
    ' sets the font size of lblTitle to 60, and appends information to a file.
    ' The information includes the sound thresholds, PA5 value for 1 kHz threshold,
    ' and PA5 values for file w6 and custom masker.
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop sound from playing
    Timer1.Enabled = False
    VolAdj = False
    dialcontrol1.visible = False
    lblNextSound.visible = False
    lblTitle.FontSize = 60
    'output info
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, (CInt(txtSoundThreshold(0).Text)), "PA5 Threshold # 1"
        Write #intfilenumber, (CInt(txtSoundThreshold(1).Text)), "PA5 Threshold # 2"
        Write #intfilenumber, (CInt(txtSoundThreshold(2).Text)), "PA5 Threshold # 3"
        Write #intfilenumber, (CInt(txtSoundThreshold(3).Text)), "PA5 Threshold # 4"
        Write #intfilenumber, (CInt(txtPA5ThreshValue.Text)), "PA5Value used for 1 kHz Threshold"
        Write #intfilenumber, (CInt(txtSoundLevelMatch(1).Text)), "PA5Value for file w6"
    Close #intfilenumber

End Sub

Private Sub Step9_ResidualInhibition()
    Dim intfilenumber, c1 As Integer
    Dim SoundOrder(1 To 8) As String
    Dim OutputBox(1 To 8) As Integer 'this keeps track of what txt box the data is written to
    Dim SO(1 To 4) As Integer
    Dim CM As String
    CM = CustomMaskerPath()
    SO(1) = 1
    SO(2) = 2
    SO(3) = 3
    SO(4) = 4
    
    lblMuchLouder(0).Caption = "TINNITUS MUCH LOUDER"
    lblMuchLouder(1).Caption = "TINNITUS MUCH LOUDER"
    lblLouder(0).Caption = "LOUDER"
    lblLouder(1).Caption = "LOUDER"
    lblNoChange(0).Caption = "NO CHANGE"
    lblNoChange(1).Caption = "NO CHANGE"
    lblSofter(0).Caption = "SOFTER"
    lblSofter(1).Caption = "SOFTER"
    lblGone(0).Caption = "TINNITUS GONE"
    lblGone(1).Caption = "TINNITUS GONE"

    'fixed order.  Only 1 sound:
    SoundOrder(1) = "C:\TinData\tintest_wav\w6_30s.wav"
    OutputBox(1) = 0
    SoundOrder(2) = "C:\TinData\tintest_wav\w6_30s.wav"
    OutputBox(2) = 1
    
    ' Set the interval of Timer1 to 1000 milliseconds (1 second)
    Timer1.Interval = 1000

    ' Set the caption of lblMainInstructions to display a message to the user
    lblMainInstructions.Caption = "Please sit back and relax. An instruction will appear on the screen in 90 seconds and you will then proceed to the last part of the test."

    ' Make lblMainInstructions visible to the user
    lblMainInstructions.Visible = True

    ' Set the text of txtTimer to 0
    txtTimer.Text = 0

    ' Enable Timer1
    Timer1.Enabled = True

    ' Set the initial value and maximum value of ProgressBar1
    ProgressBar1.Value = 0
    ProgressBar1.Max = 90

    ' Make ProgressBar1 visible to the user
    ProgressBar1.Visible = True

    ' Loop for 90 seconds
    Do While (CInt(txtTimer.Text) < 90)
        ' Update the value of ProgressBar1 to match the current value of txtTimer
        ProgressBar1.Value = CInt(txtTimer.Text)
        
        ' Allow other events to be processed
        DoEvents
    Loop

    ' Hide ProgressBar1
    ProgressBar1.Visible = False

    ' Update the caption of lblMainInstructions to prompt the user to press the dial
    lblMainInstructions.Caption = "Please press the dial to proceed to the last part of the test."
    lblMainInstructions.visible = True
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        DoEvents
    Loop
    ProgressBar1.visible = False

    ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm

    ' This section of code is responsible for playing a sound and allowing the user to rate how their tinnitus has changed.
    ' It starts by disabling the Timer1 control and setting the appropriate instructions based on the language.
    ' Then, it sets the initial values for the timer, progress bar, and enables the Timer1 control.
    ' The first loop runs for 30 seconds, allowing the user to read the instructions.
    ' After that, the Timer1 control is disabled and the progress bar is hidden.
    ' The sound is played using the sndPlaySound function, and another loop runs for 30 seconds to allow the sound to play uninterrupted.
    ' The sound is then stopped using the silence.wav file.
    ' The focus is set back to Form1, and the timer is disabled.
    ' Depending on the tinnitus localization, different instructions and frames are displayed.
    ' The user is prompted to adjust the dial to rate how their tinnitus has changed.
    ' If the tinnitus is monaural, the user adjusts the dial for either the left or right ear.
    ' If the tinnitus is bilateral, the user adjusts the dial for both ears.
    ' The user's rating is stored in the appropriate text boxes.
    ' The loop continues until all sounds have been played and rated.
    ' The documentation comment provides an overview of the functionality and purpose of this code section.
    c1 = 1
    Do While c1 <= 2
        Timer1.Enabled = False

        lblMainInstructions.Caption = "Please listen to your tinnitus now.  We will soon present a sound.  When the sound ends, we will ask you to rate how your tinnitus has changed."

        txtTimer.Text = 0
        ProgressBar1.Value = 0
        ProgressBar1.Max = 30
        ProgressBar1.visible = True
        Timer1.Enabled = True
        Do While (CInt(txtTimer.Text) < 30) 'loop for 30s, this will give the user a chance to read the above instructions.
            DoEvents
            ProgressBar1.Value = CInt(txtTimer.Text)
        Loop
        Timer1.Enabled = False
        ProgressBar1.visible = False
        lblMainInstructions.Caption = "Listen carefully to this sound"

        PA5x1.SetAtten (CInt(txtSoundLevelMatch(1).Text))  'used for UofM only.

        txtTimer.Text = 0
        Timer1.Enabled = True
        sndPlaySound SoundOrder(c1), SND_ASYNC Or SND_NODEFAULT
        Do While (CInt(txtTimer.Text) < 30) 'loop for 30000ms, this will allow the sound to play uninterupted
            DoEvents
        Loop
        sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT 'stop sound from playing
        
        Form1.SetFocus
        txtTimer.Enabled = False
        
        ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest_back.frm

        ' This code block handles the adjustment of tinnitus rating in a monaural scenario.
        ' It checks the value of txtLocalize.Text to determine if tinnitus is in the left ear (1) or right ear (3).
        ' The lblMainInstructions.Caption is updated accordingly.
        ' The caption and text color of frmMono(0) frame is changed based on the ear localization.
        ' The position of frmMono(0) frame is adjusted based on the value of vRes.
        ' The VScroll1(0) value and txtValue.Text are set to 51.
        ' The visibility of frmMono(0) and frmMono(1) frames are set to False.
        ' The user is prompted to click the knob by checking the value of chkClick.
        ' If the user makes an adjustment, the value of VScroll1(0) and txtValue.Text are updated accordingly.
        ' The loop continues until the user clicks the knob.
        ' After the loop, the visibility of frmMono(0) and frmMono(1) frames are set to False.
        ' Depending on the value of c1, the tinnitus rating values are filled in the appropriate text boxes.
        If (CInt(txtLocalize.Text) = 1) Or (CInt(txtLocalize.Text) = 3) Then   ' tinnitus is monaural

            lblMainInstructions.Caption = "Adjust the dial to rate how your tinnitus has changed, and press to move on."

            If (CInt(txtLocalize.Text) = 1) Then 'tinnitus is in left ear, change caption on frame
                frmMono(0).Caption = "Left Ear"
                frmMono(0).ForeColor = &HC00000 'chage text to blue
            Else 'tinnitus is in right ear, change caption on frame
                frmMono(0).Caption = "Right Ear"
                frmMono(0).ForeColor = &HFF& 'change text to red
            End If
            If vRes = 1024 Then 'everything needs to be moved up to fit in the slider
                frmMono(0).Left = (Form1.ScaleWidth / 2) - (frmMono(0).Width / 2)
                frmMono(0).Top = (Form1.ScaleHeight / 2) - (frmMono(0).Height / 2) + 120 'put it lower than center so it doesn't cover up text
            Else
                frmMono(0).Left = (Form1.ScaleWidth / 2) - (frmMono(0).Width / 2)
                frmMono(0).Top = (Form1.ScaleHeight / 2) - (frmMono(0).Height / 2) + 200 'put it lower than center so it doesn't cover up text
            End If
            
            VScroll1(0).Value = 51
            txtValue.Text = 51
            frmMono(0).visible = True
            chkClick.Value = 0
            Do While chkClick.Value = 0  'wait until the user clicks the knob
                If chkChange.Value = 1 Then 'user has made an adjustment
                    chkChange.Value = 0 'reset change flag
                    If (101 - CInt(txtValue.Text)) < 100 Then  'if scroll bar greater than 1 (-5)
                        If (101 - CInt(txtValue.Text)) > 0 Then 'if scroll bar is less than 101 (+5)
                            VScroll1(0).Value = 102 - CInt(txtValue.Text)
                        Else
                            VScroll1(0).Value = 1
                            txtValue.Text = 101
                        End If
                    Else 'set scroll bar to -5
                        VScroll1(0).Value = 101
                        txtValue.Text = 1
                    End If
                End If
                DoEvents
            Loop
            frmMono(0).visible = False
            frmMono(1).visible = False
            If c1 < 5 Then
                If (CInt(txtLocalize.Text) = 1) Then 'tinnitus is in left ear,Fill in appropriate box
                    txtRILeftT1(OutputBox(c1)).Text = txtValue.Text
                    txtRIRightT1(OutputBox(c1)).Text = -939
                Else 'tinnitus is in right ear, Fill in appropriate box
                    txtRILeftT1(OutputBox(c1)).Text = -939
                    txtRIRightT1(OutputBox(c1)).Text = txtValue.Text
                End If
            Else
                If (CInt(txtLocalize.Text) = 1) Then 'tinnitus is in left ear,Fill in appropriate box
                    txtRILeftT2(OutputBox(c1)).Text = txtValue.Text
                    txtRIRightT2(OutputBox(c1)).Text = -939
                Else 'tinnitus is in right ear, Fill in appropriate box
                    txtRILeftT2(OutputBox(c1)).Text = -939 'this will output -99 into the box.  Necessary due to formlua
                    txtRIRightT2(OutputBox(c1)).Text = txtValue.Text
                End If
            End If

        ' This code block adjusts the user interface based on the value entered in the txtLocalize textbox.
        ' If the value is 2, indicating bilateral tinnitus, the code sets up the UI for the user to enter information for both ears.
        ' It updates the captions, colors, positions, and visibility of various form controls.
        ' The code also handles user interactions with the scroll bars and updates the corresponding textboxes with the selected values.
        ' The loop continues until the variable c1 reaches a certain value.
        ' Finally, the Timer1 control is disabled.
        ElseIf (CInt(txtLocalize.Text) = 2) Then  'tinnitus is bilateral...user must enter info on two frames

            lblMainInstructions.Caption = "Adjust the dial to rate how your tinnitus has changed, first in your left ear, then right ear, and press to move on."
            frmMono(0).Caption = "Left Ear"
            frmMono(1).Caption = "Right Ear"
            frmMono(0).ForeColor = &HC00000 'chage text to blue
            frmMono(1).ForeColor = &HFF& 'change text to red
            If vRes = 1024 Then 'everything needs to be moved up to fit in the slider
                frmMono(0).Left = (Form1.ScaleWidth / 2) - (frmMono(0).Width / 2) - 250
                frmMono(0).Top = (Form1.ScaleHeight / 2) - (frmMono(0).Height / 2) + 120 'put it lower than center so it doesn't cover up text
                frmMono(1).Left = (Form1.ScaleWidth / 2) - (frmMono(1).Width / 2) + 250
                frmMono(1).Top = (Form1.ScaleHeight / 2) - (frmMono(1).Height / 2) + 120 'put it lower than center so it doesn't cover up text
                lblMainInstructions.Top = 80
                lblEar.Left = 160
                lblEar.Top = 200
            Else
                frmMono(0).Left = 112
                frmMono(0).Top = 352
                frmMono(1).Left = 664
                frmMono(1).Top = 352
                lblMainInstructions.Top = 144
                lblEar.Left = 160
                lblEar.Top = 300
                
            End If
            
            VScroll1(0).Value = 51
            frmMono(0).visible = True
            VScroll1(1).Value = 51
            frmMono(1).visible = True
            'frmMono(0).Enabled = True
            'frmMono(1).Enabled = False
            txtValue.Text = 51

            lblEar.ForeColor = &HC00000 'chage text to blue
                lblEar.Caption = "Adjust Left Ear First..."
            lblEar.visible = True
            chkClick.Value = 0 'do left ear first
            Form1.SetFocus
            Do While chkClick.Value = 0  'wait until the user clicks the knob
                If chkChange.Value = 1 Then 'user has made an adjustment
                    chkChange.Value = 0 'reset change flag
                    If (101 - CInt(txtValue.Text)) < 100 Then 'if scroll bar greater than 1 (-5)
                        If (101 - CInt(txtValue.Text)) > 0 Then 'if scroll bar is less than 101 (+5)
                            VScroll1(0).Value = (102 - CInt(txtValue.Text))
                        Else
                            VScroll1(0).Value = 1
                            txtValue.Text = 101
                        End If
                    Else
                        VScroll1(0).Value = 101
                        txtValue.Text = 1
                    End If
                End If
                DoEvents
            Loop
            If c1 < 5 Then
                txtRILeftT1(OutputBox(c1)).Text = txtValue.Text
            Else
                txtRILeftT2(OutputBox(c1)).Text = txtValue.Text
            End If
            'frmMono(0).Enabled = False
            'frmMono(1).Enabled = True
            txtValue.Text = 51
            If vRes = 1024 Then 'everything needs to be moved up to fit in the slider
                lblEar.Left = (Form1.ScaleWidth / 2) - (lblEar.Width / 2) + 150
                lblEar.Top = 200
            Else
                lblEar.Left = 712
                lblEar.Top = 300
            End If
            lblEar.ForeColor = &HFF& 'change text to red

            lblEar.Caption = "...Now Adjust Right Ear"

            lblEar.visible = True
            chkClick.Value = 0
            Form1.SetFocus
            Do While chkClick.Value = 0  'wait until the user clicks the knob
                If chkChange.Value = 1 Then 'user has made an adjustment
                    chkChange.Value = 0 'reset change flag
                    If (101 - CInt(txtValue.Text)) < 100 Then
                        If (101 - CInt(txtValue.Text)) > 0 Then
                            VScroll1(1).Value = (102 - CInt(txtValue.Text))
                        Else
                            VScroll1(1).Value = 1
                            txtValue.Text = 101
                        End If
                    Else
                        VScroll1(1).Value = 101
                        txtValue.Text = 1
                    End If
                End If
                DoEvents
            Loop
            lblEar.visible = False
            frmMono(0).visible = False
            frmMono(1).visible = False
            If c1 < 5 Then
                txtRIRightT1(OutputBox(c1)).Text = CInt(txtValue.Text)
            Else
                txtRIRightT2(OutputBox(c1)).Text = CInt(txtValue.Text)
            End If
        End If
        lblMainInstructions.Top = 144
        c1 = c1 + 1
    Loop
    Timer1.Enabled = False

    'output info
    intfilenumber = FreeFile ' This is safer than assigning a number
    Open (WorkingDir & WorkingFile) For Append As #intfilenumber
        Write #intfilenumber, ((txtRILeftT1(0) - 50 - 1) / 10), ((txtRIRightT1(0) - 50 - 1) / 10), "Trial1:5000Hz NBN"
        Write #intfilenumber, ((txtRILeftT1(1) - 50 - 1) / 10), ((txtRIRightT1(1) - 50 - 1) / 10), "Trial2:5000Hz NBN"
    Close #intfilenumber
    
    If CInt(txtLocalize.Text) = 2 Then   'tinnitus is in both ears
        RI5k = (((txtRILeftT1(0) - 50 - 1) / 10) + ((txtRIRightT1(0) - 50 - 1) / 10) + ((txtRILeftT1(1) - 50 - 1) / 10) + ((txtRIRightT1(1) - 50 - 1) / 10)) / 4
    ElseIf CInt(txtLocalize.Text) = 3 Then 'tinnitus is in right ear only
        RI5k = (((txtRIRightT1(0) - 50 - 1) / 10) + ((txtRIRightT1(1) - 50 - 1) / 10)) / 2
    Else 'tinnitus is in left ear only, txtLocalize.Text = 1
        RI5k = (((txtRILeftT1(0) - 50 - 1) / 10) + ((txtRILeftT1(0) - 50 - 1) / 10)) / 2
    End If

End Sub

Private Sub Command4_Click()
    Dim c9 As Integer
    For c9 = 0 To 7 Step 1
        lineClick(c9).visible = True
        lineClick(c9).ZOrder 0 'bring to front
    Next c9

    'MoveClick 632, 536
    timerClick.Enabled = True
End Sub

Private Sub Command5_Click()
'Call WriteReport
End Sub

Private Sub Command6_Click()
    Call OutputReport(22, 55, -1.2)
End Sub

Private Sub dirResume_Change()
    Dim intfilenumber, c1, c2 As Integer
    Dim TempString As String
    ChDir dirResume.Path
    cboResume.Clear 'clear the combobox
    txtLocalize.Text = ""
    txtIntensity.Text = ""
    txtBandwidth.Text = ""
    txtTemporal.Text = ""
    txtLoudness.Text = ""
    For c1 = 0 To 10
        txtLoudnessT1(c1).Text = ""
        txtLoudnessT2(c1).Text = ""
        txtPitchMatchT1(c1).Text = ""
        txtPitchMatchT2(c1).Text = ""
        txtPitchMatchT3(c1).Text = ""
    Next c1
    
    ' This code block checks if the file "MainData00.csv" exists in the specified directory.
    ' If the file exists, it clears the text boxes and reads the data from the file into the respective text boxes.
    ' The number of records in the file is counted and used to determine the available options in the combo box.
    ' The combo box is populated with different steps based on the number of records.
    ' If the file does not exist, no action is taken.
    If (dir(dirResume.Path & "\MainData00.csv")) = "MainData00.csv" Then 'the datafile exists!
        Call clearOldData 'clears the text boxes, which may be cluttered with bad values.
        cboResume.visible = True
        WorkingDir = dirResume.Path
        '^^^^^^^^^^^^The following code counts the number of records in the file^^^^^^^^^^^^^^^^^^^^^^^
        c1 = 0
        intfilenumber = FreeFile
        Open (dirResume.Path & "\MainData00.csv") For Input As #intfilenumber
        Do While Not EOF(intfilenumber)
            ' This code block reads data from a file and assigns the values to various text boxes based on the value of c1.
            ' The input file number is read using the Input statement.
            ' The TempString variable stores the value read from the file.
            ' The c1 variable is incremented by 1.
            ' The Select Case statement is used to determine which text box to assign the value to based on the value of c1.
            ' The corresponding text box is updated with the value of TempString.
            ' The loop continues until the end of the file is reached.
            Input #intfilenumber, TempString
            c1 = c1 + 1
            
            Select Case c1
                Case Is = 3 'localize data
                    txtLocalize.Text = TempString
                Case Is = 5 'sound intensity data
                    txtIntensity.Text = TempString
                Case Is = 7 'bandwidth data
                    txtIntensity2.Text = TempString
                Case Is = 9 'bandwidth data
                    txtBandwidth.Text = TempString
                Case Is = 11 'temporal data
                    txtTemporal.Text = TempString
                Case Is = 13 'loudness data
                    txtLoudness.Text = TempString
                Case Is = 19  'loudness matching data
                    txtLoudnessT1(0).Text = TempString
                Case Is = 20
                    txtLoudnessT2(0).Text = TempString
                Case Is = 23  'loudness matching data
                    txtLoudnessT1(1).Text = TempString
                Case Is = 24
                    txtLoudnessT2(1).Text = TempString
                Case Is = 27  'loudness matching data
                    txtLoudnessT1(2).Text = TempString
                Case Is = 28
                    txtLoudnessT2(2).Text = TempString
                Case Is = 31  'loudness matching data
                    txtLoudnessT1(3).Text = TempString
                Case Is = 32
                    txtLoudnessT2(3).Text = TempString
                Case Is = 35  'loudness matching data
                    txtLoudnessT1(4).Text = TempString
                Case Is = 36
                    txtLoudnessT2(4).Text = TempString
                Case Is = 39  'loudness matching data
                    txtLoudnessT1(5).Text = TempString
                Case Is = 40
                    txtLoudnessT2(5).Text = TempString
                Case Is = 43  'loudness matching data
                    txtLoudnessT1(6).Text = TempString
                Case Is = 44
                    txtLoudnessT2(6).Text = TempString
                Case Is = 47  'loudness matching data
                    txtLoudnessT1(7).Text = TempString
                Case Is = 48
                    txtLoudnessT2(7).Text = TempString
                Case Is = 51  'loudness matching data
                    txtLoudnessT1(8).Text = TempString
                Case Is = 52
                    txtLoudnessT2(8).Text = TempString
                Case Is = 55  'loudness matching data
                    txtLoudnessT1(9).Text = TempString
                Case Is = 56
                    txtLoudnessT2(9).Text = TempString
                Case Is = 59  'loudness matching data
                    txtLoudnessT1(10).Text = TempString
                Case Is = 60
                    txtLoudnessT2(10).Text = TempString
                Case Is = 68
                    txtPitchMatchT1(0).Text = TempString
                Case Is = 69
                    txtPitchMatchT2(0).Text = TempString
                Case Is = 70
                    txtPitchMatchT3(0).Text = TempString
                Case Is = 73
                    txtPitchMatchT1(1).Text = TempString
                Case Is = 74
                    txtPitchMatchT2(1).Text = TempString
                Case Is = 75
                    txtPitchMatchT3(1).Text = TempString
                Case Is = 78
                    txtPitchMatchT1(2).Text = TempString
                Case Is = 79
                    txtPitchMatchT2(2).Text = TempString
                Case Is = 80
                    txtPitchMatchT3(2).Text = TempString
                Case Is = 83
                    txtPitchMatchT1(3).Text = TempString
                Case Is = 84
                    txtPitchMatchT2(3).Text = TempString
                Case Is = 85
                    txtPitchMatchT3(3).Text = TempString
                Case Is = 88
                    txtPitchMatchT1(4).Text = TempString
                Case Is = 89
                    txtPitchMatchT2(4).Text = TempString
                Case Is = 90
                    txtPitchMatchT3(4).Text = TempString
                Case Is = 93
                    txtPitchMatchT1(5).Text = TempString
                Case Is = 94
                    txtPitchMatchT2(5).Text = TempString
                Case Is = 95
                    txtPitchMatchT3(5).Text = TempString
                Case Is = 98
                    txtPitchMatchT1(6).Text = TempString
                Case Is = 99
                    txtPitchMatchT2(6).Text = TempString
                Case Is = 100
                    txtPitchMatchT3(6).Text = TempString
                Case Is = 103
                    txtPitchMatchT1(7).Text = TempString
                Case Is = 104
                    txtPitchMatchT2(7).Text = TempString
                Case Is = 105
                    txtPitchMatchT3(7).Text = TempString
                Case Is = 108
                    txtPitchMatchT1(8).Text = TempString
                Case Is = 109
                    txtPitchMatchT2(8).Text = TempString
                Case Is = 110
                    txtPitchMatchT3(8).Text = TempString
                Case Is = 113
                    txtPitchMatchT1(9).Text = TempString
                Case Is = 114
                    txtPitchMatchT2(9).Text = TempString
                Case Is = 115
                    txtPitchMatchT3(9).Text = TempString
                Case Is = 118
                    txtPitchMatchT1(10).Text = TempString
                Case Is = 119
                    txtPitchMatchT2(10).Text = TempString
                Case Is = 120
                    txtPitchMatchT3(10).Text = TempString
                Case Is = 123
                    txtSoundThreshold(0).Text = TempString
                Case Is = 125
                    txtSoundThreshold(1).Text = TempString
                Case Is = 127
                    txtSoundThreshold(2).Text = TempString
                Case Is = 129
                    txtSoundThreshold(3).Text = TempString
                Case Is = 131
                    txtPA5ThreshValue.Text = TempString
                Case Is = 133
                    txtSoundLevelMatch(0).Text = TempString
                'Case Is = 135
                    txtSoundLevelMatch(1).Text = TempString
                'Case Is = 137
                    txtSoundLevelMatch(2).Text = TempString
                'Case Is = 139
                    txtSoundLevelMatch(3).Text = TempString
            End Select
        Loop
        'MsgBox "Found Data File. " & c1 & " records exist"
        Close #intfilenumber
        '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

        ' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
        ' This Select Case statement adds items to the cboResume ComboBox based on the value of c1.
        ' The items added depend on the range in which c1 falls.
        ' If c1 is less than or equal to 2, it adds "Step1_Localize" to the ComboBox.
        ' If c1 is between 3 and 6 (inclusive), it adds "Step1_Localize" and "Step2_SoundIntensity" to the ComboBox.
        ' If c1 is between 7 and 8 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", and "Step3_Bandwidth" to the ComboBox.
        ' If c1 is between 9 and 10 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", and "Step4_Temporal" to the ComboBox.
        ' If c1 is between 11 and 12 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", "Step4_Temporal", and "Step5_LoudnessRating" to the ComboBox.
        ' If c1 is between 13 and 14 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", "Step4_Temporal", "Step5_LoudnessRating", and "Step6_LoudnessMatching" to the ComboBox.
        ' If c1 is between 15 and 62 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", "Step4_Temporal", "Step5_LoudnessRating", "Step6_LoudnessMatching", and "Step7_PitchMatching" to the ComboBox.
        ' If c1 is between 63 and 122 (inclusive), it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", "Step4_Temporal", "Step5_LoudnessRating", "Step6_LoudnessMatching", "Step7_PitchMatching", and "Step8_Threshold" to the ComboBox.
        ' If c1 is greater than 122, it adds "Step1_Localize", "Step2_SoundIntensity", "Step3_Bandwidth", "Step4_Temporal", "Step5_LoudnessRating", "Step6_LoudnessMatching", "Step7_PitchMatching", and "Step8_Threshold" to the ComboBox.
        ' Uncomment the last line to add "Step9_ResidualInhibition" to the ComboBox when c1 is greater than 140.
        Select Case c1
            Case Is <= 2
                cboResume.AddItem "Step1_Localize"
            Case Is <= 6
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
            Case Is <= 8
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
            Case Is <= 10
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
            Case Is <= 12
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
                cboResume.AddItem "Step5_LoudnessRating"
            Case Is <= 14
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
                cboResume.AddItem "Step5_LoudnessRating"
                cboResume.AddItem "Step6_LoudnessMatching"
            Case Is <= 62
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
                cboResume.AddItem "Step5_LoudnessRating"
                cboResume.AddItem "Step6_LoudnessMatching"
                cboResume.AddItem "Step7_PitchMatching"
            Case Is <= 122
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
                cboResume.AddItem "Step5_LoudnessRating"
                cboResume.AddItem "Step6_LoudnessMatching"
                cboResume.AddItem "Step7_PitchMatching"
                cboResume.AddItem "Step8_Threshold"
            'Case Is <= 140
            Case Else
                cboResume.AddItem "Step1_Localize"
                cboResume.AddItem "Step2_SoundIntensity"
                cboResume.AddItem "Step3_Bandwidth"
                cboResume.AddItem "Step4_Temporal"
                cboResume.AddItem "Step5_LoudnessRating"
                cboResume.AddItem "Step6_LoudnessMatching"
                cboResume.AddItem "Step7_PitchMatching"
                cboResume.AddItem "Step8_Threshold"
                'cboResume.AddItem "Step9_ResidualInhibition"
        End Select
        
    Else
        'MsgBox "File Does not exist"
    End If
    
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
'
' Form_KeyDown is an event handler that is triggered when a key is pressed while the form has focus.
' It handles specific key codes and performs corresponding actions based on the key pressed.
'
' Parameters:
'   - KeyCode: An Integer representing the key code of the key that was pressed.
'   - Shift: An Integer representing the state of the Shift, Ctrl, and Alt keys when the key was pressed.
'
' Remarks:
'   - If the left arrow key (KeyCode = 37) is pressed and the value in the txtValue TextBox is greater than 1,
'     the value is decremented by 1 and the chkChange CheckBox is marked to indicate that an event has occurred.
'   - If the right arrow key (KeyCode = 39) is pressed and the value in the txtValue TextBox is less than 120,
'     the value is incremented by 1 and the chkChange CheckBox is marked to indicate that an event has occurred.
'   - If the value in the txtValue TextBox is already 120 or greater and the right arrow key is pressed,
'     and the VolAdj flag is set to True, the intMaxVolume variable is incremented by 1 and the chkChange CheckBox is marked.
'   - If the space bar (KeyCode = 32) is pressed and boolDblClick is False, the chkClick CheckBox is marked,
'     boolDblClick is set to True, and the timerDblClick Timer is enabled.
'   - If the 'c' key (KeyCode = 67 or 99) is pressed, the TinTrainComplete flag is set to True.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Then 'user hit left arrow
        If CInt(txtValue.Text) > 1 Then
            txtValue.Text = CInt(txtValue.Text) - 1
            chkChange.Value = 1 'mark that an event has occured
            'If VolAdj Then Call VolumeDown 'show volume up symbol
        End If
    ElseIf KeyCode = 39 Then 'user hit right arrow
        If CInt(txtValue.Text) < 120 Then
            txtValue.Text = CInt(txtValue.Text) + 1
            chkChange.Value = 1 'mark that an event has occured
            'If VolAdj Then Call VolumeUp 'show volume down symbol
        ElseIf CInt(txtValue.Text) >= 120 Then 'user has reached the max output volume
            If VolAdj Then 'call is being made from a volume adjusting step
                intMaxVolume = intMaxVolume + 1
                chkChange.Value = 1 'mark that an event has occured
            End If
        End If
    ElseIf KeyCode = 32 Then  'space bar
        If boolDblClick = False Then 'ok to accept click
            chkClick.Value = 1
            'imgCheck.visible = True
            'timerCheck.Enabled = True
            'Do While (timerCheck.Enabled = True)
            '    DoEvents
            'Loop
            boolDblClick = True
            timerDblClick.Enabled = True
        End If
    ElseIf KeyCode = 67 Or KeyCode = 99 Then 'user hit 'c to continue at the end of the tinnitus trainer
        TinTrainComplete = True
    End If
End Sub

' This event is triggered when the form is loaded.
' It initializes the form and sets the initial visibility of other forms.
' If an error occurs during the initialization process, an error message is displayed.
Private Sub Form_Load()
 
    VolAdj = False
    Form1.Hide
    FormReg.Hide
    formUserID.Hide
    Form2.Show
    Form2.SetFocus
    boolDblClick = False
    
    Exit Sub
    
    ErrHandler:
        MsgBox "ActiveLock Error: " & Err.Description
        FormReg.Show
        Form1.Hide
        formUserID.Hide
        Form2.Hide
        FormReg.SetFocus
End Sub

' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
' This event handler is triggered when the form is resized.
' It adjusts the position of various controls on the form to keep them centered.
Private Sub Form_Resize()
    On Error Resume Next
    
    ' Adjust the position of dialcontrol1
    dialcontrol1.Left = (Form1.ScaleWidth / 2) - (dialcontrol1.Width / 2)
    dialcontrol1.Top = (Form1.ScaleHeight / 2) - (dialcontrol1.Height / 2) + DialOffset
    
    ' Adjust the position of soundYesNo1
    soundYesNo1.Left = (Form1.ScaleWidth / 2) - (soundYesNo1.Width / 2) + 9
    soundYesNo1.Top = (Form1.ScaleHeight / 2) - (soundYesNo1.Height / 2) + YesNoTop
    
    ' Adjust the position of soundtypedial1
    soundtypedial1.Left = (Form1.ScaleWidth / 2) - (soundtypedial1.Width / 2) + 9
    soundtypedial1.Top = (Form1.ScaleHeight / 2) - (soundtypedial1.Height / 2) + SoundTypeDialTop
    
    ' Adjust the position of soundbandwidthdial1
    soundbandwidthdial1.Left = (Form1.ScaleWidth / 2) - (soundbandwidthdial1.Width / 2) - 25
    soundbandwidthdial1.Top = (Form1.ScaleHeight / 2) - (soundbandwidthdial1.Height / 2) + SoundBandwidthDialTop
    
    ' Adjust the position of whicheardial1
    whicheardial1.Left = (Form1.ScaleWidth / 2) - (whicheardial1.Width / 2) + 12
    whicheardial1.Top = (Form1.ScaleHeight / 2) - (whicheardial1.Height / 2) + WhichEarTop
    
    ' Move the cursor to the center of the form
    MoveClick (Form1.ScaleWidth / 2), (Form1.ScaleHeight / 2)
    
    ' Adjust the position of lblTitle
    lblTitle.Left = (Form1.ScaleWidth / 2) - (lblTitle.Width / 2)
    
    ' Adjust the position of lblMainInstructions
    lblMainInstructions.Left = (Form1.ScaleWidth / 2) - (lblMainInstructions.Width / 2)
    lblInstruct2.Left = (Form1.ScaleWidth / 2) - (lblInstruct2.Width / 2)
    lblInstruct3.Left = (Form1.ScaleWidth / 2) - (lblInstruct3.Width / 2)
    
    ' Adjust the position of framLoudness
    framLoudness.Left = (Form1.ScaleWidth / 2) - (framLoudness.Width / 2)
    framLoudness.Top = 500
    
    ' Adjust the position of lblNextSound
    lblNextSound.Left = (Form1.ScaleWidth / 2) - (lblNextSound.Width / 2)
    
    ' Adjust the position of lblSoft and lblLoud
    lblSoft.Left = (Form1.ScaleWidth / 2) - (lblSoft.Width / 2) - 200
    lblSoft.Top = (Form1.ScaleHeight / 2) - (lblSoft.Height / 2) + 55
    lblLoud.Left = (Form1.ScaleWidth / 2) - (lblLoud.Width / 2) + 200
    lblLoud.Top = (Form1.ScaleHeight / 2) - (lblLoud.Height / 2) + 55
    
    ' Adjust the position of frmBegin
    frmBegin.Left = (Form1.ScaleWidth / 2) - (frmBegin.Width / 2)
    
    ' Adjust the position of Choice1231 and PitchControl1
    Choice1231.Left = (Form1.ScaleWidth / 2) - (Choice1231.Width / 2) + 9
    Choice1231.Top = (Form1.ScaleHeight / 2) - (Choice1231.Height / 2) + 200
    PitchControl1.Left = (Form1.ScaleWidth / 2) - (PitchControl1.Width / 2) + 9
    PitchControl1.Top = (Form1.ScaleHeight / 2) - (PitchControl1.Height / 2) + 200
    
    ' Adjust the position of cmdTinTrain, cmdPrintReport, and cmdNext
    cmdTinTrain.Left = 24
    cmdTinTrain.Top = Form1.ScaleHeight - 100
    cmdPrintReport.Top = Form1.ScaleHeight - 100
    cmdNext.Top = Form1.ScaleHeight - 100
    cmdNext.Left = Form1.ScaleWidth - 113
    cmdPrintReport.Left = (Form1.ScaleWidth / 2) - (cmdPrintReport.Width / 2)
End Sub

' This event handler is triggered when the form is being unloaded.
' It plays a sound file and terminates the program.
Private Sub Form_Unload(Cancel As Integer)
    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    End
End Sub

End Sub

' This subroutine is triggered when the "optNew" option button is clicked.
' It hides the "dirResume" control, the "cboResume" control, and enables the "txtInitials" control.
' It also clears the text in the "txtInitials" control and disables the "cmdNext" button.
Private Sub optNew_Click()
    dirResume.Visible = False
    cboResume.Visible = False
    txtInitials.Enabled = True
    txtInitials.Text = ""
    cmdNext.Enabled = False
End Sub

' This subroutine is triggered when the "optResume" option button is clicked.
' It performs the following actions:
' - Makes the "dirResume" control visible.
' - Disables the "txtInitials" textbox.
' - Sets the path of the "dirResume" control to "C:\TinData".
' - Disables the "cmdNext" command button.
' - Sets the text of the "cboResume" combobox to "Resume From...".
' - Makes the "cboResume" combobox visible.
Private Sub optResume_Click()
    dirResume.visible = True
    txtInitials.Enabled = False
    dirResume.Path = "C:\TinData"
    cmdNext.Enabled = False
    cboResume.Text = "Resume From..."
    cboResume.visible = True
End Sub


' This subroutine is triggered by the Timer1 control's Timer event.
' It increments the value of txtTimer by 1.
Private Sub Timer1_Timer()
    txtTimer.Text = CInt(txtTimer.Text) + 1
End Sub

' This subroutine is triggered by the timerCheck timer event.
' It disables the timerCheck and hides the imgCheck control.
Private Sub timerCheck_Timer()
    'imgCheck.visible = False
    timerCheck.Enabled = False
End Sub

' This subroutine is triggered by the timerClick timer event.
' It sets the background color of various labels and frames to a light gray color.
' The timerClick timer is then disabled.
Private Sub timerClick_Timer()
    framLoudness.BackColor = &HF0F0E6
    lbl5.BackColor = &HF0F0E6
    lbl30.BackColor = &HF0F0E6       'original color:  &H00F0F0E6&
    lbl50.BackColor = &HF0F0E6
    lbl70.BackColor = &HF0F0E6
    lbl95.BackColor = &HF0F0E6
    lblScale.BackColor = &HF0F0E6
    
    lbl0(0).BackColor = &HF0F0E6
    lbl1(0).BackColor = &HF0F0E6
    lbl2(0).BackColor = &HF0F0E6
    lbl3(0).BackColor = &HF0F0E6
    lb4(0).BackColor = &HF0F0E6
    lblFive(0).BackColor = &HF0F0E6
    lblN1(0).BackColor = &HF0F0E6
    lblN2(0).BackColor = &HF0F0E6
    lblN3(0).BackColor = &HF0F0E6
    lblN4(0).BackColor = &HF0F0E6
    lblN5(0).BackColor = &HF0F0E6
    lblSofter(0).BackColor = &HF0F0E6
    lblGone(0).BackColor = &HF0F0E6
    lblNoChange(0).BackColor = &HF0F0E6
    lblLouder(0).BackColor = &HF0F0E6
    lblMuchLouder(0).BackColor = &HF0F0E6
    frmMono(0).BackColor = &HF0F0E6
    timerClick.Enabled = False
End Sub

' This subroutine is triggered when the timerDblClick interval elapses.
' It sets the boolDblClick variable to False and disables the timerDblClick.
Private Sub timerDblClick_Timer()
    boolDblClick = False
    timerDblClick.Enabled = False
End Sub

' TimerStep1_Timer event is triggered when the TimerStep1 control fires.
' It checks the value of whicheardial1 and if it is greater than or equal to 100,
' it sets the value of chkClick to 1.
Private Sub TimerStep1_Timer()
    If whicheardial1.getvalue >= 100 Then chkClick.Value = 1
End Sub

'------------------------------------------------------------------------------
' Sub: TimerStep2_Timer
'
' Description:
'   This subroutine is triggered by the TimerStep2 timer event. It updates the
'   value of the txtValue textbox with the current volume obtained from the
'   dialcontrol1 object. If the volume is within the range of 0 to 120, it sets
'   the attenuation level of the PA5x1 object and, if usePA52 is true, the
'   attenuation level of the PA5x2 object. If the volume is 999, it resets the
'   txtValue textbox to 0.
'
' Parameters:
'   None
'
' Returns:
'   None
'
'------------------------------------------------------------------------------
Private Sub TimerStep2_Timer()
    txtValue.Text = dialcontrol1.getvolume
    If CInt((txtValue.Text) >= 0) And CInt((txtValue.Text) <= 120) Then
        PA5x1.SetAtten (120 - CInt(txtValue.Text))
    ElseIf CInt(txtValue.Text) = 999 Then 'user can't hear sound?
'       MsgBox "Can you hear me now"
        txtValue.Text = 0
    End If
End Sub

' TimerStep6_Timer is an event handler that is triggered when the TimerStep6 timer interval elapses.
' It updates the value of the txtValue textbox with the current volume obtained from dialcontrol1.
' If the value is between 0 and 120 (inclusive), it sets the attenuation level of PA5x1 and PA5x2 accordingly.
' If the value is 999, it displays a message box indicating that the user can't hear sound and resets the txtValue to 0.
Private Sub TimerStep6_Timer()
    txtValue.Text = dialcontrol1.getvolume
    If CInt((txtValue.Text) >= 0) And CInt((txtValue.Text) <= 120) Then
        PA5x1.SetAtten (120 - CInt(txtValue.Text))
    ElseIf CInt(txtValue.Text) = 999 Then 'user can't hear sound?
        MsgBox "Can you hear me now"
        txtValue.Text = 0
    End If
End Sub

' This subroutine is triggered by the TimerStep8 timer event.
' It updates the value of the txtValue textbox with the current volume from the dialcontrol1.
' If the value is between 0 and 120 (inclusive), it sets the attenuation level of the PA5x1 and PA5x2 accordingly.
' If usePA52 is True, it also sets the attenuation level of the PA5x2.
Private Sub TimerStep8_Timer()

    txtValue.Text = dialcontrol1.getvolume
    
    If CInt((txtValue.Text) >= 0) And CInt((txtValue.Text) <= 120) Then
        PA5x1.SetAtten (120 - CInt(txtValue.Text))
    End If
End Sub

' This subroutine is triggered by the TimerVolume timer event.
' It hides the imgVolume(0) and imgVolume(1) controls and disables the TimerVolume.
Private Sub TimerVolume_Timer()
    imgVolume(0).Visible = False
    imgVolume(1).Visible = False
    TimerVolume.Enabled = False
End Sub

' This event handler is triggered when the text in the txtInitials textbox is changed.
' It checks if the textbox is empty and enables or disables the cmdNext button accordingly.
Private Sub txtInitials_Change()
    If txtInitials.Text = "" Then
        cmdNext.Enabled = False
    Else
        cmdNext.Enabled = True
    End If
End Sub

' RandomizeArray is a subroutine that shuffles the elements of an array in a random order.
' It takes an input array (ArrayIn) as a parameter and modifies it in place.
' The algorithm used is the Fisher-Yates shuffle algorithm.
'
' Parameters:
'   - ArrayIn: The input array to be shuffled.
'
' Example usage:
'   Dim myArray() As Variant
'   myArray = Array(1, 2, 3, 4, 5)
'   RandomizeArray myArray
'
'   After the execution of the subroutine, the elements in myArray will be randomly shuffled.
'
Private Sub RandomizeArray(ArrayIn As Variant)
   Dim c As Long
   Dim RandomIndex As Long
   Dim tmp As Variant
   Randomize
  'Ensure array was passed
   If VarType(ArrayIn) >= vbArray Then
        
     'loop through the array elements in reverse
      For c = UBound(ArrayIn) To LBound(ArrayIn) Step -1
      
        'select a random array index
         RandomIndex = Int((c - LBound(ArrayIn) + 1) * _
                       Rnd + LBound(ArrayIn))
                                    
        'c represents one array member
        'index, and RandomIndex represents
        'another, so swap the data held in
        'GayArray(c) with that in myarray(RandomIndex)
         tmp = ArrayIn(RandomIndex)
         ArrayIn(RandomIndex) = ArrayIn(c)
         ArrayIn(c) = tmp
         
      Next
   Else
     'The passed argument was not an
     'array; error handler goes here
   End If
End Sub

' This subroutine increases the volume by making the volume up button visible and starting the volume timer.
Private Sub VolumeUp()
    imgVolume(1).Visible = False
    imgVolume(0).Visible = True
    TimerVolume.Enabled = True
End Sub

' This subroutine decreases the volume by hiding the first volume image and showing the second volume image.
' It also enables the TimerVolume to start decreasing the volume gradually.
Private Sub VolumeDown()
    imgVolume(0).visible = False
    imgVolume(1).visible = True
    TimerVolume.Enabled = True
End Sub

'*******************************************************************************
' FUNCTION NAME: CanYouHearThis
'
' DESCRIPTION: This function is called when a user reaches the maximum output for volume.
'              It prompts the user to determine if they can hear a sound playing.
'              If the user can hear the sound, it prompts them to determine if the sound is quieter than their tinnitus.
'
' PARAMETERS:
'   - SoundToPlay: The path of the sound file to be played.
'
' RETURNS:
'   - Integer: 0 if the user cannot hear the sound or if the sound is quieter than their tinnitus,
'              1 if the user can hear the sound,
'              2 if the user can hear the sound and it is louder than their tinnitus.
'
'*******************************************************************************
Private Function CanYouHearThis(SoundToPlay As String) As Integer
    VolAdj = False
    lblMainInstructions.Caption = "Can you hear a sound playing right now?"
    lblMainInstructions.visible = True
    soundYesNo1.UserControl_Initialize
    'must initialize values to match the initalized control****
    txtValue.Text = 1
    CanYouHearThis = 1
    '******************
    soundYesNo1.visible = True
    soundYesNo1.SetFocus
    chkClick.Value = 0
    Do While chkClick.Value = 0  'wait until the user clicks the knob
        sndPlaySound SoundToPlay, SND_ASYNC Or SND_NODEFAULT
        txtTimer.Text = 0 'reset the textbox that holds the timer variable
        Do While (CInt(txtTimer.Text) < 30) 'loop for 4000ms
            If chkChange.Value = 1 Then 'user has made an adjustment
                chkChange.Value = 0 'reset change flag
                'MsgBox soundYesNo1.getvalue
                Select Case soundYesNo1.getvalue
                    Case Is = 0
                        'shpChoice4.BackColor = &H80FF80 'green
                        'shpChoice5.BackColor = &H80000000  'grey
                        txtValue.Text = 1
                        CanYouHearThis = 1
                    Case Is >= 1
                        'shpChoice4.BackColor = &H80000000  'grey
                        'shpChoice5.BackColor = &H80FF80 'green
                        txtValue.Text = 2
                        CanYouHearThis = 2
                End Select
                If chkClick.Value = 1 Then ' user clicked down so end early
                    txtTimer.Text = 40
                End If
            End If
            DoEvents
        Loop
        DoEvents
    Loop
    soundYesNo1.visible = False

    
    If CInt(txtValue.Text) = 2 Then 'user can hear sound
        
        chkChange.Value = 0
        chkClick.Value = 0
        lblMainInstructions.Caption = "Is this sound quieter than your tinnitus?"
        soundYesNo1.UserControl_Initialize
        'must initialize values to match the initalized control****
        txtValue.Text = 1
        CanYouHearThis = 0
        '******************
        soundYesNo1.visible = True
        soundYesNo1.SetFocus
        Do While chkClick.Value = 0  'wait until the user clicks the knob
            sndPlaySound SoundToPlay, SND_ASYNC Or SND_NODEFAULT
            txtTimer.Text = 0 'reset the textbox that holds the timer variable
            Do While (CInt(txtTimer.Text) < 30) 'loop for 3000ms
                If chkChange.Value = 1 Then 'user has made an adjustment
                    chkChange.Value = 0 'reset change flag
                    Select Case soundYesNo1.getvalue
                        Case Is = 0 'no
                            txtValue.Text = 1
                            CanYouHearThis = 0
                        Case Is >= 1 'yes
                            txtValue.Text = 2
                            CanYouHearThis = 2
                    End Select
                    If chkClick.Value = 1 Then ' user clicked down so end early
                        txtTimer.Text = 40
                    End If
                End If
                DoEvents
            Loop
            DoEvents
        Loop
        soundYesNo1.visible = False
    End If

    sndPlaySound "C:\TinData\tintest_wav\silence.wav", SND_ASYNC Or SND_NODEFAULT
    Call hide_all
    chkClick.Value = 0
    VolAdj = True
End Function

'*********************************************************************************
'This function will return a custom masker sound based on the likeness ratings in
'step 7 of the tinnitus tester.  It is only called in step 8 and 9.  It chooses a
'custom masker based on the tinnitus type (tonal, ringing or hissing) and the the
'pitch as rated by the user as sounding most like their tinnitus.
'
'If the user rated the loudness in step 6 of the tone most like their tinnitus to be
'not as loud as their tinnitus (code -102), then the program will select the next
'most-like pitch.
'
'If the user rates two values the same as most like their tinnitus, the program will
'select the lower frequency of the two
'
'If the user rates both 500hz and 5000hz as being most like thier tinnitus, the program
'will select 500Hz
'
'If the user rates 500Hz OR 5000Hz and a 2nd freq that is not 500 or 5000 the same
'as thier tinnitus, the program will select the freq that is not 500 or 5000
'
'If the user has Hissing tinnitus and rates either 500 or 5000Hz as most like thier
'tinnitus, proram will select next best value as 500hz and 5000Hz hiss are already
'pressented.
'*********************************************************************************
Private Function CustomMaskerPath() As String
    Dim MCustString As String 'string to hold custom masker name
    Dim LikenessRating(2, 11) As Integer
    Dim c1 As Integer
    Dim MostLike As Integer
    Dim cValue As Integer

    'First we have to choose the prefix - This is based on the tinnitus type selected:
        Select Case CInt(txtBandwidth.Text)
            Case Is = 1 'hissing
                MCustString = "W"
            Case Is = 2 'ringing
                MCustString = "R"
            Case Is = 3 ' tonal
                MCustString = "S"
        End Select

    'Second, we must determine which sound the user selected as being most like their tinnitus
    For c1 = 0 To 10 'populate array with average values from Likeness Rating
        'first we'll check to see if they've rated any of the sounds as being quiter than their tinnitus
        If (CInt(txtLoudnessT1(c1).Text) = -102) Or (CInt(txtLoudnessT2(c1).Text) = -102) Then
            'Enter a value of 0.  This will keep it sorted at the very bottom of the list
            LikenessRating(0, c1) = 0
            LikenessRating(1, c1) = c1 + 1 'holds filename value - necessary to preserve after sort
        ElseIf (CInt(txtLoudnessT2(c1).Text) = -101) Or (CInt(txtLoudnessT2(c1).Text) = -101) Then
            'user never heard sound at all so also keep it at 0 for the sort
            LikenessRating(0, c1) = 0
            LikenessRating(1, c1) = c1 + 1 'holds filename value - necessary to preserve after sort
        ElseIf ((MCustString = "w") And c1 = 0) Or ((MCustString = "w") And c1 = 5) Then
            'the user has hissing tinnitus.  Since they already will recieve a 500Hz hiss and a 5000Hz hiss, we don't want to play it twice, so set the value at 0 for
            'sorting purposes
            LikenessRating(0, c1) = 0
            LikenessRating(1, c1) = c1 + 1 'holds filename value - necessary to preserve after sort
        Else 'store actual likeness value
            LikenessRating(0, c1) = CInt((CInt(txtPitchMatchT1(c1).Text) + CInt(txtPitchMatchT2(c1).Text) + CInt(txtPitchMatchT3(c1).Text)) / 3) 'holds likeness rating
            LikenessRating(1, c1) = c1 + 1 'holds filename value - necessary to preserve after sort
        End If
    Next c1
    
    'sort the sounds
    MyQuickSort_Single LikenessRating, 0, 10, 0, False 'array to be sorted, start value, end value, element to sort, ascending/descending)

    cValue = 1
    'now check to see if two values are the same
    If LikenessRating(0, 0) = LikenessRating(0, 1) Then 'top 2 values are the same
        cValue = 2
        If LikenessRating(0, 1) = LikenessRating(0, 2) Then 'top 3 values are the same
            cValue = 3
            If LikenessRating(0, 2) = LikenessRating(0, 3) Then 'top 4 values are the same
                cValue = 4
                If LikenessRating(0, 3) = LikenessRating(0, 4) Then 'top 5 values are the same
                    cValue = 5
                    If LikenessRating(0, 4) = LikenessRating(0, 5) Then 'top 6 values are the same
                        cValue = 6
                        If LikenessRating(0, 5) = LikenessRating(0, 6) Then 'top 7 values are the same
                            cValue = 7
                            If LikenessRating(0, 6) = LikenessRating(0, 7) Then 'top 8 values are the same
                                cValue = 8
                                If LikenessRating(0, 7) = LikenessRating(0, 8) Then 'top 9 values are the same
                                    cValue = 9
                                    If LikenessRating(0, 8) = LikenessRating(0, 9) Then 'top 10 values are the same
                                        cValue = 10
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        'now we must go through all values and find the lowest frequency
        If ((LikenessRating(1, 0) = 1) And (LikenessRating(1, 1) = 6)) Or ((LikenessRating(1, 0) = 6) And (LikenessRating(1, 1) = 1)) Then 'both top values are 500 and 5000
            MostLike = 0 'choose 500 as base sound
        ElseIf ((LikenessRating(1, 0) = 1) Or (LikenessRating(1, 0) = 6)) Then 'one of the top two spots is 500 or 5000 so choose the other
            MostLike = LikenessRating(1, 1)
        ElseIf ((LikenessRating(1, 1) = 1) Or (LikenessRating(1, 1) = 6)) Then 'one of the top two spots is 500 or 5000 so choose the other
            MostLike = LikenessRating(1, 0)
        Else
            MostLike = LikenessRating(1, 0) 'start with the highest rated likeness frequency
        End If
        
        If cValue > 1 Then
            For c1 = 2 To cValue Step 1
                If (LikenessRating(1, c1) < MostLike) And (LikenessRating(1, c1) <> 1) And (LikenessRating(1, c1) <> 6) Then
                    MostLike = LikenessRating(1, c1)
                End If
            Next c1
        End If
    Else 'only 1 top value
        MostLike = LikenessRating(1, 0)
    End If
    MCustString = MCustString & CStr(MostLike)
    CustomMaskerPath = MCustString
End Function

' ***********************************************
'               Multidimensional Array sorted on a single dimensions
' ***********************************************
' Name: MyQuickSort_Single
' Description: Sorts a multidimensional array based on a single dimension using the QuickSort algorithm.
' Parameters:
'   - SortArray: The array to be sorted.
'   - First: The index of the first element to be sorted.
'   - Last: The index of the last element to be sorted.
'   - PrimeSort: The index of the dimension to be sorted.
'   - Ascending: Boolean value indicating whether the sorting should be in ascending order (True) or descending order (False).
' Returns: None
Private Sub MyQuickSort_Single(ByRef SortArray As Variant, ByVal First As Long, ByVal Last As Long, _
                                                            ByVal PrimeSort As Integer, ByVal Ascending As Boolean)
    Dim Low As Long, High As Long
    Dim temp As Variant, List_Separator As Variant
    Dim TempArray() As Variant
    ReDim TempArray(UBound(SortArray, 1))
    Low = First
    High = Last
    List_Separator1 = SortArray(PrimeSort, (First + Last) / 2)
    Do
        If Ascending = True Then
            Do While (SortArray(PrimeSort, Low) < List_Separator1)
                Low = Low + 1
            Loop
            Do While (SortArray(PrimeSort, High) > List_Separator1)
                High = High - 1
            Loop
        Else
            Do While (SortArray(PrimeSort, Low) > List_Separator1)
                Low = Low + 1
            Loop
            Do While (SortArray(PrimeSort, High) < List_Separator1)
                High = High - 1
            Loop
        End If
        If (Low <= High) Then
            For i = LBound(SortArray, 1) To UBound(SortArray, 1)
                TempArray(i) = SortArray(i, Low)
            Next
            For i = LBound(SortArray, 1) To UBound(SortArray, 1)
                SortArray(i, Low) = SortArray(i, High)
            Next
            For i = LBound(SortArray, 1) To UBound(SortArray, 1)
                SortArray(i, High) = TempArray(i)
            Next
            Low = Low + 1
            High = High - 1
        End If
    Loop While (Low <= High)
    If (First < High) Then MyQuickSort_Single SortArray, First, High, PrimeSort, Ascending
    If (Low < Last) Then MyQuickSort_Single SortArray, Low, Last, PrimeSort, Ascending
End Sub

Private Sub MoveClick(xc As Integer, yc As Integer)
    Dim c5 As Integer
    lineClick(0).x1 = xc - 21
    lineClick(0).X2 = xc - 5
    lineClick(0).Y1 = yc - 8
    lineClick(0).Y2 = yc

    lineClick(1).x1 = xc - 21
    lineClick(1).X2 = xc - 5
    lineClick(1).Y1 = yc
    lineClick(1).Y2 = yc
    
    lineClick(2).x1 = xc + 3
    lineClick(2).X2 = xc + 19
    lineClick(2).Y1 = yc
    lineClick(2).Y2 = yc
    
    lineClick(3).x1 = xc + 3
    lineClick(3).X2 = xc + 19
    lineClick(3).Y1 = yc
    lineClick(3).Y2 = yc + 8
    
    lineClick(4).x1 = xc - 21
    lineClick(4).X2 = xc - 5
    lineClick(4).Y1 = yc + 8
    lineClick(4).Y2 = yc
    
    lineClick(5).x1 = xc + 3
    lineClick(5).X2 = xc + 19
    lineClick(5).Y1 = yc
    lineClick(5).Y2 = yc - 8
    
    lineClick(6).x1 = xc
    lineClick(6).X2 = xc
    lineClick(6).Y1 = yc - 16
    lineClick(6).Y2 = yc - 4
    
    lineClick(7).x1 = xc
    lineClick(7).X2 = xc
    lineClick(7).Y1 = yc + 4
    lineClick(7).Y2 = yc + 16
End Sub

'---------------------------------------------------------------------------
' BubbleSortArray
'
' Description: Sorts a numeric array using the Bubble Sort algorithm.
'
' Parameters:
'   - NumericArray: ByRef Variant. The array to be sorted.
'
' Remarks:
'   - This code is based on the Bubble Sort algorithm found at:
'     http://www.freevbcode.com/ShowCode.Asp?ID=580
'   - The Bubble Sort algorithm repeatedly steps through the array,
'     compares adjacent elements, and swaps them if they are in the wrong order.
'     This process is repeated until the array is sorted.
'   - The sorted array is modified in-place, meaning the original array is modified.
'   - If an error occurs during the sorting process, the code will exit without
'     returning a value.
'
' Example usage:
'   Dim arr() As Variant
'   arr = Array(5, 2, 8, 1, 9)
'   BubbleSortArray arr
'
'---------------------------------------------------------------------------
Private Sub BubbleSortArray(ByRef NumericArray As Variant)
'http://www.freevbcode.com/ShowCode.Asp?ID=580
    Dim vAns As Variant
    Dim vTemp As Variant
    Dim bSorted As Boolean
    Dim lCtr As Long
    Dim lCount As Long
    Dim lStart As Long

    lStart = LBound(NumericArray)
    lCount = UBound(NumericArray)

    bSorted = False
   
    Do While Not bSorted
      bSorted = True

      For lCtr = lCount - 1 To lStart Step -1
        If NumericArray(lCtr + 1) < NumericArray(lCtr) Then
          DoEvents
          bSorted = False
           vTemp = NumericArray(lCtr)
           NumericArray(lCtr) = NumericArray(lCtr + 1)
           NumericArray(lCtr + 1) = vTemp
         End If
      Next lCtr
      
    Loop
    
'BubbleSortArray = vAns
Exit Sub

ErrorHandler:
'BubbleSortArray = vbEmpty
Exit Sub
End Sub


' FILEPATH: /C:/codedev/auricle/TinnTester/TinTest -UofM/TinTest.frm
' This code block defines variables and assigns values to an array of strings.
' It also contains paragraphs of text stored in the array for generating a report.
Sub OutputReport1(TL As Integer, LM1 As Integer, RI5 As Single)

    Dim oExcel As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim oRng1 As Excel.Range
    Dim oRng2 As Excel.Range
    Dim ThisLine As Excel.Shape
    Dim text1(1 To 100) As String
    Dim TL1 As Single
    Dim x11 As Integer 'counter
    Dim LMP As Integer ' holds where to plot loudness match data
    Dim TnSp(1 To 11) As Single 'holds the calculated tinnitus spectrum values

    ' TL = tinnitus loudness rating
    ' LM1 = loudness match at 1000Hz
    TL1 = (TL / 10) + 0.5 ' convert TL, which is entered on a scale of 0 to 100 to fit plot, 0 to 10

    ' The array "text1" stores paragraphs of text for generating a report.
    ' Each element of the array represents a paragraph.

    '---SECTION 1------
    '---paragraph 1----
    text1(1) = "This report describes the properties of your tinnitus based the evaluation conducted here at the clinic."
    text1(2) = "We also compare your results to those of a sample of up to 74 tinnitus patients evaluated in the"
    text1(3) = "Human Neural Plasticity Laboratory at McMaster University in Hamilton, Ontario, Canada."
    '---paragraph 1----

    '---paragraph 2----
    text1(4) = "Two important attributes of all sounds (including tinnitus sounds) are (1) the loudness of the sound,"
    text1(5) = "and (2) the frequency or pitch of the sound.  Loudness is measured in a unit called 'decibels'(dB), "
    text1(6) = "and pitch or frequency by a unit called 'Hertz' (Hz).  A loudness of 60 dB sound corresponds approx-"
    text1(7) = "imately to the loudness of normal speech. Middle 'C' on the piano corresponds to a pitch or frequency"
    text1(8) = "of 256 Hz.  The frequencies contained in most speech sounds fall in the range of 100 - 2000 Hz,"
    text1(9) = "while very high pitched sounds heard by the human ear can range as high as 20,000 Hz.  However,"
    text1(10) = "sounds above about 12,000 Hz comparatively rare in the natural human environment and not"
    text1(11) = "everyone can detect sounds above this frequency."
    '---paragraph 2----

    '---paragraph 3----
    text1(12) = "We measured your tinnitus with two independent methods.  The graphs below show your results"
    text1(13) = "compared to 74 patients with stable chronic tinnitus measured at McMaster University.  We call these"
    text1(14) = "74 patients our 'baseline' sample."
    text1(15) = "In the first method, you rated your tinnitus on a Borg CR100 scale which is used in tinnitus perception."
    text1(16) = "Your Tinnitus Loudness Rating on a Borg CR100 scale was _" & TL & "_, out of a maximum of _100_."
    '---paragraph 3----

    '---paragraph 4----
    text1(17) = "Figure 1 on the next page compares your Borg CR100 loudness rating to that of the baseline sample. "
    text1(18) = "The baseline sample reported an average rating of 43.9, which corresponds to the midpoint "
    text1(19) = "between 'moderate' to 'strong' tinnitus on the Borg CR100 scale."
    '---paragraph 4----

    '---paragraph 5----
    text1(20) = "In the second method,  you adjusted the loudness of several sounds to equal the loudness of your"
    text1(21) = "tinnitus.  The loudness match you gave for a sound of 1000 Hz (a high pitched tone) was " & LM1 & " dB"
    text1(22) = "Figure 2 below shows how your loudness match in dB compares to that of the baseline sample"
    text1(23) = "measured at McMaster University. "
    '---paragraph 5----

    '---SECTION 2------
    '---paragraph 1----
    text1(24) = "The next step of tinnitus measurement asked you to rate the similarity of each of several tones"
    text1(25) = "differing in pitch for their similarity or 'likeness' to your tinnitus.  Your results are shown below in"
    text1(26) = "Figure 3 and are compared to a baseline  group of tinnitus subjects measured at McMaster University."
    text1(27) = "From our research we consider any likeness rating above 40 to mean that the sound is beginning to"
    text1(28) = "resemble your tinnitus.  Ratings above this value are your 'tinnitus spectrum'."
    '---paragraph 1----

    '---SECTION 3------
    '---paragraph 1----
    text1(29) = "Some tinnitus patients report that sounds in the environment can �mask� their tinnitus.  This means "
    text1(30) = "that when the sound is present they cannot hear their tinnitus.   Examples of sounds that can mask "
    text1(31) = "tinnitus are a noisy mistuned radio, the sound of running water, or nature sounds such as crickets or "
    text1(32) = "birds.   "
    '---paragraph 1----

    '---paragraph 2----
    text1(33) = "When masking sounds are presented in the laboratory and then switched off, tinnitus may be reduced"
    text1(34) = "or even eliminated for a brief period of time.  This phenomenon is called �Residual Inhibition� or RI.  "
    text1(35) = "RI typically lasts about 30 seconds to a minute, but it can last longer."
    '---paragraph 2----

    '---paragraph 3----
    text1(36) = "Figure 4 below shows RI induced in 47 tinnitus patients by a noise-like masking sound with a center "
    text1(37) = "frequency of 5000 Hz.   Because this masking sound contains the frequencies usually reported in the "
    text1(38) = "tinnitus spectrum, it is more effective than most other maskers at inducing RI.  RI is measured on a "
    text1(39) = "scale ranging from 0 (meaning tinnitus did not change after listening to the masker) to minus 5 "
    text1(40) = "(meaning the tinnitus was gone when the sound was switched off).  In some cases tinnitus can get "
    text1(41) = "louder (a score of plus 5 means much louder).   "
    '---paragraph 3----

    '---paragraph 4----
    text1(42) = "Your RI score measured with this masker was _" & RI5 & "_.  The graph below shows how your score compares "
    text1(43) = "to that of 47 people measured at McMaster University.  "
    '---paragraph 4----
    
    ' This code sets up the formatting for the selected range in the Excel worksheet.
    ' It merges the cells, sets the font style, size, and boldness, and applies interior formatting.
    ' The range is specified as "A1:I1" and the header text "Tinnitus Report" is added.
    ' The formatting is applied to the merged range.
    ' The range is then centered horizontally and aligned to the bottom vertically.
    ' The wrap text, orientation, and other formatting options are set.
    ' Finally, the font color, background color, and other visual properties are set.
    Set oExcel = New Excel.Application
    Set oWB = oExcel.Workbooks.Add
    Set oWS = oWB.Worksheets("Sheet1")
    Set oRng1 = oWS.Range("A1")
    Set oRng2 = oWS.Range("B2:E5")

    oExcel.visible = False ' <-- *** Optional *** 'true = show actions. False = don't show actions
    'Set up header info:
    '-----------------------------------------'
    oWS.Range("A1").Value = "Tinnitus Report"
    oWS.Range("A1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Font
        .Name = "Calibri"
        .SIZE = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    '-----------------------------------------'
    ' Purpose: Set up username/info/Data in the selected range of the worksheet.
    '-----------------------------------------'
    oWS.Range("A3").Value = "Name:"
    oWS.Range("B3").Value = UserName
    oWS.Range("A4").Value = "Report Date:"
    oWS.Range("B4").Value = Now()
    oWS.Range("F3").Value = "City:"
    oWS.Range("G3").Value = UserCity
    oWS.Range("F4").Value = "State/Prov:"
    oWS.Range("G4").Value = UserProv
    oWS.Range("F5").Value = "Country:"
    oWS.Range("G5").Value = UserCountry

    oWS.Range("A7").Value = "Self-Reported Data:"
    oWS.Range("A8").Value = "Tinitus Location:"
    oWS.Range("B8").Value = UserTL
    oWS.Range("A9").Value = "Steady or Pulsing:"
    oWS.Range("B9").Value = UserSorP
    oWS.Range("A10").Value = "Bandwidth:"
    oWS.Range("B10").Value = UserBW

    oWS.Range("F8").Value = "Age:"
    oWS.Range("G8").Value = UserAge
    oWS.Range("F9").Value = "Sex:"
    oWS.Range("G9").Value = UserSex
    oWS.Range("F10").Value = "YY/MM of Onset:"
    oWS.Range("G10").Value = UserOnset
    '-----------------------------------------'
    
    'insert text into body of excel file:

    oWS.Range("A12").Value = text1(1)
    oWS.Range("A13").Value = text1(2)
    oWS.Range("A14").Value = text1(3)
    
    oWS.Range("A16").Value = text1(4)
    oWS.Range("A17").Value = text1(5)
    oWS.Range("A18").Value = text1(6)
    oWS.Range("A19").Value = text1(7)
    oWS.Range("A20").Value = text1(8)
    oWS.Range("A21").Value = text1(9)
    oWS.Range("A22").Value = text1(10)
    oWS.Range("A23").Value = text1(11)
    
    oWS.Range("A25").Value = "(1) YOUR TINNITUS LOUDNESS"
    oWS.Range("A25").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    oWS.Range("A27").Value = text1(12)
    oWS.Range("A28").Value = text1(13)
    oWS.Range("A29").Value = text1(14)
    oWS.Range("A31").Value = text1(15)
    oWS.Range("A32").Value = text1(16)
    oWS.Range("A34").Value = text1(17)
    oWS.Range("A35").Value = text1(18)
    oWS.Range("A36").Value = text1(19)
    oWS.Range("A65").Value = text1(20)
    oWS.Range("A66").Value = text1(21)
    oWS.Range("A68").Value = text1(22)
    oWS.Range("A69").Value = text1(23)
    
'plot Loudness Rating:
    ' This code adds a chart to the active sheet and customizes its properties.
    ' The chart type is set to a clustered column chart.
    ' The chart title is set to "Tinnitus Loudness Rating".
    ' Two series are added to the chart: "Avg Tinnitus Loudness" and "Your Tinnitus Loudness".
    ' The X values for the first series are set to an array of values.
    ' The Y values for both series are set to arrays of values.
    ' The position of the chart is adjusted to align with cell B48.
    ' An image is inserted and positioned below the chart.
    ' The text "Figure 1" is added below the image with italic formatting and font size 10.
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Tinnitus Loudness Rating"
    ActiveChart.SeriesCollection(1).Name = "=""Avg Tinnitus Loudness"""
    ActiveChart.SeriesCollection(1).XValues = "={5,15,25,35,45,55,65,75,85,95}"
    'ActiveChart.SeriesCollection(1).XValues = "= {""Extremely Weak"","""",""Moderate"","""",""Strong"","""",""Very Strong"","""","""",""Extremely Strong""}"
    ActiveChart.SeriesCollection(1).Values = "={0,0.02,0.12,0.28,0.19,0.17,0.04,0.11,0.04,0}"
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your Tinnitus Loudness"""
    ActiveChart.SeriesCollection(2).Values = "={0.225}"
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    'ActiveChart.SeriesCollection(2).XValues = "=Sheet1!$C$53"
    ActiveChart.SeriesCollection(2).XValues = TL1
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -58.5
    With ActiveChart.Parent
        .Top = Range("B48").Top
        .Left = Range("B48").Left
    End With
    'insert pic
    oWS.Range("B62").Select
    ActiveSheet.Pictures.Insert("C:\TinData\axis1.JPG").Select
    With Selection.ShapeRange
       .Top = Range("B60").Top + 2
       .Left = Range("B60").Left + 34
       .Height = Application.InchesToPoints(0.36)
       .Width = Application.InchesToPoints(2.97)
    End With
    oWS.Range("E63").Value = "Figure 1"
    oWS.Range("E63").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10

'plot Loudness Match - New Plot - based on 1khz level match (created Jan 2011):
    ' This code generates a chart to display tinnitus loudness match data.
    ' It calculates the position of the user's loudness on the chart based on the LM1 value.
    ' The chart displays the average loudness values for different loudness ranges,
    ' as well as the user's loudness value.
    ' The chart is positioned on the worksheet and labeled with appropriate titles.
    ' A figure number is added to the worksheet for reference.
    ' The code also includes commented out lines for inserting a picture into the worksheet.
    ' Note: The file path mentioned in the comments is for reference and may need to be updated.
    'first, we have have to calculate where to put users' loudness in the plot:
    Select Case LM1
        Case Is < 10
            LMP = 1
        Case Is <= 20
            LMP = 2
        Case Is <= 30
            LMP = 3
        Case Is <= 40
            LMP = 4
        Case Is <= 50
            LMP = 5
        Case Is <= 60
            LMP = 6
        Case Is > 60
            LMP = 7
    End Select
    oWS.Range("A72:G73").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$73:$G$73")
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.Axes(xlValue).MajorGridlines.Select
    ActiveSheet.ChartObjects("Chart 3").Activate
    
    ActiveChart.SeriesCollection(1).Values = "={0.152777778,0.236111111,0.236111111,0.180555556,0.111111111,0.041666667,0.041666667}"
    ActiveChart.SeriesCollection(1).Name = "=""Avg Loudness"""
    'ActiveChart.SeriesCollection(1).XValues = "={1,2,3,4,5,6,7,8,9,10}"
    ActiveChart.SeriesCollection(1).XValues = "={""<10"",""11 to 20"",""21 to 30"",""31 to 40"",""41 to 50"",""51 to 60"","">60""}"
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Tinnitus Loudness Match"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your Loudness"""
    ActiveChart.SeriesCollection(2).Values = "={.17}"
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    ActiveChart.SeriesCollection(2).XValues = LMP
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -58.5
    With ActiveChart.Parent
        .Top = Range("B71").Top
        .Left = Range("B71").Left
    End With
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Loudness in dB SL"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    'insert pic
    'oWS.Range("B62").Select
    'ActiveSheet.Pictures.Insert("C:\TinData\axis1.JPG").Select
    'With Selection.ShapeRange
    '   .Top = Range("B60").Top + 2
    '   .Left = Range("B60").Left + 34
    '   .Height = Application.InchesToPoints(0.36)
    '   .Width = Application.InchesToPoints(2.97)
    'End With
    oWS.Range("E86").Value = "Figure 2"
    oWS.Range("E86").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10


    'SECTION 2
        ' Set the value of cell A94 to "(2)  YOUR TINNITUS SPECTRUM"
        oWS.Range("A94").Value = "(2)  YOUR TINNITUS SPECTRUM"
        
        ' Select cell A94
        oWS.Range("A94").Select
        
        ' Apply formatting to the selected cell
        Selection.Font.Bold = True
        Selection.Font.SIZE = 12
        Selection.Font.Underline = xlUnderlineStyleSingle
        
        ' Set the values of cells A96, A97, A98, A100, and A101 to the corresponding elements in the text1 array
        oWS.Range("A96").Value = text1(24)
        oWS.Range("A97").Value = text1(25)
        oWS.Range("A98").Value = text1(26)
        oWS.Range("A100").Value = text1(27)
        oWS.Range("A101").Value = text1(28)
    
    'PLOT LIKENESS
        'first, calculate users avg likeness values:
        For x11 = 1 To 11 Step 1
            TnSp(x11) = (CInt(txtPitchMatchT1(x11 - 1).Text) + CInt(txtPitchMatchT2(x11 - 1).Text) + CInt(txtPitchMatchT3(x11 - 1).Text)) / 3
        Next x11
        ActiveSheet.Shapes.AddChart.Select
        ActiveChart.SetSourceData Source:=Range("Sheet1!$A$104:$K$105")
        ActiveChart.ChartType = xlColumnClustered
        ActiveChart.SeriesCollection(1).Values = "={15,21.8,31.9,42.5,45.7,56.5,58.4,54.5,54,54.3,29.8}"
        ActiveChart.SeriesCollection(1).XValues = "={.5,1,2,3,4,5,6,7,8,10,12}"
        ActiveChart.SeriesCollection(1).Name = "=""Avg Spectrum"""
        ActiveChart.SeriesCollection(1).ChartType = xlXYScatterLines
        ActiveChart.HasTitle = True
        ActiveChart.ChartTitle.Text = "Tinnitus Likeness Match"
        With ActiveChart
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Frequency (kHz)"
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Tinnitus Spectrum"
        End With
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(2).Name = "=""Your Spectrum"""
        'ActiveChart.SeriesCollection(2).Values = "={5,   10,  40,  48,  65,  85,  90,  95,  70,  42,  12}"
        ActiveChart.SeriesCollection(2).Values = TnSp
        ActiveChart.SeriesCollection(2).ChartType = xlXYScatterLines
        ActiveChart.SeriesCollection(2).XValues = "={.5,1,2,3,4,5,6,7,8,10,12}"
        ActiveChart.Axes(xlCategory).MaximumScale = 13
        ActiveChart.Axes(xlCategory).MinimumScale = 0
        ActiveChart.Axes(xlCategory).MajorUnit = 1
        With ActiveChart.Parent
            .Top = Range("B103").Top
            .Left = Range("B103").Left
        End With
        oWS.Range("E118").Value = "Figure 3"
        oWS.Range("E118").Select
        Selection.Font.Italic = True
        Selection.Font.SIZE = 10
        'PLOT LIKENESS
    
    
'RI data
    oWS.Range("A122").Value = "(3) RESIDUAL INHIBITION"
    oWS.Range("A122").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    
    'INSERT TEXT HERE
    oWS.Range("A125").Value = text1(29)
    oWS.Range("A126").Value = text1(30)
    oWS.Range("A127").Value = text1(31)
    oWS.Range("A128").Value = text1(32)
    
    oWS.Range("A131").Value = text1(33)
    oWS.Range("A132").Value = text1(34)
    oWS.Range("A133").Value = text1(35)
    
    oWS.Range("A142").Value = text1(36)
    oWS.Range("A143").Value = text1(37)
    oWS.Range("A144").Value = text1(38)
    oWS.Range("A145").Value = text1(39)
    oWS.Range("A146").Value = text1(40)
    oWS.Range("A147").Value = text1(41)
    
    oWS.Range("A150").Value = text1(42)
    oWS.Range("A151").Value = text1(43)
     
    'Plot RI Data
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$154:$K$154")
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SeriesCollection(1).Values = "={0, 0.021276596, 0,0.063829787,0.085106383,0.085106383,0.276595745,0.14893617,0.063829787,0.170212766,0.085106383}"
    'ActiveChart.SeriesCollection(1).XValues = "={1,2,3,4,5,6,7,8,9,10,11}"
    ActiveChart.SeriesCollection(1).XValues = "={""5"",""4"",""3"",""2"",""1"",""0"",""-1"",""-2"",""-3"",""-4"",""-5""}"
    ActiveChart.SeriesCollection(1).Name = "=""Avg RI"""
    ActiveChart.SeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Residual Inhibition"
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Residual Inhibition"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your RI"""
    ActiveChart.SeriesCollection(2).Values = RI5
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    '**********NEED TO CHANGE TO ACTUAL USER INPUT
    '**********NEED TO CHANGE TO ACTUAL USER INPUT
    '**********NEED TO CHANGE TO ACTUAL USER INPUT
    '**********NEED TO CHANGE TO ACTUAL USER INPUT
    ActiveChart.SeriesCollection(2).XValues = 8 '**********NEED TO CHANGE TO ACTUAL USER INPUT
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -58.5
    With ActiveChart.Parent
        .Top = Range("B155").Top
        .Left = Range("B155").Left
    End With
    Set ThisLine = ActiveChart.Shapes.AddLine(175, 25, 175, 160) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 2
    End With
    
    Set ThisLine = ActiveChart.Shapes.AddLine(95, 30, 60, 30) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BeginArrowheadStyle = 1 'none
        .EndArrowheadStyle = 2 'triangle

        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 1
    End With
    Set ThisLine = ActiveChart.Shapes.AddLine(255, 30, 290, 30) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BeginArrowheadStyle = 1
        .EndArrowheadStyle = 2 'msoArrowheadTriangle

        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 1
    End With

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Change in Tinnitus"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    oWS.Range("E170").Value = "Figure 4"
    oWS.Range("E170").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10
    
Cleanup:
    ActiveWorkbook.SaveAs FileName:=(WorkingDir & "\TinnitusReport.xlsx"), FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close
    Set oWB = Nothing
    
    oExcel.Quit
    Set oExcel = Nothing
End Sub

Sub OutputReport(TL As Integer, LM1 As Integer, RI5 As Single)
    Dim oExcel As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim oRng1 As Excel.Range
    Dim oRng2 As Excel.Range
    Dim ThisLine As Excel.Shape
    Dim text1(1 To 100) As String
    Dim TL1 As Single
    Dim x11 As Integer 'counter
    Dim LMP As Integer ' holds where to plot loudness match data
    Dim TnSp(1 To 11) As Single 'holds the calculated tinnitus spectrum values
    Dim i As Integer
    Dim w As Integer
    'TL = tinitus loudness rating
    'LM1 = loudness match at 1000Hz
    TL1 = (TL / 10) + 0.5 ' convert TL, which is entered on a scale of 0 to 100 to fit plot, 0 to 10
    
    '---SECTION 0------
    '---paragraph 1----
    'text1(1) = "This report describes your tinnitus based an evaluation conducted at the Montreal   "
    'text1(2) = "Tinnitus Clinic.  We also compare your results to those of a large sample of tinnitus"
    'text1(3) = "patients evaluated in the Human Neural Plasticity Laboratory at McMaster University in"
    'text1(4) = "Hamilton, Ontario, Canada."
    text1(1) = "This report describes your tinnitus based the evaluation just completed.  We also "
    text1(2) = "compare your results to those of a large sample of tinnitus patients evaluated in the "
    text1(3) = "Human Neural Plasticity Laboratory at McMaster University in Hamilton, Ontario, "
    text1(4) = "Canada."

    '---paragraph 1----
    
    '---paragraph 2----
    text1(5) = "Two important attributes of all sounds (including tinnitus sounds) are (1) the loudness of "
    text1(6) = "the sound, and (2) the frequency or pitch of the sound.  We measured both attributes of "
    text1(7) = "your tinnitus.  "
    '---paragraph 2----
    
    '---paragraph 3----
    text1(8) = "Loudness is measured in a unit called 'decibels' (dB), and pitch or frequency by a unit called "
    text1(9) = "Hertz' (Hz). A 60 dB sound corresponds approximately to the loudness of normal speech.  "
    text1(10) = "Middle 'C' on the piano corresponds to a pitch or frequency of 256 Hz. The frequencies "
    text1(11) = "contained in speech and in the normal environment typically range between 100 and 3000 Hz.  "
    text1(12) = "The human ear can hear pitches as high as 20,000 Hz, although not everyone can hear pitches "
    text1(13) = "that high.    "
    '---paragraph 3----
    
    '---paragraph 4----
    text1(14) = "We also measured an attribute of your tinnitus called 'Residual Inhibition' or RI.  RI is a "
    text1(15) = "temporary suppression of tinnitus that can sometimes be experienced after listening to a "
    text1(16) = "masking sound.  All three attributes of your tinnitus  (tinnitus loudness, pitch, and RI) are "
    text1(17) = "reported below."
    
    '---paragraph 4----
    '---SECTION 1------
    '---paragraph 1----
    text1(18) = "We measured your tinnitus loudness with two independent methods. The graphs below show your "
    text1(19) = "results for both methods, compared to 74 patients with stable chronic tinnitus measured at"
    text1(20) = "McMaster University.   We call these 74 patients our 'baseline' sample.  "
    '---paragraph 1----
    
    '---paragraph 2----
    text1(21) = "In the first method, you rated the loudness of your tinnitus on a Borg CR100 scale which is "
    text1(22) = "used in auditory research.   Scores on this scale range from 0 (tinnitus loudness rated "
    text1(23) = " 'extremely weak') to 100 (tinnitus loudness rated 'extremely loud').  The baseline sample "
    text1(24) = "reported an average rating of 43.9, which corresponds to the midpoint between 'moderate' "
    text1(25) = "to 'strong' tinnitus on the Borg CR100 scale. "
    '---paragraph 2----
    '---paragraph 3----
    text1(26) = "Your Tinnitus Loudness Rating on a Borg CR100 scale was _" & TL & "_, out of a maximum of 100.  "
    text1(27) = "Figure 1 below compares your Borg CR100 loudness rating to that of the baseline sample."
    '---paragraph 3----
    '---paragraph 4----
    text1(28) = "In the second method, you adjusted the loudness of several sounds measured in dB to equal"
    text1(29) = "the loudness of your tinnitus.   The loudness match you gave for a sound of 1000 Hz (a high "
    text1(30) = "pitched tone) was _" & LM1 & "_ dB.    Figure 2 below shows how your loudness match in dB"
    text1(31) = "compares to that of the baseline sample. Remember that a 60 dB sound corresponds"
    text1(32) = "approximately to the loudness of normal speech."
    '---paragraph 4----
    
    '---SECTION 2------
    '---paragraph 1----
    text1(33) = "Another step in our tinnitus measurement asked you to rate the similarity of each of several "
    text1(34) = "tones differing in pitch for their similarity or 'likeness' to your tinnitus. From our research we"
    text1(35) = "consider any likeness rating above 40 to indicate that the sound is beginning to resemble "
    text1(36) = "your tinnitus. Ratings above this value are your 'tinnitus spectrum'.  For most patients, "
    text1(37) = "the tinnitus spectrum contains pitches spanning the range 3000 to 10,000 Hz.  "
    '---paragraph 1----
    '---paragraph 2----
    text1(38) = "Your results are shown below in Figure 3 and are compared to a baseline group of tinnitus "
    text1(39) = "subjects measured at McMaster University.  "
    '---paragraph 2----
    
    '---SECTION 3------
    '---paragraph 1----
    text1(40) = "Some tinnitus patients report that sounds in the environment can 'mask' their tinnitus.  This"
    text1(41) = "means that when the sound is present they cannot hear their tinnitus.   Examples of sounds "
    text1(42) = "that can mask tinnitus are a noisy mistuned radio, the sound of running water, or nature "
    text1(43) = "sounds such as crickets or birds.   "

    '---paragraph 1----
    
    '---paragraph 2----
    text1(44) = "When masking sounds are presented in the laboratory and then switched off, tinnitus may "
    text1(45) = "be reduced or even eliminated for a brief period of time.  This phenomenon is called "
    text1(46) = "�Residual Inhibition� or RI.  RI typically lasts about 30 seconds to a minute, but it can last "
    text1(47) = "longer."
    '---paragraph 2----
    
    '---paragraph 3----
    text1(48) = "We measured your RI by presenting a noise-like masking sound with a center frequency of "
    text1(49) = "5000 Hz.  When the masker was switched off, we asked you to rate how your tinnitus had "
    text1(50) = "changed.   Because this masking sound contains the frequencies usually reported to be in "
    text1(51) = "the tinnitus spectrum, it is more effective than most other maskers at inducing RI.  RI was "
    text1(52) = "measured on a scale ranging from 0 (meaning tinnitus did not change after listening to the "
    text1(53) = "masker) to minus 5 (meaning the tinnitus was gone when the sound was switched off).  In "
    text1(54) = "some cases tinnitus can get louder (a score of plus 5 means much louder) meaning that no RI "
    text1(55) = "was experienced.   Not everyone will report RI (tinnitus suppression) with the masker we used,"
    text1(56) = "but many people with tinnitus do."
    '---paragraph 3----
    
    '---paragraph 4----
    text1(57) = "Your RI score measured with this masker was _" & RI5 & "_. The graph below shows how your"
    text1(58) = "score compares to that of 47 people with tinnitus measured at McMaster University.  "
    '---paragraph 4----
    
    '---SECTION 3------
    '---paragraph 1----
    text1(59) = "The measurements reported above are intended to help you understand your tinnitus and "
    text1(60) = "how it compares to tinnitus experienced by others.  Often it can be reassuring to see that "
    text1(61) = "while it is irritating, the tinnitus sound is not mysterious.  "
    '---paragraph 1----
    
    '---paragraph 2----
    text1(62) = "Much has been learned in the last decade about how tinnitus is generated by the brain.  "
    text1(63) = "Most cases are associated with some degree of high frequency hearing loss that often occurs"
    text1(64) = "with normal aging.   Hearing loss can be induced or accelerated by exposure to very loud "
    text1(65) = "sounds (for example, loud rock concerts, motorcycle engines, or fire crackers), so such "
    text1(66) = "sounds should be avoided because they can damage your ears.  The tinnitus sound however "
    text1(67) = "will not harm your ears.  Over time most individuals with tinnitus find that their tinnitus "
    text1(68) = "becomes less distrurbing and intrusive."
    text1(69) = "� L. Roberts and D. Bosnyak, 2011"
    '---paragraph 2----
    
    Set oExcel = New Excel.Application
    Set oWB = oExcel.Workbooks.Add
    Set oWS = oWB.Worksheets("Sheet1")
    Set oRng1 = oWS.Range("A1")
    Set oRng2 = oWS.Range("B2:E5")

    oExcel.visible = False ' <-- *** Optional *** 'true = show actions. False = don't show actions
    'Set up header info:
    '-----------------------------------------'
    oWS.Range("A1").Value = "Tinnitus Report"
    oWS.Range("A1:I1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection.Font
        .Name = "Calibri"
        .SIZE = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    '-----------------------------------------'
    
    'Next, set up username/info/Data:
    '-----------------------------------------'
    oWS.Range("A3").Value = "Name:"
    oWS.Range("B3").Value = UserName
    oWS.Range("A4").Value = "Date:"
    oWS.Range("B4").Value = Format(Now(), "MMM-dd-yy")
    oWS.Range("F3").Value = "City:"
    oWS.Range("G3").Value = UserCity
    oWS.Range("F4").Value = "Prov:"
    oWS.Range("G4").Value = UserProv
    oWS.Range("F5").Value = "Country:"
    oWS.Range("G5").Value = UserCountry
    
    oWS.Range("A7").Value = "Self-Reported Data:"
    oWS.Range("A8").Value = "Tinitus Location:"
    oWS.Range("C8").Value = UserTL
    oWS.Range("A9").Value = "Steady or Pulsing:"
    oWS.Range("C9").Value = UserSorP
    oWS.Range("A10").Value = "Bandwidth:"
    oWS.Range("C10").Value = UserBW
    
    oWS.Range("F8").Value = "Age:"
    oWS.Range("G8").Value = UserAge
    oWS.Range("F9").Value = "Sex:"
    oWS.Range("G9").Value = UserSex
    oWS.Range("F10").Value = "Onset:"
    oWS.Range("G10").Value = UserOnset
    '-----------------------------------------'
    '-----------------------------------------'
    'format user data:
    oWS.Range("G8:G10").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'insert text into body of excel file:
    oWS.Range("A12").Value = text1(1)
    oWS.Range("A13").Value = text1(2)
    oWS.Range("A14").Value = text1(3)
    oWS.Range("A15").Value = text1(4)
    
    oWS.Range("A17").Value = text1(5)
    oWS.Range("A18").Value = text1(6)
    oWS.Range("A19").Value = text1(7)
    
    oWS.Range("A21").Value = text1(8)
    oWS.Range("A22").Value = text1(9)
    oWS.Range("A23").Value = text1(10)
    oWS.Range("A24").Value = text1(11)
    oWS.Range("A25").Value = text1(12)
    oWS.Range("A26").Value = text1(13)
        
    oWS.Range("A28").Value = text1(14)
    oWS.Range("A29").Value = text1(15)
    oWS.Range("A30").Value = text1(16)
    oWS.Range("A31").Value = text1(17)
    
    oWS.Range("A33").Value = "(1) YOUR TINNITUS LOUDNESS"
    oWS.Range("A33").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    
    oWS.Range("A35").Value = text1(18)
    oWS.Range("A36").Value = text1(19)
    oWS.Range("A37").Value = text1(20)
    
    oWS.Range("A39").Value = text1(21)
    oWS.Range("A40").Value = text1(22)
    oWS.Range("A41").Value = text1(23)
    oWS.Range("A42").Value = text1(24)
    oWS.Range("A43").Value = text1(25)
    
    oWS.Range("A45").Value = text1(26)
    oWS.Range("A46").Value = text1(27)
    
    'plot Loudness Rating:
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.Location Where:=xlLocationAsObject, Name:="Sheet1"
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Tinnitus Loudness Rating"
    ActiveChart.SeriesCollection(1).Name = "=""Baseline Sample"""
    ActiveChart.SeriesCollection(1).XValues = "={5,15,25,35,45,55,65,75,85,95}"
    'ActiveChart.SeriesCollection(1).XValues = "={""Extremely Weak"","""","""",""Moderate"","""",""Strong"","""",""Very Strong"","""",""Extremely Strong""}"
    'ActiveChart.SeriesCollection(1).XValues = "= {""Extremely Weak"","""",""Moderate"","""",""Strong"","""",""Very Strong"","""","""",""Extremely Strong""}"
    ActiveChart.SeriesCollection(1).Values = "={0,0.02,0.12,0.28,0.19,0.17,0.04,0.11,0.04,0}"
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabelPosition = xlNone
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your Tinnitus Loudness"""
    'ActiveChart.SeriesCollection(2).FontSize = 9
    ActiveChart.SeriesCollection(2).Values = "={0.225}"
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    'ActiveChart.SeriesCollection(2).XValues = "=Sheet1!$C$53"
    ActiveChart.SeriesCollection(2).XValues = TL1
    ActiveChart.PlotArea.Select
    Selection.Height = 142

    'add textbox with labels                                        '(orient, left, top, width, height)
    ActiveChart.ChartArea.Select
    'insert access number labels
    i = 30
    w = 19
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "0"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w - 2
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "10"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "20"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "30"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w - 1
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "40"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "50"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w + 1
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "60"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w - 1
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "70"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "80"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 25, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "90"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
        i = i + w - 1
        Set ThisLine = ActiveChart.Shapes.AddTextbox(1, i, 175, 30, 15) 'msoTextOrientationHorizontal = 1,
        ThisLine.DrawingObject.Text = "100"
        ThisLine.Select
        With Selection
            .Font.SIZE = 9
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlTop
        End With
    
    'add "Extremely weak" text box
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 25, 185, 50, 25) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Extremely Weak"
    ThisLine.Select
    With Selection
        .Font.SIZE = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    'add "moderate" text box
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 72, 185, 50, 25) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Moderate"
    ThisLine.Select
    With Selection
        .Font.SIZE = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    'add "strong" text box
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 110, 185, 50, 25) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Strong"
    ThisLine.Select
    With Selection
        .Font.SIZE = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    'add "Verystrong" text box
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 147, 185, 50, 25) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Very Strong"
    ThisLine.Select
    With Selection
        .Font.SIZE = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    'add "Verystrong" text box
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 195, 185, 50, 25) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Extremely Strong"
    ThisLine.Select
    With Selection
        .Font.SIZE = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    With ActiveChart.Parent
        .Top = Range("B48").Top
        .Left = Range("B48").Left
    End With
    oWS.Range("E63").Value = "Figure 1"
    oWS.Range("E63").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10

'plot Loudness Match - New Plot - based on 1khz level match (created Jan 2011):
'enter text:
    oWS.Range("A66").Value = text1(28)
    oWS.Range("A67").Value = text1(29)
    oWS.Range("A68").Value = text1(30)
    oWS.Range("A69").Value = text1(31)
    oWS.Range("A70").Value = text1(32)
    'first, we have have to calculate where to put users' loudness in the plot:
    Select Case LM1
        Case Is < 10
            LMP = 1
        Case Is <= 20
            LMP = 2
        Case Is <= 30
            LMP = 3
        Case Is <= 40
            LMP = 4
        Case Is <= 50
            LMP = 5
        Case Is <= 60
            LMP = 6
        Case Is > 60
            LMP = 7
    End Select
    oWS.Range("A73:G73").Select
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$73:$G$73")
    ActiveChart.ChartType = xlColumnClustered
    'ActiveChart.Axes(xlValue).MajorGridlines.Select
    'ActiveSheet.ChartObjects("Chart 3").Activate
    
    ActiveChart.SeriesCollection(1).Values = "={0.152777778,0.236111111,0.236111111,0.180555556,0.111111111,0.041666667,0.041666667}"
    ActiveChart.SeriesCollection(1).Name = "=""Baseline Sample"""
    'ActiveChart.SeriesCollection(1).XValues = "={1,2,3,4,5,6,7,8,9,10}"
    ActiveChart.SeriesCollection(1).XValues = "={""<10"",""11-20"",""21-30"",""31-40"",""41-50"",""51-60"","">60""}"
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Tinnitus Loudness Match"
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your Loudness"""
    ActiveChart.SeriesCollection(2).Values = "={.17}"
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    ActiveChart.SeriesCollection(2).XValues = LMP
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -58.5
    With ActiveChart.Parent
        .Top = Range("B73").Top
        .Left = Range("B73").Left
    End With
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Loudness in dB SL"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    'insert pic
    'oWS.Range("B62").Select
    'ActiveSheet.Pictures.Insert("C:\TinData\axis1.JPG").Select
    'With Selection.ShapeRange
    '   .Top = Range("B60").Top + 2
    '   .Left = Range("B60").Left + 34
    '   .Height = Application.InchesToPoints(0.36)
    '   .Width = Application.InchesToPoints(2.97)
    'End With
    oWS.Range("E88").Value = "Figure 2"
    oWS.Range("E88").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10

    'SECTION 2
    oWS.Range("A94").Value = "(2)  YOUR TINNITUS SPECTRUM"
    oWS.Range("A94").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    oWS.Range("A96").Value = text1(33)
    oWS.Range("A97").Value = text1(34)
    oWS.Range("A98").Value = text1(35)
    oWS.Range("A99").Value = text1(36)
    oWS.Range("A100").Value = text1(37)
    
    oWS.Range("A102").Value = text1(38)
    oWS.Range("A103").Value = text1(39)
    
    'PLOT LIKENESS
    'first, calculate users avg likeness values:
    For x11 = 1 To 11 Step 1
        TnSp(x11) = (CInt(txtPitchMatchT1(x11 - 1).Text) + CInt(txtPitchMatchT2(x11 - 1).Text) + CInt(txtPitchMatchT3(x11 - 1).Text)) / 3
        If TnSp(x11) < 0 Then TnSp(x11) = 0 'just in case there are some -101 codes
    Next x11
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$104:$K$105")
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SeriesCollection(1).Values = "={15,21.8,31.9,42.5,45.7,56.5,58.4,54.5,54,54.3,29.8}"
    ActiveChart.SeriesCollection(1).XValues = "={.5,1,2,3,4,5,6,7,8,10,12}"
    ActiveChart.SeriesCollection(1).Name = "=""Baseline Spectrum"""
    ActiveChart.SeriesCollection(1).ChartType = xlXYScatterLines
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Tinnitus Likeness Match"
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Frequency (kHz)"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Tinnitus Likeness"
    End With
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your Spectrum"""
    'ActiveChart.SeriesCollection(2).Values = "={5,   10,  40,  48,  65,  85,  90,  95,  70,  42,  12}"
    ActiveChart.SeriesCollection(2).Values = TnSp
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatterLines
    ActiveChart.SeriesCollection(2).XValues = "={.5,1,2,3,4,5,6,7,8,10,12}"
    ActiveChart.Axes(xlCategory).MaximumScale = 13
    ActiveChart.Axes(xlCategory).MinimumScale = 0
    ActiveChart.Axes(xlCategory).MajorUnit = 1
    With ActiveChart.Parent
        .Top = Range("B105").Top
        .Left = Range("B105").Left
    End With
    oWS.Range("E120").Value = "Figure 3"
    oWS.Range("E120").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10
    'PLOT LIKENESS
    
    
    'RI data
    oWS.Range("A121").Value = "(3) RESIDUAL INHIBITION"
    oWS.Range("A121").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    
    'INSERT TEXT HERE
    oWS.Range("A123").Value = text1(40)
    oWS.Range("A124").Value = text1(41)
    oWS.Range("A125").Value = text1(42)
    oWS.Range("A126").Value = text1(43)
    
    oWS.Range("A128").Value = text1(44)
    oWS.Range("A129").Value = text1(45)
    oWS.Range("A130").Value = text1(46)
    oWS.Range("A131").Value = text1(47)
    
    oWS.Range("A141").Value = text1(48)
    oWS.Range("A142").Value = text1(49)
    oWS.Range("A143").Value = text1(50)
    oWS.Range("A144").Value = text1(51)
    oWS.Range("A145").Value = text1(52)
    oWS.Range("A146").Value = text1(53)
    oWS.Range("A147").Value = text1(54)
    oWS.Range("A148").Value = text1(55)
    oWS.Range("A149").Value = text1(56)
    
    oWS.Range("A151").Value = text1(57)
    oWS.Range("A152").Value = text1(58)
    

    
    'Plot RI Data
    ActiveSheet.Shapes.AddChart.Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$A$154:$K$154")
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.SeriesCollection(1).Values = "={0, 0.021276596, 0,0.063829787,0.085106383,0.085106383,0.276595745,0.14893617,0.063829787,0.170212766,0.085106383}"
    'ActiveChart.SeriesCollection(1).XValues = "={1,2,3,4,5,6,7,8,9,10,11}"
    ActiveChart.SeriesCollection(1).XValues = "={""5"",""4"",""3"",""2"",""1"",""0"",""-1"",""-2"",""-3"",""-4"",""-5""}"
    ActiveChart.SeriesCollection(1).Name = "=""Baseline RI"""
    ActiveChart.SeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "Residual Inhibition"
    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Residual Inhibition"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.SeriesCollection(2).Name = "=""Your RI"""
    ActiveChart.SeriesCollection(2).Values = "{.215}"
    ActiveChart.SeriesCollection(2).ChartType = xlXYScatter
    ActiveChart.SeriesCollection(2).XValues = (6 - RI5) 'user input
    'ActiveSheet.Shapes("Chart 1").IncrementLeft -58.5
    With ActiveChart.Parent
        .Top = Range("B154").Top
        .Left = Range("B154").Left
    End With
    Set ThisLine = ActiveChart.Shapes.AddLine(175, 25, 175, 38) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 2
    End With
    
    Set ThisLine = ActiveChart.Shapes.AddLine(95, 30, 60, 30) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BeginArrowheadStyle = 1 'none
        .EndArrowheadStyle = 2 'triangle

        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 1
    End With
    Set ThisLine = ActiveChart.Shapes.AddLine(255, 30, 290, 30) 'expression.AddLine(BeginX, BeginY, EndX, EndY)
    With ThisLine.Line
        .BeginArrowheadStyle = 1
        .EndArrowheadStyle = 2 'msoArrowheadTriangle

        .BackColor.RGB = 0
        .ForeColor.RGB = 0
        .Weight = 1
    End With

    With ActiveChart
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "Change in Tinnitus"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Percentage of Patients"
    End With
    'add textbox with caption "Louder"
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 92, 20, 100, 10) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Tinnitus Louder"
    Set ThisLine = ActiveChart.Shapes.AddTextbox(1, 178, 20, 100, 10) 'msoTextOrientationHorizontal = 1
    ThisLine.DrawingObject.Text = "Tinnitus Softer"
    
    oWS.Range("E169").Value = "Figure 4"
    oWS.Range("E169").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 10
    
'Final Comment
    oWS.Range("A171").Value = "(4) A FINAL COMMENT"
    oWS.Range("A171").Select
    Selection.Font.Bold = True
    Selection.Font.SIZE = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    
    'INSERT TEXT HERE
    oWS.Range("A174").Value = text1(59)
    oWS.Range("A175").Value = text1(60)
    oWS.Range("A176").Value = text1(61)
    
    oWS.Range("A178").Value = text1(62)
    oWS.Range("A179").Value = text1(63)
    oWS.Range("A180").Value = text1(64)
    oWS.Range("A181").Value = text1(65)
    oWS.Range("A182").Value = text1(66)
    oWS.Range("A183").Value = text1(67)
    oWS.Range("A184").Value = text1(68)
    
    oWS.Range("A185") = text1(69)
    oWS.Range("A185").Select
    Selection.Font.Italic = True
    Selection.Font.SIZE = 8
    
Cleanup:
    ActiveWorkbook.SaveAs FileName:=(WorkingDir & "\TinnitusReport.xlsx"), FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Set oWS = Nothing
    If Not oWB Is Nothing Then oWB.Close
    Set oWB = Nothing
    
    oExcel.Quit
    Set oExcel = Nothing
End Sub

Private Sub CheckLicense()
End Sub

' Clears old data from the form controls.
Private Sub clearOldData()
    Dim z As Integer
    
    ' Clear text boxes
    txtLocalize.Text = ""
    txtIntensity.Text = ""
    txtIntensity2.Text = ""
    txtBandwidth.Text = ""
    txtTemporal.Text = ""
    txtLoudness.Text = ""
    
    ' Clear loudness text boxes
    For z = 0 To 10 Step 1
        txtLoudnessT1(z).Text = ""
        txtLoudnessT2(z).Text = ""
    Next z
    
    ' Clear pitch match text boxes
    For z = 0 To 10 Step 1
        txtPitchMatchT1(z).Text = ""
        txtPitchMatchT2(z).Text = ""
        txtPitchMatchT3(z).Text = ""
    Next z
    
    ' Clear sound threshold text boxes
    For z = 0 To 3 Step 1
        txtSoundThreshold(z).Text = ""
    Next z
    
    ' Clear PA5 threshold value text box
    txtPA5ThreshValue.Text = ""
    
    ' Clear sound level match text boxes
    For z = 0 To 3 Step 1
        txtSoundLevelMatch(z).Text = ""
    Next z
End Sub
