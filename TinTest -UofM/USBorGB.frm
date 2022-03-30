VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "USB or GB"
   ClientHeight    =   10170
   ClientLeft      =   8835
   ClientTop       =   3540
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   4680
   Begin VB.Frame framReport 
      Caption         =   "Tinnitus Report"
      Height          =   1815
      Left            =   360
      TabIndex        =   14
      Top             =   7440
      Width           =   3975
      Begin VB.OptionButton optReportNo 
         Caption         =   "No"
         Height          =   255
         Left            =   960
         TabIndex        =   16
         Top             =   1440
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optReportYes 
         Caption         =   "Yes"
         Height          =   195
         Left            =   960
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblReport 
         Caption         =   "Output Tinnitus Report?       (Microsoft Excel 2007 or later must be installed!)"
         Height          =   615
         Left            =   480
         TabIndex        =   17
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame framLang 
      Caption         =   "Language"
      Height          =   1815
      Left            =   360
      TabIndex        =   11
      Top             =   5400
      Width           =   3975
      Begin VB.OptionButton optFrench 
         Caption         =   "French"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optEnglish 
         Caption         =   "English"
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame framPA5s 
      Caption         =   "PA5s"
      Height          =   1695
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   3975
      Begin VB.OptionButton opt2PA5 
         Caption         =   "Use 2 PA5s"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   960
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton opt1PA5 
         Caption         =   "Use 1 PA5"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame framvRes 
      Caption         =   "Vertical Resolution"
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   3975
      Begin VB.OptionButton opt1200 
         Caption         =   "1200"
         Height          =   255
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.OptionButton opt1024 
         Caption         =   "1024"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   9600
      Width           =   1455
   End
   Begin VB.Frame framUSBGB 
      Caption         =   "USB or GB"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.OptionButton optGB 
         Caption         =   "Gigabit"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton optUSB 
         Caption         =   "USB"
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Please choose your TDT PA5 Interface:"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim intfilenumber As Integer
If (optUSB = False And optGB = False) Or (opt1024 = False And opt1200 = False) Then
    MsgBox "Please choose an interface and a screen resolution"

Else
    If optUSB = True Then 'connect using the usb protocol
        If Form1.PA5x1.ConnectPA5("USB", 1) >= 1 Then
        '    Debug.Print ("connected to pa5 #1")
        Else
            Debug.Print ("unable to connect to pa5 #1")
            MsgBox ("Unable to connect to pa5 #1")
            End
        End If
        If opt2PA5 Then 'connect to the 2nd pa5
            usePA52 = True
            Debug.Print ("Connecting to PA5 #2 ....")
            If Form1.PA5x2.ConnectPA5("USB", 2) >= 1 Then
            '    Debug.Print ("connected to pa5 #2")
            Else
                Debug.Print ("unable to connect to pa5 #2")
                MsgBox ("Unable to connect to pa5 #2")
                End
            End If
        Else
            usePA52 = False
        End If
    Else 'connect using GB
        If Form1.PA5x1.ConnectPA5("GB", 1) >= 1 Then
        '    Debug.Print ("connected to pa5 #1")
        Else
            Debug.Print ("unable to connect to pa5 #1")
            MsgBox ("Unable to connect to pa5 #1")
            End
        End If
        If opt2PA5 Then 'connect to the 2nd pa5
            usePA52 = True
            Debug.Print ("Connecting to PA5 #2 ....")
            If Form1.PA5x2.ConnectPA5("GB", 2) >= 1 Then
            '    Debug.Print ("connected to pa5 #2")
            Else
                Debug.Print ("unable to connect to pa5 #2")
                MsgBox ("Unable to connect to pa5 #2")
                End
            End If
        Else
            usePA52 = False
        End If
    End If
    If opt1024.Value = True Then
        vRes = 1024
        Form1.SoundTypeDialTop = 80
        Form1.SoundBandwidthDialTop = 80
        Form1.WhichEarTop = 78
        Form1.YesNoTop = 79
        Form1.cmdTinTrain.Top = 700
        Form1.cmdNext.Top = 700
    Else
        vRes = 1200
        Form1.SoundTypeDialTop = 80
        Form1.SoundBandwidthDialTop = 80
        Form1.WhichEarTop = 78
        Form1.YesNoTop = 79
        Form1.cmdTinTrain.Top = 968
        Form1.cmdNext.Top = 968
    End If
    Form1.PA5x1.SetAtten (90)
    If optEnglish Then
        English = True 'set language as english
        Form1.lblMainInstructions.Font = "Arial"
        Form1.lblMainInstructions.FontSize = 32
        Form1.lblInstruct2.Font = "Arial"
        Form1.lblInstruct2.FontSize = 32
        Form1.lblInstruct3.Font = "Arial"
        Form1.lblInstruct3.FontSize = 32
    Else
        English = False 'set language as french
        Form1.lblMainInstructions.Font = "Arial Narrow"
        Form1.lblMainInstructions.FontSize = 32
        Form1.lblInstruct2.Font = "Arial Narrow"
        Form1.lblInstruct2.FontSize = 32
        Form1.lblInstruct3.Font = "Arial Narrow"
        Form1.lblInstruct3.FontSize = 32
    End If
    If opt2PA5 Then 'user is using 2 pa5s so set level for 2nd pa5
        Form1.PA5x2.SetAtten (90)
    End If
    If optReportNo Then 'no tinnitus report so go right to program
        PReport = False
        Form2.Hide
        Form1.Show
        Form1.SetFocus
    Else 'tinitus report, so ask questions
        PReport = True
        Form2.Hide
        formUserID.Show
    End If
    'option selected
End If
    'save setup for future sessions
    intfilenumber = FreeFile
    Open ("C:\TinData\init.ini") For Output As #intfilenumber
        If optUSB = True Then Write #intfilenumber, 0 Else Write #intfilenumber, 1
        If opt1024 = True Then Write #intfilenumber, 0 Else Write #intfilenumber, 1
        If opt1PA5 = True Then Write #intfilenumber, 0 Else Write #intfilenumber, 1
        If optEnglish = True Then Write #intfilenumber, 0 Else Write #intfilenumber, 1
        If optReportNo = True Then Write #intfilenumber, 0 Else Write #intfilenumber, 1
    Close #intfilenumber
End Sub



Private Sub Form_Load()
Dim intfilenumber As Integer
Dim tempS As Single
Dim c1 As Integer
'MsgBox (Dir("C:\TinData\CalibrationData.csv"))
If (dir("C:\TinData\init.ini")) = "init.ini" Then 'the datafile exists
    intfilenumber = FreeFile
    Open ("C:\TinData\init.ini") For Input As #intfilenumber
        Input #intfilenumber, tempS
        If tempS = 0 Then
            optUSB = True
            optGB = False
        Else
            optUSB = False
            optGB = True
        End If
        Input #intfilenumber, tempS
        If tempS = 0 Then
            opt1024 = True
            opt1200 = False
        Else
            opt1024 = False
            opt1200 = True
        End If
        Input #intfilenumber, tempS
        If tempS = 0 Then
            opt1PA5 = True
            opt2PA5 = False
        Else
            opt1PA5 = False
            opt2PA5 = True
        End If
        Input #intfilenumber, tempS
        If tempS = 0 Then
            optEnglish = True
            optFrench = False
        Else
            optEnglish = False
            optFrench = True
        End If
        Input #intfilenumber, tempS
        If tempS = 0 Then
            optReportYes = False
            optReportNo = True
        Else
            optReportYes = True
            optReportNo = False
        End If
    Close #intfilenumber
Else 'data file does not exist; set all values to standard
            opt1PA5 = False
            opt2PA5 = True
            opt1024 = False
            opt1200 = True
            optUSB = False
            optGB = True
            optEnglish = True
            optFrench = False
            optReportYes = False
            optReportNo = True
End If
End Sub
