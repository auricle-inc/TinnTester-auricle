VERSION 5.00
Begin VB.Form formUserID 
   BackColor       =   &H00F0F0E6&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Information"
   ClientHeight    =   4335
   ClientLeft      =   9690
   ClientTop       =   6120
   ClientWidth     =   7080
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7080
   Begin VB.CommandButton cmdUserInfoOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Frame framUserData 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Information"
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox lblOnset 
         Height          =   285
         Left            =   4920
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox lblSex 
         Height          =   285
         Left            =   4920
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox lblAge 
         Height          =   285
         Left            =   4920
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox lblCountry 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox lblProvince 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox lblCity 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox lblName 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Date of Tinnitus Onset:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   6
         Left            =   2520
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Sex:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   7
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   4080
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Prov/State:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00F0F0E6&
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label lblUserDataTitle 
      BackColor       =   &H00F0F0E6&
      Caption         =   "Subject Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "formUserID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdUserInfoOK_Click()
If lblName.Text = "" Then
    UserName = "Not Entered"
Else
    UserName = lblName.Text
End If

If lblCity.Text = "" Then
    UserCity = "Not Entered"
Else
    UserCity = lblCity.Text
End If

If lblProvince.Text = "" Then
    UserProv = "Not Entered"
Else
    UserProv = lblProvince.Text
End If

If lblCountry.Text = "" Then
    UserCountry = "Not Entered"
Else
    UserCountry = lblCountry.Text
End If

If lblAge.Text = "" Then
    UserAge = "NA"
Else
    UserAge = lblAge.Text
End If

If lblSex.Text = "" Then
    UserSex = "NA"
Else
    UserSex = lblSex.Text
End If

If lblOnset.Text = "" Then
    UserOnset = "NA"
Else
    UserOnset = lblOnset.Text
End If

formUserID.Hide
Form1.Show
Form1.SetFocus
End Sub

