VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Copy Class Example"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      Height          =   1305
      Left            =   225
      TabIndex        =   11
      Top             =   3420
      Width           =   5385
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   60
         Top             =   810
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   2535
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StatusLabel"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   18
         Top             =   735
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seconds Taken:"
         Height          =   195
         Index           =   7
         Left            =   315
         TabIndex        =   17
         Top             =   735
         Width           =   1185
      End
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StatusLabel"
         Height          =   195
         Index           =   1
         Left            =   1590
         TabIndex        =   16
         Top             =   495
         Width           =   840
      End
      Begin VB.Label StatusLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "StatusLabel"
         Height          =   195
         Index           =   0
         Left            =   1590
         TabIndex        =   15
         Top             =   255
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage ready:"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   13
         Top             =   465
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         Height          =   195
         Index           =   3
         Left            =   1110
         TabIndex        =   12
         Top             =   255
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Job"
      Height          =   2610
      Left            =   195
      TabIndex        =   1
      Top             =   675
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "&Copy NOW"
         Height          =   300
         Left            =   4005
         TabIndex        =   10
         Top             =   2190
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   1170
         TabIndex        =   4
         Text            =   "c:\temp\matrix.qt"
         Top             =   330
         Width           =   3330
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1185
         TabIndex        =   3
         Text            =   "c:\temp\matrix.qtb"
         Top             =   765
         Width           =   3345
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   975
         TabIndex        =   2
         Top             =   1275
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   32
         SmallChange     =   16
         Max             =   256
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Source File"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   8
         Top             =   795
         Width           =   795
      End
      Begin VB.Label CacheSizeLabel 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cache Size: 0K"
         Height          =   195
         Left            =   2160
         TabIndex        =   7
         Top             =   1740
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0K"
         Height          =   195
         Index           =   4
         Left            =   1095
         TabIndex        =   6
         Top             =   1710
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "256K"
         Height          =   195
         Index           =   5
         Left            =   4080
         TabIndex        =   5
         Top             =   1755
         Width           =   405
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   630
      Index           =   0
      Left            =   330
      TabIndex        =   0
      Top             =   45
      Width           =   5025
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Xerox As FileCopyClass

Private Sub Command1_Click()

' hope this does not confuse u all :)

Command1.Enabled = False

ProgressBar1.Value = 0
ProgressBar1.Visible = True

StatusLabel(0).Caption = "Copying"

Set Xerox = New FileCopyClass

Xerox.CacheSize = Slider1.Value * 1000

Timer1.Enabled = True
Xerox.Copy Text1.Text, Text2.Text

StatusLabel(1).Caption = "100%"
ProgressBar1.Value = 100

StatusLabel(2).Caption = Xerox.SecondsTaken

Timer1.Enabled = False
Set Xerox = Nothing

Command1.Enabled = True
StatusLabel(0).Caption = "Ready"
End Sub

Private Sub Form_Load()

Timer1.Enabled = False  'had to be so, u'll see

StatusLabel(0).Caption = "Ready"
StatusLabel(1).Caption = ""
StatusLabel(2).Caption = ""

End Sub

Private Sub Slider1_Click()
CacheSizeLabel.Caption = "Cache Size: " & Slider1.Value & "K"
End Sub

Private Sub Timer1_Timer()
StatusLabel(1).Caption = Xerox.PercentREady & "%"
ProgressBar1.Value = Xerox.PercentREady
End Sub
