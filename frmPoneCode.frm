VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPoneCode 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   10320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10320
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "O&k"
      Height          =   375
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9700
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Phone Code Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9250
      Begin MSComctlLib.ListView lstRace 
         Height          =   9000
         Left            =   240
         TabIndex        =   17
         Top             =   300
         Visible         =   0   'False
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   15875
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Country"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ethnicity and Race"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.ListView lstPhoneCode 
         Height          =   4335
         Left            =   240
         TabIndex        =   4
         Top             =   4890
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "City Name"
            Object.Width           =   4623
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Phone Code"
            Object.Width           =   2788
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "City Name"
            Object.Width           =   4623
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Phone Code"
            Object.Width           =   2788
         EndProperty
      End
      Begin VB.Label lblElec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   7160
         TabIndex        =   15
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label lblphone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   7160
         TabIndex        =   14
         Top             =   3745
         Width           =   45
      End
      Begin VB.Image imgPHStat 
         Height          =   735
         Index           =   3
         Left            =   7160
         Stretch         =   -1  'True
         Top             =   2920
         Width           =   975
      End
      Begin VB.Image imgElStat 
         Height          =   735
         Index           =   3
         Left            =   7160
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblphone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   13
         Top             =   3745
         Width           =   45
      End
      Begin VB.Label lblphone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2440
         TabIndex        =   12
         Top             =   3745
         Width           =   45
      End
      Begin VB.Label lblElec 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   4800
         TabIndex        =   11
         Top             =   2250
         Width           =   45
      End
      Begin VB.Image imgElStat 
         Height          =   735
         Index           =   2
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgPHStat 
         Height          =   735
         Index           =   2
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   2920
         Width           =   975
      End
      Begin VB.Image imgPHStat 
         Height          =   735
         Index           =   1
         Left            =   2440
         Stretch         =   -1  'True
         Top             =   2920
         Width           =   975
      End
      Begin VB.Label lblphone 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   3745
         Width           =   45
      End
      Begin VB.Label lblElec 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   2440
         TabIndex        =   9
         Top             =   2250
         Width           =   45
      End
      Begin VB.Label lblElec 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2250
         Width           =   45
      End
      Begin VB.Image imgPHStat 
         Height          =   735
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   2920
         Width           =   975
      End
      Begin VB.Image imgElStat 
         Height          =   735
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   975
      End
      Begin VB.Image imgElStat 
         Height          =   735
         Index           =   1
         Left            =   2440
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblNoCity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "lblNoCity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   3810
         TabIndex        =   6
         Top             =   4400
         Width           =   1275
      End
      Begin VB.Label lblIDDInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "lblIDDInfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1680
         Left            =   240
         TabIndex        =   5
         Top             =   2560
         Width           =   8415
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "lblInfo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1500
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   8415
      End
      Begin VB.Image ImgCntFlag 
         Height          =   495
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblcontryName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Toronto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   420
         Width           =   7215
      End
   End
   Begin VB.Label lblRanking 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6255
      TabIndex        =   16
      Top             =   9795
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblCityCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   9795
      Width           =   75
   End
End
Attribute VB_Name = "frmPoneCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  frmPoneCode.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmPoneCode.Width / 2)
  frmPoneCode.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmPoneCode.Height / 2)
  Set cmdOk.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmWeatherMain.Timer1.Enabled = True
  Set frmPoneCode = Nothing
End Sub

