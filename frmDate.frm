VERSION 5.00
Begin VB.Form frmDate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Holiday Date Picker"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   2355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   700
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1200
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Select Date"
      ForeColor       =   &H8000000D&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin KDweather.Duncan_DatePicker Duncan_DatePicker1 
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         FirstDayOfWeek  =   1
         ShortDayNames   =   -1  'True
         DescriptionFormat=   "d mmm yyyy"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Duncan_DatePicker1_DateChanged(ByVal FromDate As Date, ByVal ToDate As Date)
  HolDateSelect = ToDate
  cmdOk_Click
End Sub

Private Sub Form_Load()
  Set cmdOk.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  frmDate.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmDate.Height / 2)
  frmDate.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmDate.Width / 2)
  HolDateSelect = Duncan_DatePicker1.DateSelected
End Sub
