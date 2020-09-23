VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAnimate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5800
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   8300
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   9180
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   5636
            MinWidth        =   5644
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4048
            MinWidth        =   4049
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4128
            MinWidth        =   4129
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4128
            MinWidth        =   4129
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   4657
            MinWidth        =   4657
         EndProperty
      EndProperty
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
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12615
      ExtentX         =   22251
      ExtentY         =   15266
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
   Animation = False
   PlayAnimation = False
   bGPS = False
   frmWeatherMain.Timer1.Enabled = True
   Unload Me
End Sub

Private Sub Form_Load()
   On Error GoTo errohandler
   frmWeatherMain.Timer1.Enabled = False
   Set cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
   If bGPS Then
      StatusBar1.Visible = True
      StatusBar1.Panels(1).Text = sStatArea
      StatusBar1.Panels(2).Text = sStatRegion
      StatusBar1.Panels(3).Text = sStatCountry
      StatusBar1.Panels(4).Text = sStatState
      StatusBar1.Panels(5).Text = sStatCounty
      frmAnimate.Width = 12975
      frmAnimate.Height = 9660
      WebBrowser1.Height = 8655
      WebBrowser1.Width = 12615
   ElseIf AQICanShowTool Or bMapView Or AQIMonitorShowTool Then
      StatusBar1.Style = sbrSimple
      StatusBar1.SimpleText = sStatusText
      frmAnimate.Width = 8850
      frmAnimate.Height = 8100
      WebBrowser1.Height = 6550
      WebBrowser1.Width = 8510
      cmdExit.Top = frmAnimate.Height - 1270
   Else
      StatusBar1.Style = sbrSimple
      StatusBar1.SimpleText = sStatusText
      frmAnimate.Width = 12500
      frmAnimate.Height = 9660
      WebBrowser1.Height = 7980
      WebBrowser1.Width = 12145
   End If
   cmdExit.Left = frmAnimate.Left + (frmAnimate.Width / 2 - cmdExit.Width / 2)
   WebBrowser1.Navigate AnimationLink
   Exit Sub
errohandler:
   MsgBox "Unable To Display GPS Location", vbInformation, "Weather Of The World"
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  sStatusText = ""
  Set frmAnimate = Nothing
End Sub
