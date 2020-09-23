VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCountry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "Current PM 2.5"
      Height          =   375
      Index           =   5
      Left            =   2840
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   5
      Left            =   6480
      Picture         =   "frmCountry.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   15
      Top             =   4600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   4
      Left            =   5400
      Picture         =   "frmCountry.frx":01E3
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   14
      Top             =   4600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   3
      Left            =   4320
      Picture         =   "frmCountry.frx":0418
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   13
      Top             =   4600
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   2
      Left            =   6480
      Picture         =   "frmCountry.frx":060B
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   4000
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   1
      Left            =   5400
      Picture         =   "frmCountry.frx":078D
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   11
      Top             =   4000
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   0
      Left            =   4320
      Picture         =   "frmCountry.frx":0959
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   10
      Top             =   4000
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "AQI Animation"
      Height          =   375
      Index           =   2
      Left            =   2840
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "Current PM 2.5"
      Height          =   375
      Index           =   4
      Left            =   2840
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "Current Ozone"
      Height          =   375
      Index           =   3
      Left            =   1480
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "Current AQI"
      Height          =   375
      Index           =   1
      Left            =   1480
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdAQITabs 
      Caption         =   "Forecast"
      Height          =   375
      Index           =   0
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtshow 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   3600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000002&
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
      Begin RichTextLib.RichTextBox rchTxtInfo 
         Height          =   1455
         Left            =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2566
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCountry.frx":0ADB
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
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Play Animation"
      Height          =   375
      Left            =   1320
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1680
   End
   Begin VB.Image picSource 
      Height          =   1095
      Left            =   960
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgLgCountry 
      Height          =   1080
      Left            =   45
      MousePointer    =   99  'Custom
      ToolTipText     =   "Right Click To Close"
      Top             =   120
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmCountry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bOldIndex As Integer
Dim zooming As Boolean
Dim MinLeft As Integer
Dim MinTop As Integer
Dim MaxLeft As Integer
Dim MaxTop As Integer
Dim AQITooltip(5) As New Tooltip

Private Sub cmdAnimate_Click()
  If cmdAnimate.Caption = "Exit" Then
    Unload Me
  Else
    Animation = True
    Timer1.Enabled = True
  End If
End Sub

Private Sub cmdAQITabs_Click(Index As Integer)
  If bMapView Or AQICanShowTool Then
    If Index <> 2 Then
      If cmdAQITabs(bOldIndex).Enabled = False Then
        cmdAQITabs(bOldIndex).Enabled = True
        cmdAQITabs(Index).Enabled = False
        bOldIndex = Index
      End If
    End If
  Else
    If Index <> 2 And Index <> 3 Then
      If cmdAQITabs(bOldIndex).Enabled = False Then
        cmdAQITabs(bOldIndex).Enabled = True
        cmdAQITabs(Index).Enabled = False
        bOldIndex = Index
      End If
    End If
  End If
  
  If bMapView Then
    picTureName = AQIPicArray(Index)
  Else
    picTureName = AQICanPicArray(Index)
  End If
  If bMapView Then
    If Index <> 2 Then
      SizePic picTureName
    Else
      AnimationLink = AQIPicArray(Index)
      frmAnimate.Caption = sFrmName
      frmAnimate.StatusBar1.SimpleText = sFrmName & " Index Animation Map"
      frmAnimate.Show vbModal
      cmdAQITabs(Index).Enabled = True
    End If
  ElseIf AQICanShowTool Then
    If Index = 0 Then
      SizePic picTureName
    Else
      AnimationLink = AQICanPicArray(Index)
      frmAnimate.Caption = sFrmName
      frmAnimate.StatusBar1.SimpleText = sFrmName & " Air Quality Index Animation Map"
      frmAnimate.Show vbModal
      cmdAQITabs(Index).Enabled = True
    End If
  Else
    If Index <> 2 And Index <> 3 Then
      SizePic picTureName
    Else
      AnimationLink = AQICanPicArray(Index)
      frmAnimate.Caption = sFrmName
      frmAnimate.StatusBar1.SimpleText = "USA " & cmdAQITabs(Index).Caption & " Air Quality Monitor Map"
      frmAnimate.Show vbModal
      cmdAQITabs(Index).Enabled = True
    End If
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set cmdAnimate.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  frmCountry.Caption = sFrmName
  If AQIShowTool Or AQICanShowTool Or AQIMonitorShowTool Or bNoAQIndex Then
    Dim cnt As Integer
    LoadAQITips
    For cnt = 0 To 5
      AQITooltip(cnt).Style = TTBalloon
      AQITooltip(cnt).BackColor = AQIbgCol(cnt)
      AQITooltip(cnt).Title = AQITitle(cnt)
      AQITooltip(cnt).TipText = AQITxt(cnt)
      If cnt > 2 Then
        AQITooltip(cnt).ForeColor = vbWhite
      End If
      Set AQITooltip(cnt).ParentControl = AQIPic(cnt)
      AQITooltip(cnt).Create
    Next
  End If
  bOldIndex = 0
  cmdAQITabs(0).Enabled = False
  SizePic picTureName
  If Len(sStatusText) <> 0 Then
    fraInfo.Visible = True
    fraInfo.Caption = sFrmName
    rchTxtInfo.TextRTF = sStatusText
    rchTxtInfo.Visible = True
    txtshow.Top = frmCountry.Height + 200
    txtshow.Visible = True
    txtshow.Text = "1"
  Else
    fraInfo.Visible = False
  End If
  If bPicError Then
    Unload frmCountry
  Else
    frmCountry.Show vbModal
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Erase AQITooltip
  Erase AQIbgCol
  sStatusText = ""
  sFrmName = ""
  PlayAnimation = False
  Set frmCountry = Nothing
End Sub

Private Sub imgLgCountry_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Timer1.Enabled = True
  End If
End Sub

Public Sub ZoomPicture(BoxWidth As Integer, BoxHeight As Integer)
   Dim X, Y As Single
   On Error Resume Next
  
   X = BoxWidth
   Y = BoxHeight
   X = X / 1.0851
   Y = Y / 1.0851
   BoxWidth = X
   BoxHeight = Y
    
   Call ShowZoom(BoxWidth, BoxHeight, picTureName)
   'Center picture
   frmCountry.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmCountry.Width / 2)
   frmCountry.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmCountry.Height / 2)
   imgLgCountry.Visible = True
   If Y < 425 Or X < 200 Then
      Timer1.Enabled = False
      Unload Me
   End If
End Sub

Private Sub Timer1_Timer()
   Call ZoomPicture(imgLgCountry.Width, imgLgCountry.Height)
End Sub

Private Sub ShowZoom(BoxWidth As Integer, BoxHeight As Integer, picName As String)
   imgLgCountry.Height = BoxHeight
   imgLgCountry.Width = BoxWidth
   frmCountry.Height = imgLgCountry.Height + 300
   frmCountry.Width = imgLgCountry.Width + 25
End Sub
