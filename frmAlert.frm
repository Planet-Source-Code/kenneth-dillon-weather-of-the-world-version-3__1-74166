VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmAlert 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Severe Weather Alert"
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
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Find City"
      Height          =   375
      Left            =   2760
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   9700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox AQIPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Index           =   5
      Left            =   8440
      Picture         =   "frmAlert.frx":0000
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   23
      Top             =   9650
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
      Left            =   7420
      Picture         =   "frmAlert.frx":01E3
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   22
      Top             =   9650
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
      Left            =   6400
      Picture         =   "frmAlert.frx":0418
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   21
      Top             =   9650
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
      Left            =   5400
      Picture         =   "frmAlert.frx":060B
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   20
      Top             =   9650
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
      Left            =   4400
      Picture         =   "frmAlert.frx":078D
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   19
      Top             =   9650
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
      Left            =   3360
      Picture         =   "frmAlert.frx":0959
      ScaleHeight     =   525
      ScaleWidth      =   900
      TabIndex        =   18
      Top             =   9650
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CommandButton cmdMapView 
      Caption         =   "View Map"
      Height          =   375
      Left            =   1560
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   9700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1615
      Left            =   240
      TabIndex        =   15
      Top             =   10720
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2858
      _Version        =   393217
      TextRTF         =   $"frmAlert.frx":0ADB
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1440
      Top             =   11200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9250
      Begin RichTextLib.RichTextBox rchTxtAnthem 
         Height          =   9070
         Left            =   200
         TabIndex        =   16
         Top             =   200
         Visible         =   0   'False
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   16007
         _Version        =   393217
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmAlert.frx":0B66
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5000
         ScaleHeight     =   300
         ScaleWidth      =   3000
         TabIndex        =   12
         Top             =   8950
         Visible         =   0   'False
         Width           =   3000
         Begin VB.Label lblHur2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1000
         ScaleHeight     =   300
         ScaleWidth      =   3000
         TabIndex        =   11
         Top             =   8950
         Visible         =   0   'False
         Width           =   3000
         Begin VB.Label lblHur1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   2895
         End
      End
      Begin VB.PictureBox picHur2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2000
         Left            =   5000
         MousePointer    =   99  'Custom
         ScaleHeight     =   1995
         ScaleWidth      =   3000
         TabIndex        =   10
         ToolTipText     =   " Click To Enlarge "
         Top             =   6900
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.PictureBox picHur1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2000
         Left            =   1000
         MousePointer    =   99  'Custom
         ScaleHeight     =   1995
         ScaleWidth      =   3000
         TabIndex        =   9
         ToolTipText     =   " Click To Enlarge "
         Top             =   6900
         Visible         =   0   'False
         Width           =   3000
      End
      Begin MSComctlLib.ListView lsvStormName 
         Height          =   3030
         Left            =   195
         TabIndex        =   6
         Top             =   5160
         Visible         =   0   'False
         Width           =   8610
         _ExtentX        =   15187
         _ExtentY        =   5345
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   3704
         EndProperty
      End
      Begin MSComctlLib.ListView lstWeatherAlert 
         Height          =   8950
         Left            =   195
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   15796
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "123"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "123"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "123"
            Object.Width           =   4939
         EndProperty
      End
      Begin MSComctlLib.ListView lstPopulation 
         Height          =   9050
         Left            =   200
         TabIndex        =   8
         Top             =   200
         Visible         =   0   'False
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   15954
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
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
            Text            =   "Country"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Capital"
            Object.Width           =   3353
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Area Sq/Mi  "
            Object.Width           =   1941
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Population      "
            Object.Width           =   2470
         EndProperty
      End
      Begin VB.TextBox txtAlert 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9000
         Left            =   200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   200
         Visible         =   0   'False
         Width           =   8855
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9700
      Width           =   1215
   End
   Begin VB.ComboBox cmbcntyName 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   10680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7320
      Top             =   10560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   11
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   225
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":0BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":0D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":0F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":15C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1754
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1EED
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2072
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":21FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2383
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":250C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":26A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":29B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2B45
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3037
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":34CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":39D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4427
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":48DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":584D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":671C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":7945
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":8303
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":8C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":8FA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":912E
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":980F
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":A405
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":A7A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":AC31
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":B55F
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":BF22
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":CA3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":CFF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":D187
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":D869
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":DF6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":E70F
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":EBC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":F0B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":F69C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":FA3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":10271
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":10851
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":10EA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1135C
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":11DC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":123CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":12A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":12F87
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":13388
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":13946
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":143B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":14C30
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":15561
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":16114
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":167BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1703B
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1784E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":18218
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1853C
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":18DB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":19460
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1A0E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1A50D
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1ADFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1B6F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1BBA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1C054
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1C3F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1C809
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1CB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1D02F
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1D3D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1D8BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1E26B
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1EECC
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1F46A
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":1F91F
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2018E
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":20AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":21035
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":213D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":218F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":221D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":224E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":23432
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":23B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":24051
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":241DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":24690
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":24819
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":249A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":24F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":25E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":267BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":27415
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":27884
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2853E
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":28A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":28D37
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2946D
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":29C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2A17E
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2A3CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2AB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2AECD
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2B26E
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2BDEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2C229
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2CBF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2D464
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2D9A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2DE5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2E529
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2F021
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2F677
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":2FA7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":305CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":30B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3187C
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":31B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3259A
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":344F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":34A21
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":352B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":35CF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3620D
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":36AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":36E55
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":378F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":37F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":38443
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":38878
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":38ED6
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":39445
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":39D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3A3CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3A8A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3AF12
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3B92B
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3C15E
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3CCD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3D57A
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3D872
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3E364
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3E7DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3EC27
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3EFC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":3F845
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":40042
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":41064
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":416FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4238A
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4299D
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":43CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4446B
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4480C
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":44E06
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":458F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":460B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":468ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":471C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":47655
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":480FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4930A
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":49C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4A2EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4AA3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4AF33
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4B48E
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4C06E
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4C47E
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4C937
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4CE0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4D484
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4DBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4E3CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4E780
            Key             =   ""
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4ED5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4F167
            Key             =   ""
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4F2FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4F955
            Key             =   ""
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":4FEE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":50F77
            Key             =   ""
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":51CA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":52306
            Key             =   ""
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":52426
            Key             =   ""
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5271E
            Key             =   ""
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":52B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":52D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":52E97
            Key             =   ""
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5351C
            Key             =   ""
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":53E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":54ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":55091
            Key             =   ""
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":55543
            Key             =   ""
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":55B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":55D15
            Key             =   ""
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":55E95
            Key             =   ""
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":56017
            Key             =   ""
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":568EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":571D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":57914
            Key             =   ""
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":590F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":59B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5B66E
            Key             =   ""
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5C86F
            Key             =   ""
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5D816
            Key             =   ""
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5E011
            Key             =   ""
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5E2E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5EF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":5F86F
            Key             =   ""
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":60734
            Key             =   ""
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":61A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":61E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":62AF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":6610B
            Key             =   ""
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":69714
            Key             =   ""
         EndProperty
         BeginProperty ListImage222 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":698B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage223 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":6CEC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage224 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":6D0BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage225 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlert.frx":6D2BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   9200
      TabIndex        =   3
      Top             =   9770
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label lblCount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   200
      TabIndex        =   2
      Top             =   9770
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim selIndex As Integer
Dim sumCityName As String
Dim AQISummeryMapIndex As String
Dim AQITooltip(5) As New Tooltip

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdMapView_Click()
  On Error Resume Next
  bMapView = True
  bPicError = False
  MousePointer = 11
  If AQISummeryMap Then
    If AQISummeryMapIndex = "" Then
      MousePointer = 0
      Exit Sub
    End If
    GetUSSummeryMap AQISummeryMapIndex
    sFrmName = sumCityName & " " & frmAlert.Caption
    lstPopulation.SetFocus
  Else
    sFrmName = frmAlert.Caption
  End If
  picTureName = AQIPicArray(0)
  MousePointer = 0
  Load frmCountry
  bMapView = False
End Sub

Private Sub cmdSearch_Click()
  If cmdSearch.Caption = "Airport Info" Then
    If AirPortSummery(selIndex) = "" Then
      MsgBox "No AirPort Selected", vbInformation, "The Weather Of The World"
    Else
      GetAirportInfo AirPortSummery(selIndex)
      lstPopulation.SetFocus
    End If
  Else
    Dim sFindString As String
    sFindString = InputBox("Enter City To Find", "Weather Of The World", "Toronto", frmAlert.Left + 2000, frmAlert.Top + 4000)
    If Len(sFindString) <> 0 Then
      SearchListVw lstPopulation, sFindString
      lstPopulation.SetFocus
    End If
  End If
End Sub

Private Sub Form_Load()
  Dim cnt As Integer
 
  frmWeatherMain.Timer1.Enabled = False
  Set cmdClose.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  If AQIShowTool Then
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
    Set cmdMapView.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    cmdClose.Left = 120
  Else
    Set picHur1.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    Set picHur2.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    cmdClose.Left = 3960
    cmbcntyName.Clear
    For cnt = 0 To UBound(CountriesArray, 1)
      cmbcntyName.AddItem CountriesArray(cnt), cnt
      cmbcntyName.ListIndex = 0
    Next
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Erase AQITooltip
  Erase AQITxt
  Erase AQIbgCol
  Erase AQICanPicArray()
  Erase AQIPicArray()
  Erase AQICityMapArray()
  Erase AirPortSummery()
  frmWeatherMain.Timer1.Enabled = True
  Set frmAlert = Nothing
End Sub

Private Sub lstPopulation_DblClick()
  If cmdSearch.Caption = "Airport Info" Then
    selIndex = lstPopulation.SelectedItem.Index
    cmdSearch_Click
  End If
  If AQISummeryMap Then
    cmdMapView_Click
  End If
End Sub

Private Sub lstPopulation_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If AQISummeryMap Then
    AQISummeryMapIndex = AQICityMapArray(Item.Index)
    cmdMapView.Enabled = True
    sumCityName = Item.Text
  End If
  selIndex = Item.Index
End Sub

Private Sub picHur1_Click()
  GetMapPage slargeMapLink1
  DisplayHurMap
End Sub

Private Sub picHur1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picHur1.BorderStyle = 1
End Sub

Private Sub picHur1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picHur1.BorderStyle = 0
End Sub

Private Sub picHur2_Click()
  GetMapPage slargeMapLink2
  DisplayHurMap
End Sub

Private Sub GetMapPage(Page As String)
  RichTextBox1.Text = ""
  RichTextBox1.Text = Inet1.OpenURL(Page)
End Sub

Private Sub DisplayHurMap()
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  Dim sFieName As String
  
  On Error Resume Next
  nFileNum = FreeFile
  iIndexSt = InStr(1, RichTextBox1.Text, "Content Container", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, " ", vbTextCompare)
  myFile = Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5))
  myData() = Inet1.OpenURL(myFile, icByteArray)
  sFieName = "Large-" & Mid(myFile, InStrRev(myFile, "/") + 1)
      
  Open App.Path + "\Icons\" & sFieName For Binary Access Write As #nFileNum
    Put #nFileNum, , myData()
  Close #nFileNum
  picTureName = App.Path + "\Icons\" & sFieName
  Erase myData()
  Load frmCountry
End Sub

Private Sub picHur2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picHur2.BorderStyle = 1
End Sub

Private Sub picHur2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  picHur2.BorderStyle = 0
End Sub

Private Sub GetUSSummeryMap(sNameTag As String)
  Dim iIndex As Long, iIndex2 As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sStartPos As String
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  Dim sFieName As String
  Dim x As Integer
  'On Error Resume Next
  GetMapPage "http://www.airnow.gov/" & sNameTag
  sStartPos = """TabbedPanelsContent"""
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  
  Do
    nFileNum = FreeFile
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, " ", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5)), "legends", vbTextCompare) <> 0 Then
      iIndex = InStr(iIndexEnd, RichTextBox1.Text, "src=", vbTextCompare)
      iIndexEnd = InStr(iIndex, RichTextBox1.Text, " ", vbTextCompare)
      myFile = Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5))
    Else
      myFile = Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 1) - (iIndex + 5))
    End If
    myData() = Inet1.OpenURL(myFile, icByteArray)
    sFieName = "AQI-" & Mid(myFile, InStrRev(myFile, "/") + 1)
    
    Open App.Path + "\Icons\" & sFieName For Binary Access Write As #nFileNum
    Put #nFileNum, , myData()
    Close #nFileNum
    ReDim Preserve AQIPicArray(x)
    If x = 2 Then
      AQIPicArray(x) = myFile
    Else
      AQIPicArray(x) = App.Path + "\Icons\" & sFieName
    End If
    x = x + 1
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, """TabbedPanelsContent""", vbTextCompare)
    If iIndexSt = 0 Then Exit Do
  Loop
  Erase myData()
End Sub

Private Sub SearchListVw(LV As ListView, sFindString As String, Optional ByRef Start As Long = 1)
  'On Error GoTo SearchErr
  Dim lFound As Long
  Dim lIndex As Long
  Dim lItem As Long
  Dim lPos As Long
  Dim nameFound() As String
  Dim sCnt As Integer
  
  LV.HideSelection = False 'Needed
  LV.FullRowSelect = True 'optional....

  LV.ListItems(LV.SelectedItem.Index).Selected = False
  lIndex = 0
  For lItem = Start To LV.ListItems.Count
    sFindString = UCase$(sFindString)
    If InStr(1, UCase$(LV.ListItems(lItem).Text), sFindString, vbTextCompare) <> 0 Then
      If lIndex = 0 Then lIndex = lItem 'If the first item hasnt been selected select it now
        ReDim Preserve nameFound(lFound)
        nameFound(lFound) = lItem
        lFound = lFound + 1
     Else
       'Not Found
       LV.ListItems(lItem).Selected = False
    End If
  Next
  If lIndex = 0 Then
    MsgBox "No " & sFindString & " In Listview!", vbInformation, "Weather Of The World City Search"
    Start = 1
  Else
    For sCnt = 0 To UBound(nameFound) + 1
      If sCnt > UBound(nameFound) Then
        'Didn't find any more items
        MsgBox "No More " & sFindString & " In Countries!", vbInformation, "Weather Of The World City Search"
        Exit For
      End If
      lIndex = nameFound(sCnt)
      LV.ListItems(lIndex).Selected = True
      LV.ListItems(lIndex).EnsureVisible
      If MsgBox("Found " & LV.ListItems(lIndex).Text & vbNewLine & "Find next matching item? ", vbQuestion + vbYesNo, "Weather Of The World City Search") = vbNo Then
        LV.ListItems(lIndex).Selected = True
        LV.ListItems(lIndex).EnsureVisible
        Exit For
      End If
    Next
  End If
  Erase nameFound()
End Sub

Private Sub GetAirportInfo(sUrl As String)
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sStartPos As String
  Dim x, sCnt As Integer
  Dim sTitle, sName As String
  Dim sGPSLong As String
  Dim sLatitude, sLongitude As String
  'On Error Resume Next
  MousePointer = 11
  GetMapPage "http://www.world-airport-codes.com" & sUrl
  
  iIndex = InStr(1, RichTextBox1.Text, "var point", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "(", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ")", vbTextCompare)
  sGPSLong = Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
  sLongitude = Mid(sGPSLong, 1, InStr(1, sGPSLong, ",", vbTextCompare) - 1)
  sLatitude = Mid(sGPSLong, InStr(1, sGPSLong, ",", vbTextCompare) + 1)
  AnimationLink = "http://www.mappingsupport.com/p/gmap4.php?ll=" & sLatitude & "," & sLongitude & "&z=10&t=m&icon=pgs"
  
  sStartPos = "airportheader"
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then Exit Sub
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "<b>", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</div>", vbTextCompare)
  sTitle = Mid(RichTextBox1.Text, iIndex + 3, (iIndexEnd) - (iIndex + 3))
  sTitle = Replace(sTitle, "</b>", "")
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "airportdetails", vbTextCompare)
  
  For x = 0 To 15
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "<label class=", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
    frmAportStatus.lstAirportStat.ListItems.Add , , Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1))
    sCnt = sCnt + 1
    frmAportStatus.lstAirportStat.ListItems(sCnt).ForeColor = vbBlue
    frmAportStatus.lstAirportStat.ListItems(sCnt).Bold = True
    iIndex = InStr(iIndexSt, RichTextBox1.Text, "<span class=", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<", vbTextCompare)
    sName = Replace(Mid(RichTextBox1.Text, iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1)), "&#176;", Chr(176))
    sName = Replace(Replace(Replace(sName, "&#8217;", Chr(39)), "&#8221;", Chr(34)), Chr(10), "")
    sName = Replace(sName, "   ", "")
    frmAportStatus.lstAirportStat.ListItems(sCnt).ListSubItems.Add = sName
  Next
  frmAportStatus.lblTitle = sTitle
  frmAportStatus.Caption = sTitle
  MousePointer = 0
  frmAportStatus.Show vbModal
End Sub
