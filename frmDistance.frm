VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmDistance 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Travel Time"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmboFlags 
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   6120
      Width           =   1815
   End
   Begin VB.PictureBox picHidden 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   960
      ScaleHeight     =   1095
      ScaleWidth      =   1575
      TabIndex        =   11
      Top             =   6760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   43
      ImageHeight     =   46
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":0335
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":0908
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":0DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":12BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":17F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":20CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":29A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":2CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":33D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":375E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":3B8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDistance.frx":3F91
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9840
      Top             =   6040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   6680
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   2778
      _Version        =   393217
      TextRTF         =   $"frmDistance.frx":45A3
   End
   Begin VB.Frame frDistance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5805
      Left            =   150
      TabIndex        =   3
      Top             =   100
      Width           =   9975
      Begin VB.CommandButton cmbLocalCity 
         Caption         =   "Get Local Cities"
         Height          =   375
         Left            =   5380
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   1580
         Width           =   1335
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6220
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   960
         Width           =   1100
      End
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4860
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   960
         Width           =   1100
      End
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   7800
         ScaleHeight     =   1725
         ScaleWidth      =   1965
         TabIndex        =   12
         Top             =   260
         Visible         =   0   'False
         Width           =   1965
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   8040
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   525
         ImageHeight     =   350
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   11
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":462E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":18C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":20268
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":2764E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":2E0AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":3C488
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":43E0F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":55EDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":950B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":A3295
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":A932B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txterror 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   220
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2540
         Visible         =   0   'False
         Width           =   9515
      End
      Begin MSComctlLib.ListView lstCities 
         Height          =   1100
         Left            =   240
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "test"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "test"
            Object.Width           =   5468
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "test"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.ComboBox cmboDistance 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1060
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   2900
      End
      Begin VB.TextBox txtTo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1060
         TabIndex        =   0
         Top             =   1560
         Width           =   2900
      End
      Begin VB.TextBox txtFrom 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1060
         TabIndex        =   1
         Top             =   960
         Width           =   2900
      End
      Begin MSComctlLib.ImageList ImageList3 
         Left            =   7320
         Top             =   5120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   11
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   218
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BE181
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BE30E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BE49C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BEB55
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BECE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BEDF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BF47D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BF602
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BF78C
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BF913
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BFA9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BFC32
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":BFF47
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C00D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C05C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C0A5B
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C0F64
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C10F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C19B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C1E6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C2DDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C3200
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C3CAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C4ED5
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C5893
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C61D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C6533
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C66BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C6D9F
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C7995
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C7D36
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C81C1
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C8AEF
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C94B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":C9FCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CA583
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CA717
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CADF9
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CB4FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CBC9F
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CC154
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CC641
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CCC2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CCFCC
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CD801
               Key             =   ""
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CDDE1
               Key             =   ""
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CE437
               Key             =   ""
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CE8EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CF357
               Key             =   ""
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":CF95F
               Key             =   ""
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D0028
               Key             =   ""
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D0517
               Key             =   ""
            EndProperty
            BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D0918
               Key             =   ""
            EndProperty
            BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D0ED6
               Key             =   ""
            EndProperty
            BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D1940
               Key             =   ""
            EndProperty
            BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D21C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D2AF1
               Key             =   ""
            EndProperty
            BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D36A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D3D4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D45CB
               Key             =   ""
            EndProperty
            BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D4DDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D57A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D5ACC
               Key             =   ""
            EndProperty
            BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D6345
               Key             =   ""
            EndProperty
            BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D69F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D7672
               Key             =   ""
            EndProperty
            BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D7A9D
               Key             =   ""
            EndProperty
            BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D838F
               Key             =   ""
            EndProperty
            BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D8C82
               Key             =   ""
            EndProperty
            BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D9137
               Key             =   ""
            EndProperty
            BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D95E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D9985
               Key             =   ""
            EndProperty
            BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":D9D99
               Key             =   ""
            EndProperty
            BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DA0BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DA5BF
               Key             =   ""
            EndProperty
            BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DA965
               Key             =   ""
            EndProperty
            BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DAE4B
               Key             =   ""
            EndProperty
            BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DB7FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DC45C
               Key             =   ""
            EndProperty
            BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DC9FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DCEAF
               Key             =   ""
            EndProperty
            BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DD71E
               Key             =   ""
            EndProperty
            BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DE088
               Key             =   ""
            EndProperty
            BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DE5C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DE966
               Key             =   ""
            EndProperty
            BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DEE87
               Key             =   ""
            EndProperty
            BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DF761
               Key             =   ""
            EndProperty
            BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":DFA72
               Key             =   ""
            EndProperty
            BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E09C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E112C
               Key             =   ""
            EndProperty
            BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E15E1
               Key             =   ""
            EndProperty
            BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E176B
               Key             =   ""
            EndProperty
            BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E1C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E1DA9
               Key             =   ""
            EndProperty
            BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E1F33
               Key             =   ""
            EndProperty
            BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E2525
               Key             =   ""
            EndProperty
            BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E3390
               Key             =   ""
            EndProperty
            BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E3D4A
               Key             =   ""
            EndProperty
            BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E49A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E4E14
               Key             =   ""
            EndProperty
            BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E5ACE
               Key             =   ""
            EndProperty
            BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E5FCA
               Key             =   ""
            EndProperty
            BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E62C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E69FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E71E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E770E
               Key             =   ""
            EndProperty
            BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E795D
               Key             =   ""
            EndProperty
            BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E80BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E845D
               Key             =   ""
            EndProperty
            BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E87FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E937A
               Key             =   ""
            EndProperty
            BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":E97B9
               Key             =   ""
            EndProperty
            BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EA189
               Key             =   ""
            EndProperty
            BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EA9F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EAF38
               Key             =   ""
            EndProperty
            BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EB3ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EBAB9
               Key             =   ""
            EndProperty
            BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EC5B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":ECC07
               Key             =   ""
            EndProperty
            BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":ED00B
               Key             =   ""
            EndProperty
            BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EDB5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EE0D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EEE0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EF11D
               Key             =   ""
            EndProperty
            BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":EFB2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F1A82
               Key             =   ""
            EndProperty
            BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F1FB1
               Key             =   ""
            EndProperty
            BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F2848
               Key             =   ""
            EndProperty
            BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F3287
               Key             =   ""
            EndProperty
            BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F379D
               Key             =   ""
            EndProperty
            BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F4044
               Key             =   ""
            EndProperty
            BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F43E5
               Key             =   ""
            EndProperty
            BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F4E87
               Key             =   ""
            EndProperty
            BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F5508
               Key             =   ""
            EndProperty
            BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F59D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F5E08
               Key             =   ""
            EndProperty
            BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F6466
               Key             =   ""
            EndProperty
            BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F69D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F7291
               Key             =   ""
            EndProperty
            BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F795A
               Key             =   ""
            EndProperty
            BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F7E33
               Key             =   ""
            EndProperty
            BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F84A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F8EBB
               Key             =   ""
            EndProperty
            BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":F96EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FA268
               Key             =   ""
            EndProperty
            BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FAB0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FAE02
               Key             =   ""
            EndProperty
            BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FB8F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FBD6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FC1B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FC558
               Key             =   ""
            EndProperty
            BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FCDD5
               Key             =   ""
            EndProperty
            BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FD5D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FE5F4
               Key             =   ""
            EndProperty
            BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FEC8F
               Key             =   ""
            EndProperty
            BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FF91A
               Key             =   ""
            EndProperty
            BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":FFF2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":101234
               Key             =   ""
            EndProperty
            BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1019FB
               Key             =   ""
            EndProperty
            BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":101D9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":102396
               Key             =   ""
            EndProperty
            BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":102E89
               Key             =   ""
            EndProperty
            BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":103648
               Key             =   ""
            EndProperty
            BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":103E7D
               Key             =   ""
            EndProperty
            BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":104758
               Key             =   ""
            EndProperty
            BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":104BE5
               Key             =   ""
            EndProperty
            BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10568F
               Key             =   ""
            EndProperty
            BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10689A
               Key             =   ""
            EndProperty
            BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":107192
               Key             =   ""
            EndProperty
            BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10787F
               Key             =   ""
            EndProperty
            BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":107FCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1084C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":108A1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1095FE
               Key             =   ""
            EndProperty
            BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":109A0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":109EC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10A39B
               Key             =   ""
            EndProperty
            BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10AA14
               Key             =   ""
            EndProperty
            BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10B16F
               Key             =   ""
            EndProperty
            BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10B95B
               Key             =   ""
            EndProperty
            BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10BD10
               Key             =   ""
            EndProperty
            BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10C2EF
               Key             =   ""
            EndProperty
            BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10C6F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10C88C
               Key             =   ""
            EndProperty
            BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10CEE5
               Key             =   ""
            EndProperty
            BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10D473
               Key             =   ""
            EndProperty
            BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10E507
               Key             =   ""
            EndProperty
            BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10F233
               Key             =   ""
            EndProperty
            BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10F896
               Key             =   ""
            EndProperty
            BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10F9B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":10FCAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11010B
               Key             =   ""
            EndProperty
            BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11029C
               Key             =   ""
            EndProperty
            BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":110427
               Key             =   ""
            EndProperty
            BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":110AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1113F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11204C
               Key             =   ""
            EndProperty
            BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":112621
               Key             =   ""
            EndProperty
            BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":112AD3
               Key             =   ""
            EndProperty
            BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11311C
               Key             =   ""
            EndProperty
            BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1132A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":113425
               Key             =   ""
            EndProperty
            BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1135A7
               Key             =   ""
            EndProperty
            BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":113E7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":114768
               Key             =   ""
            EndProperty
            BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":114EA4
               Key             =   ""
            EndProperty
            BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":116683
               Key             =   ""
            EndProperty
            BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":1170F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":118BFE
               Key             =   ""
            EndProperty
            BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":119DFF
               Key             =   ""
            EndProperty
            BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11ADA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11B5A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11B875
               Key             =   ""
            EndProperty
            BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11C4F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11CDFF
               Key             =   ""
            EndProperty
            BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11DCC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11F019
               Key             =   ""
            EndProperty
            BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistance.frx":11F429
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image cmdimgReverse 
         Height          =   400
         Left            =   4100
         MousePointer    =   99  'Custom
         Stretch         =   -1  'True
         ToolTipText     =   "Reverse Location"
         Top             =   1250
         Width           =   400
      End
      Begin VB.Image imgFlag2 
         Height          =   250
         Left            =   9320
         Stretch         =   -1  'True
         Top             =   4090
         Width           =   400
      End
      Begin VB.Image imgFlag1 
         Height          =   250
         Left            =   1000
         Stretch         =   -1  'True
         Top             =   4090
         Width           =   400
      End
      Begin VB.Label sMjCityInfo 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   320
         TabIndex        =   31
         Top             =   3750
         Width           =   9495
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   7
         Left            =   6080
         TabIndex        =   30
         Top             =   5370
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   6080
         TabIndex        =   29
         Top             =   5055
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   3
         Left            =   6080
         TabIndex        =   28
         Top             =   4755
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   6080
         TabIndex        =   27
         Top             =   4440
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   5000
         TabIndex        =   26
         Top             =   5370
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   5000
         TabIndex        =   25
         Top             =   5055
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5000
         TabIndex        =   24
         Top             =   4755
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblSecCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   5000
         TabIndex        =   23
         Top             =   4440
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   7
         Left            =   1400
         TabIndex        =   22
         Top             =   5370
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   5
         Left            =   1400
         TabIndex        =   21
         Top             =   5055
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   3
         Left            =   1400
         TabIndex        =   20
         Top             =   4755
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Index           =   1
         Left            =   1400
         TabIndex        =   19
         Top             =   4440
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   320
         TabIndex        =   18
         Top             =   5370
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   320
         TabIndex        =   17
         Top             =   5055
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   320
         TabIndex        =   16
         Top             =   4755
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblFrsCity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   320
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label lblSecTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   5000
         TabIndex        =   14
         Top             =   4080
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblFrsTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   315
         TabIndex        =   13
         Top             =   4080
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Image Image2 
         Height          =   600
         Left            =   5800
         Top             =   180
         Width           =   600
      End
      Begin VB.Image imgPicture 
         Appearance      =   0  'Flat
         Height          =   1725
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   260
         Width           =   1965
      End
      Begin VB.Label lblMiles 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   4875
         TabIndex        =   8
         Top             =   2865
         Width           =   135
      End
      Begin VB.Label lblCountry 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   580
         Left            =   240
         TabIndex        =   7
         Top             =   2175
         Width           =   9495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   580
         TabIndex        =   6
         Top             =   1600
         Width           =   360
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   350
         TabIndex        =   5
         Top             =   1020
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Get:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   520
         TabIndex        =   4
         Top             =   400
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isCity As Boolean
Dim bAirError As Boolean
Dim oldtxtTo As String
Dim oldtxtFrom As String
Dim ComboIndesTxt As String
Dim iNameCnt As Integer
Dim ctycnt As Integer
Dim NewFormH As Long
Dim NewFrameH As Long

Private Sub cmboDistance_Click()
  Dim rndNumber As Integer
  On Error Resume Next
  ComboIndesTxt = cmboDistance.List(cmboDistance.ListIndex)
  Image2.Left = 5800
  
  Set Image2.Picture = ImageList1.ListImages(6).Picture
  lstCities.Visible = False
  cmbLocalCity.Visible = False
  
  Select Case ComboIndesTxt
    Case "Latitude/Longitude"
      SetPictureBox frmDistance, "", 3
      txtTo.Visible = False
      cmdimgReverse.Visible = False
      Label3.Visible = False
      Label2.Caption = "Of:"
      cmdCalculate.Enabled = True
    Case "Airlines Flying", "Major Cities", "Hotels In The Area", "Major Airports"
      If ComboIndesTxt = "Major Cities" Then
        SetPictureBox frmDistance, "", 6
      ElseIf ComboIndesTxt = "Major Airports" Then
        SetPictureBox frmDistance, "", 8
      ElseIf ComboIndesTxt = "Airlines Flying" Then
        SetPictureBox frmDistance, "", 1
        Label3.Visible = False
        Label2.Caption = "To:"
      Else
        SetPictureBox frmDistance, "", 7
      End If
      txtTo.Visible = False
      cmdimgReverse.Visible = False
      If ComboIndesTxt <> "Airlines Flying" Then
        Label3.Visible = False
        Label2.Caption = "Near/In:"
      End If
      cmdCalculate.Enabled = True
    Case "Flight Time", "Flight Distance", "Driving Distance", "Drive Time", "Time Difference"
      If ComboIndesTxt = "Flight Time" Then
        SetPictureBox frmDistance, "", 9
      ElseIf ComboIndesTxt = "Flight Distance" Then
        SetPictureBox frmDistance, "", 10
      ElseIf ComboIndesTxt = "Driving Distance" Then
        SetPictureBox frmDistance, "", 2
      ElseIf ComboIndesTxt = "Drive Time" Then
        SetPictureBox frmDistance, "", 11
      Else
        SetPictureBox frmDistance, "", 4
      End If
      txtTo.Visible = True
      cmdimgReverse.Visible = True
      Label2.Caption = "From:"
      Label3.Visible = True
      If Len(Trim(txtTo.Text)) = 0 Or Len(Trim(txtFrom.Text)) = 0 Then
        cmdCalculate.Enabled = False
      Else
        cmdCalculate.Enabled = True
      End If
  End Select
  If Len(Trim(txtFrom.Text)) = 0 Then
    cmdCalculate.Enabled = False
  End If
  lblCountry.Caption = ""
  lblMiles.Caption = ""
  txterror.Text = ""
  frmDistance.Caption = ComboIndesTxt
  NewFrameH = 2175
  NewFormH = 2800
  centerForm
  cmdExit.SetFocus
End Sub

Private Sub cmdCalculate_Click()
  Dim sPageLink As String
  
  lblCountry.ForeColor = vbBlack
  lblCountry.Caption = "Calculating..."
  lstCities.Visible = False
  
  Select Case ComboIndesTxt
    Case "Latitude/Longitude"
      sPageLink = "http://www.travelmath.com/lat-long/"
    Case "Major Cities"
      GetCityTag
      If isCity Or InStr(1, txtFrom, ",", vbTextCompare) <> 0 Or Nozip Then
        sPageLink = "http://www.travelmath.com/cities-near/"
      Else
        sPageLink = "http://www.travelmath.com/cities-in/"
      End If
    Case "Hotels In The Area"
      sPageLink = "http://www.travelmath.com/hotels-near/"
    Case "Airlines Flying"
      sPageLink = "http://www.travelmath.com/airlines/"
    Case "Major Airports"
      GetCityTag
      If Not isCity Or InStr(1, txtFrom, ",", vbTextCompare) <> 0 Or Nozip Then
        sPageLink = "http://www.travelmath.com/airports-in/"
      Else
        sPageLink = "http://www.travelmath.com/closest-airport/"
      End If
    Case Else
      sPageLink = "http://www.travelmath.com/" & Replace(LCase(ComboIndesTxt), " ", "-") & "/from/"
  End Select
  GetDistance sPageLink, txtFrom, txtTo, ComboIndesTxt
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdimgReverse_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  cmdimgReverse.BorderStyle = 1
End Sub

Private Sub cmdimgReverse_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  cmdimgReverse.BorderStyle = 0
End Sub

Private Sub Form_Load()
  Dim cnt As Integer
  
  Set cmdExit.MouseIcon = ImageList1.ListImages(8).Picture
  Set cmdCalculate.MouseIcon = ImageList1.ListImages(8).Picture
  Set cmdimgReverse.MouseIcon = ImageList1.ListImages(8).Picture
  Set cmbLocalCity.MouseIcon = ImageList1.ListImages(8).Picture
  Set Image2.Picture = ImageList1.ListImages(6).Picture
  Set cmdimgReverse.Picture = ImageList1.ListImages(9).Picture
  SetPictureBox frmDistance, "", 1
  cmboDistance.AddItem "Flight Distance"
  cmboDistance.AddItem "Flight Time"
  cmboDistance.AddItem "Airlines Flying"
  cmboDistance.AddItem "Major Airports"
  cmboDistance.AddItem "Driving Distance"
  cmboDistance.AddItem "Drive Time"
  cmboDistance.AddItem "Time Difference"
  cmboDistance.AddItem "Major Cities"
  cmboDistance.AddItem "Hotels In The Area"
  cmboDistance.AddItem "Latitude/Longitude"
  cmboDistance.ListIndex = 0
  txtFrom.Text = frmWeatherMain.lblCity.Caption
  
  cmboFlags.Clear
  
  For cnt = 0 To UBound(CountriesArray, 1)
   cmboFlags.AddItem CountriesArray(cnt), cnt
   cmboFlags.ListIndex = 0
  Next
End Sub

Private Sub GetDistance(sPageName As String, sStringFrom As String, sStringTo As String, cmboIndex As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sWeblink As String
  Dim sStartPos As String
  Dim sLongLatmin As String
  Dim sLongLatlon As String
  Dim x As Integer
  Dim cnt As Integer
  Dim slimits As Integer
  Dim errMessage As String
  
  On Error GoTo errorTrap
  lblMiles.Caption = ""
  sMjCityInfo.Caption = ""
  cmbLocalCity.Visible = False
  If InStr(1, sStringFrom, "+", vbTextCompare) = 0 Then
    sStringFrom = Replace(sStringFrom, " ", "+")
  End If
  If InStr(1, sStringTo, "+", vbTextCompare) = 0 Then
    sStringTo = Replace(sStringTo, " ", "+")
  End If
  
  If cmboIndex = "Latitude/Longitude" Then
    sWeblink = sPageName & sStringFrom
    GetWebpage sWeblink
  ElseIf cmboIndex = "Major Cities" Or cmboIndex = "Hotels In The Area" Or ComboIndesTxt = "Airlines Flying" Or ComboIndesTxt = "Major Airports" Then
    sWeblink = sPageName & sStringFrom
    GetWebpage sWeblink
  Else
    sWeblink = sPageName & sStringFrom & "/to/" & sStringTo
    GetWebpage sWeblink
  End If
  sStartPos = "boxtop"
  
  bAirError = False
  Image2.Left = 4100
  lblCountry.Visible = True
  lblMiles.FontSize = 14
  lblCountry.ForeColor = vbBlack
  txterror.Visible = False
  Select Case cmboIndex
    Case "Flight Distance"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "flight-d"
        Exit Sub
      Else
        Set Image2.Picture = ImageList1.ListImages(1).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "flight-d", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
        
        lblCountry.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<", vbTextCompare)
        lblMiles.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<", vbTextCompare)
        lblMiles.Caption = lblMiles.Caption & Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, ">", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
        
        lblMiles.Caption = lblMiles.Caption & Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
        GetCityInfo iIndexSt, True
      End If
    Case "Flight Time"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "flight-t"
        Exit Sub
      Else
        Set Image2.Picture = ImageList1.ListImages(2).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndex = InStr(iIndex, RichTextBox1.Text, "flight-t", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblMiles.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        GetCityInfo iIndexEnd, True
      End If
    Case "Driving Distance"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "driving-d"
        Exit Sub
      Else
        Set Image2.Picture = ImageList1.ListImages(10).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndex = InStr(iIndex, RichTextBox1.Text, "driving-d", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<", vbTextCompare)
        lblMiles.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblMiles.Caption = lblMiles.Caption & " / " & Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        If InStr(1, lblMiles.Caption, "font", vbTextCompare) <> 0 Then
          lblMiles.Caption = "Unable To Calculate Driving Distance"
        End If
        GetCityInfo iIndexEnd, True
      End If
    Case "Drive Time"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "drive-t"
        Exit Sub
      Else
        Set Image2.Picture = ImageList1.ListImages(11).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndex = InStr(iIndex, RichTextBox1.Text, "drive-t", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblMiles.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        If InStr(1, lblMiles.Caption, "font", vbTextCompare) <> 0 Then
          lblMiles.Caption = "Unable To Calculate Drive Time"
        End If
        GetCityInfo iIndexEnd, True
      End If
    Case "Time Difference"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "time-d"
        Exit Sub
      Else
        lblCountry.Caption = ""
        Set Image2.Picture = ImageList1.ListImages(3).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndex = InStr(iIndex, RichTextBox1.Text, "time-d", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<", vbTextCompare)
        lblMiles.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblMiles.Caption = lblMiles.Caption & Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "</span>", "")
        If Len(lblMiles.Caption) > 65 Then
          lblMiles.FontSize = 12
        Else
          lblMiles.FontSize = 14
        End If
      End If
      GetCityInfo iIndexEnd, True
    Case "Major Cities"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "grid-g"
        Exit Sub
      Else
        lstCities.ListItems.Clear
        lstCities.Visible = True
        Set Image2.Picture = ImageList1.ListImages(5).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "grid-g", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "<br />", ""), Chr(10), " ")
        'Get cities
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, """top"">", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "</table>", vbTextCompare)
        RemoveHttpLink Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt) - (iIndex + 6))
        GetMoreInfo iIndexSt
        GetCityInfo iIndexEnd, False
        bAirError = False
        If InStr(1, lblFrsCity(7).Caption, "cities") = 0 Then
          cmbLocalCity.Caption = "Local Cities"
          cmbLocalCity.Visible = True
        End If
      End If
    Case "Hotels In The Area"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "hotels"
        Exit Sub
      Else
        lstCities.ListItems.Clear
        lstCities.Visible = True
        Set Image2.Picture = ImageList1.ListImages(12).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "hotels", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "<br />", ""), Chr(10), " ")
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, """top"">", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "</table>", vbTextCompare)
        RemoveHttpLink Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt) - (iIndex + 6))
        GetMoreInfo iIndexSt
        GetCityInfo iIndexSt, False
      End If
    Case "Airlines Flying"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        cmbLocalCity.Caption = "Closest Airport"
        bAirError = True
        DisplayError "airport"
        Exit Sub
      Else
        lstCities.ListItems.Clear
        lstCities.Visible = True
        Set Image2.Picture = ImageList1.ListImages(13).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, "airport", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        lblCountry.Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "<br />", ""), Chr(10), "")
        iIndex = iIndexEnd
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "style=", vbTextCompare)
        If InStr(1, Mid(RichTextBox1.Text, iIndex, 25), "class=", vbTextCompare) <> 0 Then
          txterror.Visible = True
          iIndex = InStr(iIndex, RichTextBox1.Text, "class=", vbTextCompare)
          iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
          iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<a", vbTextCompare)
          errMessage = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), Chr(10), " ")
          If InStr(1, errMessage, "<big>", vbTextCompare) <> 0 Then
            iIndex = InStr(1, errMessage, "<big>", vbTextCompare)
            iIndexSt = InStr(iIndex, errMessage, "</big>", vbTextCompare)
            errMessage = Mid(errMessage, iIndex + 5, (iIndexSt) - (iIndex + 5))
            If InStr(1, errMessage, "<big>", vbTextCompare) <> 0 Then
              errMessage = Mid(Replace(errMessage, "<big>", ""), 1, InStr(1, errMessage, "</big>"))
            End If
            iNameCnt = 0
            txterror.Visible = False
            lstCities.ListItems.Clear
            lstCities.Visible = True
            ParseString errMessage
            GetCityInfo iIndexSt, False
            NewFrameH = 5800
            NewFormH = 6420
            centerForm
            Exit Sub
          End If
          iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<strong>", vbTextCompare)
          iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
          errMessage = errMessage & " " & Mid(RichTextBox1.Text, iIndex + 8, (iIndexSt) - (iIndex + 8))
          txterror.Text = errMessage
          cmbLocalCity.Caption = "Closest Airport"
          cmbLocalCity.Visible = True
          cmbLocalCity.SetFocus
          bAirError = True
          lstCities.Visible = False
          NewFrameH = 3575
          NewFormH = 4200
          centerForm
          Exit Sub
        End If
        iNameCnt = 0
        ctycnt = 0
        iIndex = InStr(iIndexEnd, RichTextBox1.Text, "valign=", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</table>", vbTextCompare)
        ParseString Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
        GetCityInfo iIndexEnd, False
        cmbLocalCity.Caption = "Departure City"
        cmbLocalCity.Visible = True
        cmbLocalCity.SetFocus
      End If
    Case "Major Airports"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "airport"
        Exit Sub
      Else
        If isCity Then
          GetClosestAirport
          Set Image2.Picture = ImageList1.ListImages(13).Picture
        Else
          lstCities.ListItems.Clear
          lstCities.Visible = True
          Set Image2.Picture = ImageList1.ListImages(13).Picture
          iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
          iIndexSt = InStr(iIndex, RichTextBox1.Text, "airport", vbTextCompare)
          iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
          iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
          lblCountry.Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "<br />", ""), Chr(10), " ")
          'Get Airports
          iIndex = InStr(iIndexEnd, RichTextBox1.Text, """top"">", vbTextCompare)
          iIndexSt = InStr(iIndex, RichTextBox1.Text, "</table>", vbTextCompare)
          RemoveHttpLink Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt) - (iIndex + 6))
          GetMoreInfo iIndexSt
          GetCityInfo iIndexSt, False
          cmbLocalCity.Caption = "Local Airports"
          cmbLocalCity.Visible = True
          cmbLocalCity.SetFocus
        End If
      End If
    Case "Latitude/Longitude"
      If InStr(1, RichTextBox1.Text, ">ERROR:<", vbTextCompare) <> 0 Then
        DisplayError "lat-l"
        Exit Sub
      Else
        Set Image2.Picture = ImageList1.ListImages(4).Picture
        iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
        iIndex = InStr(iIndex, RichTextBox1.Text, "lat-l", vbTextCompare)
        iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h1", vbTextCompare)
        lblCountry.Caption = Replace(Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), "<br />", ""), Chr(10), " ")
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<", vbTextCompare)
        sLongLatmin = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "</span", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
        sLongLatmin = sLongLatmin & " / " & Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "Latitude:", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "<br ", vbTextCompare)
        sLongLatlon = "Latitude:" & Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "Longitude:", vbTextCompare)
        iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
        iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</div>", vbTextCompare)
        sLongLatlon = sLongLatlon & "  Longitude:" & Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
        lblMiles.FontSize = 11
        lblMiles.Caption = Replace(sLongLatmin, "&deg;", Chr(176)) & "  OR  " & sLongLatlon
      End If
  End Select
xxx:
  If ComboIndesTxt = "Latitude/Longitude" Then
    NewFrameH = 3575
    NewFormH = 4150
  Else
    NewFormH = 6420
    NewFrameH = 5800
  End If
  centerForm
  Exit Sub
errorTrap:
  If Err.Number = 5 Then
    MsgBox "Could not Find any " & cmboIndex & " in " & sStringFrom, vbInformation, "Weather Of The World"
  Else
    MsgBox "Error Number " & Err.Number & " Has occurred, Please select another location", vbInformation, "Weather Of The World"
  End If
  Image2.Left = 5800
  Set Image2.Picture = ImageList1.ListImages(7).Picture
End Sub

Private Sub GetWebpage(sWebPage)
  RichTextBox1.Text = ""
  RichTextBox1.Text = Inet1.OpenURL(sWebPage)
End Sub

Private Sub cmdimgReverse_Click()
  oldtxtTo = txtTo.Text
  oldtxtFrom = txtFrom.Text
  If txtFrom.Text <> oldtxtTo Then
    txtTo.Text = txtFrom.Text
    txtFrom.Text = oldtxtTo
  Else
    txtTo.Text = oldtxtTo
    txtFrom.Text = oldtxtFrom
  End If
End Sub

Private Sub cmbLocalCity_Click()
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sStartPos As String
  
  lstCities.ListItems.Clear
  lstCities.Visible = True
  cmbLocalCity.Visible = False
  If Not bAirError Then
    sStartPos = "<div id=""travelmap"""
    iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "<h2 class=", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1)), "Map", vbTextCompare) <> 0 Then
      lblCountry.Caption = Mid(RichTextBox1.Text, iIndex + 8, (iIndexEnd) - (iIndex + 8))
    Else
      lblCountry.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
    End If
    'Get cities
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, """top"">", vbTextCompare)
    If iIndex = 0 Then
      MsgBox "No more cities to display", vbInformation, "Weather Of The World"
      NewFrameH = 2175
      NewFormH = 2800
      centerForm
      Image2.Left = 5800
      Set Image2.Picture = ImageList1.ListImages(6).Picture
      cmdCalculate.SetFocus
      Exit Sub
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</table>", vbTextCompare)
    RemoveHttpLink Mid(RichTextBox1.Text, iIndex + 6, (iIndexSt) - (iIndex + 6))
    GetMoreInfo iIndexSt
  Else
    If ComboIndesTxt = "Airlines Flying" Then
      Dim sWeblink As String
      Dim sStrFind As String
      
      lstCities.Visible = False
      sStrFind = Replace(txtFrom.Text, " ", "+")
      sWeblink = "http://www.travelmath.com/closest-airport/" & sStrFind
      GetWebpage sWeblink
      NewFormH = 6420
      NewFrameH = 5800
      centerForm
    End If
    GetClosestAirport
  End If
End Sub

Private Sub txtFrom_Change()
  cmbLocalCity.Visible = False
  Select Case ComboIndesTxt
    Case "Major Cities", "Hotels In The Area", "Airlines Flying", "Latitude/Longitude", "Major Airports"
      If Len(Trim(txtFrom.Text)) = 0 Then
        cmdCalculate.Enabled = False
      Else
        cmdCalculate.Enabled = True
      End If
    Case Else
      If Len(Trim(txtTo.Text)) = 0 Or Len(Trim(txtFrom.Text)) = 0 Then
        cmdCalculate.Enabled = False
      Else
        cmdCalculate.Enabled = True
      End If
    End Select
End Sub

Private Sub txtTo_Change()
  If Len(Trim(txtTo.Text)) = 0 Or Len(Trim(txtFrom.Text)) = 0 Then
    cmdCalculate.Enabled = False
  Else
    cmdCalculate.Enabled = True
  End If
End Sub

Private Sub DisplayError(sStartPos As String)
  Dim iStartIndex As Long
  Dim iEndIndex As Long
  Dim iNewIndes As Long

  Image2.Left = 5800
  Set Image2.Picture = ImageList1.ListImages(7).Picture
  iNewIndes = InStr(1, RichTextBox1.Text, "boxtop", vbTextCompare)
  iStartIndex = InStr(iNewIndes, RichTextBox1.Text, sStartPos, vbTextCompare)
  iEndIndex = InStr(iStartIndex, RichTextBox1.Text, ">", vbTextCompare)
  iNewIndes = InStr(iEndIndex, RichTextBox1.Text, "</", vbTextCompare)
  lblCountry.ForeColor = vbRed
  lblCountry.Caption = Mid(RichTextBox1.Text, iEndIndex + 1, (iNewIndes) - (iEndIndex + 1)) & vbCrLf & _
                       "ERROR: Please try using country only for your entries:"
  
  iStartIndex = InStr(iNewIndes, RichTextBox1.Text, "<big>", vbTextCompare)
  iEndIndex = InStr(iStartIndex, RichTextBox1.Text, "</big>", vbTextCompare)
  lblMiles.Caption = "The location " & Mid(RichTextBox1.Text, iStartIndex + 5, (iEndIndex) - (iStartIndex + 5)) & " could not be found."
  NewFormH = 4000
  NewFrameH = 3400
  centerForm
End Sub

Private Sub ParseString(StringToParse As String)
  Dim x As Integer
  Dim NameArray() As String
  lstCities.ListItems.Clear
  StringToParse = Replace(Replace(Replace(Replace(StringToParse, "</td>", "<br />"), "</tr>", ""), "<td valign=""top"">", ""), "</big>", "")
  NameArray() = Split(StringToParse, "<br />")
  For x = 0 To UBound(NameArray, 1) - 1
    If x Mod 3 = 0 Then
      lstCities.ListItems.Add , , NameArray(x)
      iNameCnt = iNameCnt + 1
    Else
      lstCities.ListItems(iNameCnt).ListSubItems.Add , , NameArray(x)
      ctycnt = ctycnt + 1
    End If
  Next
  If UBound(NameArray, 1) >= 1 Then
    lblCountry.Caption = UBound(NameArray, 1) & " " & lblCountry.Caption
  Else
    lblCountry.Caption = 1 & " " & lblCountry.Caption
  End If
  Erase NameArray
End Sub

Private Sub RemoveHttpLink(StringToParse As String)
  Dim iStartIndex As Long
  Dim iEndIndex As Long
  Dim iNewIndes As Long
  Dim x As Integer
  Dim sCityNames As String
  Dim newString As String
  Dim NameArray() As String
  lstCities.ListItems.Clear
  
  NameArray() = Split(StringToParse, "</a>")
  newString = StringToParse
  For x = 0 To UBound(NameArray, 1)
    If InStr(1, StringToParse, "<a href=", vbTextCompare) <> 0 Then
      iNewIndes = InStr(1, StringToParse, "<a href=", vbTextCompare)
      iStartIndex = InStr(iNewIndes, StringToParse, ">", vbTextCompare)
      iEndIndex = InStr(iStartIndex, StringToParse, "</a>", vbTextCompare)
      sCityNames = sCityNames & Mid(StringToParse, iStartIndex + 1, (iEndIndex + 4) - (iStartIndex + 1))
      newString = Mid(StringToParse, iStartIndex + 1)
    End If
    StringToParse = newString
  Next
  If Len(sCityNames) < 1 Then
    txterror.Visible = True
    txterror.Text = "No " & ComboIndesTxt
    Exit Sub
  End If
  'Get cities
  NameArray() = Split(sCityNames, "</a>")
  iNameCnt = 0
  ctycnt = 0
  For x = 0 To UBound(NameArray, 1) - 1
    If x Mod 3 = 0 Then
      lstCities.ListItems.Add , , NameArray(x)
      iNameCnt = iNameCnt + 1
    Else
      lstCities.ListItems(iNameCnt).ListSubItems.Add , , NameArray(x)
      ctycnt = ctycnt + 1
    End If
  Next
  If UBound(NameArray, 1) >= 1 Then
    lblCountry.Caption = UBound(NameArray, 1) & " " & lblCountry.Caption
  Else
    lblCountry.Caption = 1 & " " & lblCountry.Caption
  End If
  NewFormH = 6420
  NewFrameH = 5800
  Erase NameArray
End Sub

Private Sub GetClosestAirport()
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim Limits As Integer
  Dim sStartPos As String
  Dim sPageLink As String
  
  ctycnt = 0
  iNameCnt = 0
  lstCities.ListItems.Clear
  lstCities.Visible = True
  sStartPos = "boxtop"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "airport", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</h", vbTextCompare)
  lblCountry.Caption = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  Do
    If InStr(1, Mid(RichTextBox1.Text, iIndexEnd + 1, 5), "Find", vbTextCompare) = 0 Then
      If iIndex = 0 Then
        MsgBox "No more cities to display", vbInformation, "Weather Of The World"
        NewFrameH = 2175
        NewFormH = 2800
        centerForm
        Image2.Left = 5800
        Set Image2.Picture = ImageList1.ListImages(6).Picture
        cmdCalculate.SetFocus
        Exit Sub
      End If
      iIndexEnd = InStr(iIndexEnd, RichTextBox1.Text, "/airport", vbTextCompare)
      If iIndexEnd = 0 Then Exit Sub
      iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<strong>", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
      If ctycnt Mod 3 = 0 Then
        lstCities.ListItems.Add , , Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
        iNameCnt = iNameCnt + 1
      Else
        lstCities.ListItems(iNameCnt).ListSubItems.Add , , Mid(RichTextBox1.Text, iIndexSt + 8, (iIndexEnd) - (iIndexSt + 8))
      End If
    Else
      Limits = 1
      txterror.Visible = False
    End If
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    ctycnt = ctycnt + 1
  Loop Until Limits = 1
  GetCityInfo iIndexEnd, False
End Sub

Private Sub GetCityTag()
  Dim txtToFind As String
  If InStr(1, txtFrom.Text, ",", vbTextCompare) <> 0 Then
    txtToFind = Mid(txtFrom.Text, 1, InStr(1, txtFrom.Text, ",", vbTextCompare) - 1)
  Else
    txtToFind = txtFrom.Text
  End If
   Dim oFndNode As Node
   Set oFndNode = TreeFindNode(frmWeatherMain.TView, txtToFind, True, 1)
   Set oFndNode = Nothing
End Sub

Function TreeFindNode(tvFind As TreeView, ByVal sFindItem As String, Optional bSearchAll As Boolean = True, Optional lItemIndex As Long = 1) As Node
   Dim oThisNode As Node, bSearch As Boolean, lInstance As Long
    
   sFindItem = UCase$(sFindItem)
   bSearch = True
   isCity = False
   
   For Each oThisNode In tvFind.Nodes
      If bSearchAll = False Then
         'Only Search Top Level Nodes
         If (oThisNode.Parent Is Nothing) = False Then
            bSearch = False
         Else
            bSearch = True
         End If
      End If
      If bSearch Then
         If (UCase$(oThisNode.Text) Like sFindItem) = True Then
            lInstance = lInstance + 1
            If lInstance >= lItemIndex Then
              If Len(oThisNode.Parent) > 2 Then
                isCity = True
              Else
                isCity = False
              End If
              Exit For
            End If
         End If
      End If
   Next
End Function

Private Sub GetCityInfo(iStart As Long, bCityOnly As Boolean)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim Limits As Integer
  Dim sStartPos As String
  Dim sLastPos As String
  Dim iCnt As Integer
  
  For iCnt = 0 To 7
    lblFrsCity(iCnt).Caption = ""
    lblSecCity(iCnt).Caption = ""
    lblFrsCity(iCnt).Visible = False
    lblSecCity(iCnt).Visible = False
  Next
  lblFrsTitle.Visible = False
  lblSecTitle.Visible = False
  imgFlag1.Visible = False
  imgFlag2.Visible = False
  
  iCnt = 0
   If bCityOnly Then
    Do
      iIndex = InStr(iStart + 22, RichTextBox1.Text, "</table>", vbTextCompare)
      iStart = iIndex
    Loop Until InStr(iIndex + 22, RichTextBox1.Text, "</table>", vbTextCompare) = 0
  Else
    Do
      iIndex = InStr(iStart + 22, RichTextBox1.Text, "clear:both", vbTextCompare)
      iStart = iIndex
    Loop Until InStr(iIndex + 22, RichTextBox1.Text, "clear:both", vbTextCompare) = 0
  End If
  If ComboIndesTxt = "Hotels In The Area" Then
    iStart = InStr(iStart + 10, RichTextBox1.Text, "travelmap", vbTextCompare)
  End If
  sStartPos = "traveling"
  iIndex = InStr(iStart, RichTextBox1.Text, "traveling", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
  lblFrsTitle.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
  lblFrsTitle.Visible = True
  For Limits = 0 To 3
    iIndex = InStr(iIndexEnd + 7, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    sLastPos = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "</a>", "")
    lblFrsCity(iCnt).Caption = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "</a>", "")
    lblFrsCity(iCnt).Visible = True
    iCnt = iCnt + 1
    lblFrsCity(iCnt).Visible = True
    iIndex = InStr(iIndexSt + 13, RichTextBox1.Text, "href=", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 1, 20), "href=", vbTextCompare) = 0 Then
      iIndex = iIndexSt
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<br />", vbTextCompare)
    lblFrsCity(iCnt).Caption = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), "</a>", "")
    lblFrsCity(iCnt).Visible = True
    If sLastPos = "Category:" Then
      cmdExit.SetFocus
      Exit For
    End If
    If lblFrsCity(iCnt - 1).Caption = "Country:" Then
      sfndResult = FindStringinListControl(cmboFlags, lblFrsCity(iCnt).Caption)
      If sfndResult <> -1 Then
        imgFlag1.Visible = True
        imgFlag1.Picture = ImageList3.ListImages(sfndResult + 1).Picture
        imgFlag1.Top = lblFrsCity(iCnt).Top + 5
        imgFlag1.Left = lblFrsCity(iCnt).Left + lblFrsCity(iCnt).Width + 70
      End If
    End If
    iCnt = iCnt + 1
  Next
  If InStr(iIndexEnd, RichTextBox1.Text, "clear:both", vbTextCompare) = 0 Then
    cmdExit.SetFocus
    Exit Sub
  End If
  iCnt = 0
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "traveling", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
  lblSecTitle.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1))
  lblSecTitle.Visible = True
  For Limits = 0 To 3
    iIndex = InStr(iIndexEnd + 7, RichTextBox1.Text, ">", vbTextCompare)
    iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    sLastPos = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "</a>", "")
    lblSecCity(iCnt).Caption = Replace(Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt) - (iIndex + 1)), "</a>", "")
    lblSecCity(iCnt).Visible = True
    iCnt = iCnt + 1
    iIndex = InStr(iIndexSt + 13, RichTextBox1.Text, "href=", vbTextCompare)
    If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 1, 20), "href=", vbTextCompare) = 0 Then
      iIndex = iIndexSt
    End If
    iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "<br />", vbTextCompare)
    lblSecCity(iCnt).Caption = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd) - (iIndexSt + 1)), "</a>", "")
    lblSecCity(iCnt).Visible = True
    If sLastPos = "Category:" Then
      
      Exit For
    End If
    If lblSecCity(iCnt - 1).Caption = "Country:" Then
      sfndResult = FindStringinListControl(cmboFlags, lblSecCity(iCnt).Caption)
      If sfndResult <> -1 Then
        imgFlag2.Visible = True
        imgFlag2.Picture = ImageList3.ListImages(sfndResult + 1).Picture
        imgFlag2.Top = lblSecCity(iCnt).Top + 5
        imgFlag2.Left = lblSecCity(iCnt).Left + lblSecCity(iCnt).Width + 70
      End If
    End If
    iCnt = iCnt + 1
  Next
End Sub

Private Sub centerForm()
  frmDistance.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmDistance.Width / 2)
  frmDistance.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (NewFormH / 2)
  frDistance.Height = NewFrameH
  frmDistance.Height = NewFormH
End Sub

Private Sub GetMoreInfo(iIndexSt As Long)
  Dim iIndex As Long
  Dim tempStr As String
  
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "class=", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  tempStr = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndex) - (iIndexSt + 1))
  tempStr = Mid(tempStr, 1, InStr(1, tempStr, ".") + 1)
  sMjCityInfo.Caption = tempStr
End Sub
