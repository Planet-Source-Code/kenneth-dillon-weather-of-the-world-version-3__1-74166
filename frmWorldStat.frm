VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWorldStat 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "7 Wonders of the Modern World"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   9360
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1275
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9360
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   12360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   9900
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4260
      _Version        =   393217
      TextRTF         =   $"frmWorldStat.frx":0000
   End
   Begin VB.PictureBox picHidden 
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   1920
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   9940
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame fmMap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   5495
      Left            =   120
      TabIndex        =   3
      Top             =   100
      Width           =   6255
      Begin VB.PictureBox picMain 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4895
         Left            =   240
         ScaleHeight     =   4860
         ScaleWidth      =   5745
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.Image imgPicture 
         Height          =   4895
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   5775
      End
   End
   Begin VB.Label lblStatInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   3100
      Left            =   220
      TabIndex        =   8
      Top             =   6120
      Width           =   6055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "label"
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
      Height          =   390
      Left            =   2970
      TabIndex        =   7
      Top             =   5610
      Width           =   555
   End
End
Attribute VB_Name = "frmWorldStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ipageNum As Integer
Dim ipageCnt As Integer

Private Sub cmdNext_Click()
  If ipageNum <= ipageCnt Then
    ipageNum = ipageNum + 1
    cmdPrevious.Enabled = True
  End If
  If ipageNum = ipageCnt Then
    cmdNext.Enabled = False
  End If
  If Not isTallest Then
    GetsevenWonderStat "http://fun.familyeducation.com/slideshow/historic-sites/61628.html?page=" & ipageNum
  Else
    GetTallestBuilding "http://fun.familyeducation.com/slideshow/historic-sites/61490.html?page=" & ipageNum
  End If
End Sub

Private Sub cmdPrevious_Click()
  If ipageNum > 1 Then
    ipageNum = ipageNum - 1
    cmdNext.Enabled = True
  End If
  If ipageNum = 1 Then
    cmdPrevious.Enabled = False
  End If
  If Not isTallest Then
    GetsevenWonderStat "http://fun.familyeducation.com/slideshow/historic-sites/61628.html?page=" & ipageNum
  Else
    GetTallestBuilding "http://fun.familyeducation.com/slideshow/historic-sites/61490.html?page=" & ipageNum
  End If
End Sub

Private Sub Form_Load()
  Set cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  Set cmdPrevious.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  Set cmdNext.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
  If Not isTallest Then
    ipageCnt = 7
    lblStatInfo.FontSize = 11
    GetsevenWonderStat "http://fun.familyeducation.com/slideshow/historic-sites/61628.html"
  Else
    ipageCnt = 9
    lblStatInfo.FontSize = 9.5
    GetTallestBuilding "http://fun.familyeducation.com/slideshow/historic-sites/61490.html"
  End If
  ipageNum = 1
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub GetWebpage(sWebPage)
  RichTextBox1.Text = ""
  RichTextBox1.Text = Inet1.OpenURL(sWebPage)
End Sub

Private Sub GetsevenWonderStat(sWeblink As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sStartPos As String
  Dim picName As String
  Dim picPath As String
  Dim picUrl As String
  Dim infoText As String
  
  GetWebpage sWeblink
  sStartPos = "BodyText"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "slideH", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=""", vbTextCompare)
  iIndexEnd = InStr(iIndex + 6, RichTextBox1.Text, " ", vbTextCompare)
  
  picUrl = "http://fun.familyeducation.com" & Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 11) - iIndex + 5)
  picName = Mid(picUrl, InStrRev(picUrl, "/") + 1)
  SavePicture picUrl, picName
  picPath = App.Path + "\Icons\" & picName
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<h", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
  lblTitle.Caption = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd - 2) - iIndexSt + 1)
  
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "class=", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</p", vbTextCompare)
  infoText = Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd - 2) - iIndexSt + 1)
  
  lblStatInfo.Caption = Mid(infoText, 1, InStr(1, infoText, "<a", vbTextCompare) - 1)
  
  infoText = Mid(infoText, InStr(1, infoText, "target=", vbTextCompare))
  iIndex = InStr(1, infoText, ">", vbTextCompare)
  iIndexSt = InStr(iIndex, infoText, "</", vbTextCompare)
  lblStatInfo.Caption = lblStatInfo.Caption & Mid(infoText, iIndex + 1, (iIndexSt - 2) - iIndex + 1)
  
  infoText = Mid(infoText, InStr(1, infoText, "</a>", vbTextCompare))
  If InStr(1, infoText, "href=", vbTextCompare) = 0 Then
    lblStatInfo.Caption = lblStatInfo.Caption & Mid(infoText, InStr(1, infoText, "</a>", vbTextCompare))
  Else
    iIndex = InStr(1, infoText, ">", vbTextCompare)
    iIndexSt = InStr(iIndex, infoText, "<a", vbTextCompare)
    lblStatInfo.Caption = lblStatInfo.Caption & Mid(infoText, iIndex + 1, (iIndexSt) - iIndex + 1)
    infoText = Mid(infoText, InStr(1, infoText, "target=", vbTextCompare))
    lblStatInfo.Caption = lblStatInfo.Caption & Mid(infoText, InStr(1, infoText, ">", vbTextCompare))
  End If
  
  lblStatInfo.Caption = Replace(lblStatInfo.Caption, "</a>", "")
  SetPictureBox frmWorldStat, picPath, 0
End Sub

Private Sub SavePicture(myUrl As String, pngFile As String)
  Dim nFileNum As Integer
  Dim myFile As String
  Dim myData() As Byte
  
  myData() = Inet1.OpenURL(myUrl, icByteArray)
 
  nFileNum = FreeFile
  Open App.Path + "\Icons\" & pngFile For Binary Access Write As #nFileNum
    Put #nFileNum, , myData()
  Close #nFileNum
  Erase myData()
End Sub

Private Sub GetTallestBuilding(sWeblink As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sStartPos As String
  Dim picName As String
  Dim picPath As String
  Dim picUrl As String
  
  GetWebpage sWeblink
  sStartPos = "BodyText"
  iIndex = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "slideH", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "src=""", vbTextCompare)
  iIndexEnd = InStr(iIndex + 6, RichTextBox1.Text, " ", vbTextCompare)
  
  picUrl = "http://fun.familyeducation.com" & Mid(RichTextBox1.Text, iIndex + 5, (iIndexEnd - 11) - iIndex + 5)
  picName = Mid(picUrl, InStrRev(picUrl, "/") + 1)
  SavePicture picUrl, picName
  picPath = App.Path + "\Icons\" & picName
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<strong", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</h", vbTextCompare)
  lblTitle.Caption = Replace(Mid(RichTextBox1.Text, iIndexSt + 1, (iIndexEnd - 2) - iIndexSt + 1), "</strong><br>", " ")
  
  iIndex = InStr(iIndexEnd, RichTextBox1.Text, "desc"">", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</p", vbTextCompare)
  lblStatInfo.Caption = Mid(RichTextBox1.Text, iIndexSt + 4, (iIndexEnd - 8) - iIndexSt + 4)
  
  lblStatInfo.Caption = Replace(lblStatInfo.Caption, "</a>", "")
  SetPictureBox frmWorldStat, picPath, 0
End Sub
