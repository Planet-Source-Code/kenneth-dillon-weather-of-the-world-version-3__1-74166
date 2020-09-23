VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAportStatus 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   12195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdStateAir 
      Caption         =   "State Airport"
      Height          =   375
      Left            =   2400
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9400
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdState 
      Caption         =   "State Info"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9400
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSComctlLib.ListView lstAirUSAState 
      Height          =   8275
      Left            =   120
      TabIndex        =   11
      Top             =   525
      Visible         =   0   'False
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   14605
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
         Object.Width           =   8538
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   8538
      EndProperty
   End
   Begin MSComctlLib.ListView lstAirportCountry 
      Height          =   8275
      Left            =   120
      TabIndex        =   10
      Top             =   525
      Visible         =   0   'False
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   14605
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
         Object.Width           =   7480
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   6774
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   3704
      EndProperty
   End
   Begin VB.CommandButton cmdAirLine 
      Caption         =   "Airline Info"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   9000
      Visible         =   0   'False
      Width           =   1155
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   10200
      Top             =   9720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   9840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5106
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmAportStatus.frx":0000
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   9000
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   9000
      Visible         =   0   'False
      Width           =   1155
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3715
      Left            =   120
      TabIndex        =   2
      Top             =   5020
      Width           =   11895
      ExtentX         =   20981
      ExtentY         =   6553
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
      Location        =   "http:///"
   End
   Begin MSComctlLib.ListView lstAirportStat 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   15875
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5380
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   9000
      Width           =   1155
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
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
      Left            =   180
      TabIndex        =   8
      Top             =   9075
      Width           =   75
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   8960
      Visible         =   0   'False
      Width           =   4600
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   80
      Width           =   11775
   End
End
Attribute VB_Name = "frmAportStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sfndpage As Integer

Private Sub cmdAirLine_Click()
  If AQIShowTool Then
    If lstAirportCountry.Visible Then
      GetAirPortInformation AirPortSummery(lstAirportCountry.SelectedItem.Index - 1)
      lstAirportCountry.Visible = False
      cmdAirLine.Caption = "View Airport"
      lblInfo.Caption = lstAirportCountry.ListItems(lstAirportCountry.SelectedItem.Index).Text & vbCrLf & "GPS Location Airport Map"
      lblInfo.Visible = True
      lblCount.Visible = False
    Else
      lstAirportCountry.Visible = True
      cmdAirLine.Caption = "Airport Info"
      lblTitle.Caption = sCityTitle
      lblCount.Visible = True
      lblInfo.Visible = False
    End If
  Else
    GetAirlineInfo AirPortSummery(lstAirportStat.SelectedItem.Index)
    lstAirportStat.SetFocus
  End If
End Sub

Private Sub cmdExit_Click()
  Unload Me
End Sub

Private Sub cmdNext_Click()
  sfndpage = sfndpage + 1
  If sfndpage >= iApPage Then
    cmdNext.Enabled = False
  End If
  lblCount.Caption = sfndpage & "/" & iApPage
  cmdPrevious.Enabled = True
  GetArrivalPage Replace(Arrivalink, ".html", "/" & sfndpage & ".html")
End Sub

Private Sub cmdPrevious_Click()
  If sfndpage >= 2 Then
    sfndpage = sfndpage - 1
  End If
  If sfndpage = 1 Then
    cmdPrevious.Enabled = False
  End If
  lblCount.Caption = sfndpage & "/" & iApPage
  cmdNext.Enabled = True
  GetArrivalPage Replace(Arrivalink, ".html", "/" & sfndpage & ".html")
End Sub

Private Sub cmdState_Click()
  GetUSAStateInfo AirPortUSAState(lstAirUSAState.SelectedItem.Index - 1)
  cmdStateAir.Left = 3720
  cmdState.Visible = False
  Set cmdStateAir.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
End Sub

Private Sub cmdStateAir_Click()
  If cmdStateAir.Caption = "Airport Info" Then
    cmdStateAir.Caption = "State Airport"
    lblInfo.Visible = False
    lstAirportCountry.Visible = True
    lblTitle.Caption = sCityTitle
    lblCount.Visible = True
  Else
    GetUSAAirport AirPortSummery(lstAirportCountry.SelectedItem.Index - 1)
    lstAirportCountry.Visible = False
    cmdState.Enabled = True
    lblCount.Visible = False
    lblInfo.Caption = lstAirportCountry.ListItems(lstAirportCountry.SelectedItem.Index).Text & vbCrLf & "GPS Location Airport Map"
    lblInfo.Visible = True
  End If
End Sub

Private Sub Form_Load()
  frmAportStatus.Height = 9975
  If ApArival Then
    sfndpage = 1
    cmdNext.Enabled = True
    Set cmdPrevious.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    Set cmdNext.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    Set cmdAirLine.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    sCityTitle = lblTitle.Caption
  Else
    Set cmdAirLine.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    WebBrowser1.Navigate AnimationLink
  End If
  Set cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  iApPage = 0
  AQIShowTool = False
  Set frmAportStatus = Nothing
End Sub

Private Sub GetArrivalPage(sUrl As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sNum As Integer
  Dim sStringToParse As String
  Dim sCnt, x As Integer
  Dim NameArray() As String
  Dim sInfo As String
  
  MousePointer = 11
  lstAirportStat.ListItems.Clear
  sPageName = "http://www.airwise.com" & sUrl
  GetWebpage sPageName
  sStartPos = "name=""content"""
  
  iIndex = InStr(1, RichTextBox1.Text, "<table", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</table", vbTextCompare)
  sStringToParse = Mid(RichTextBox1.Text, iIndex, (iIndexEnd) - (iIndex))
  
  NameArray() = Split(sStringToParse, "valign=")
  iIndex = 1
  For sNum = 1 To UBound(NameArray, 1)
    For x = 0 To 3
      iIndexSt = InStr(iIndex, NameArray(sNum), "<td>", vbTextCompare)
      iIndex = InStr(iIndexSt, NameArray(sNum), "</", vbTextCompare)
      If InStr(1, Mid(NameArray(sNum), iIndexSt + 4, (iIndex) - (iIndexSt + 4)), "href=", vbTextCompare) <> 0 Then
        iIndexEnd = InStr(1, NameArray(sNum), "href=", vbTextCompare)
        iIndexSt = InStr(iIndexEnd, NameArray(sNum), ">", vbTextCompare)
        ReDim Preserve AirPortSummery(sCnt)
        AirPortSummery(sCnt) = Mid(NameArray(sNum), iIndexEnd + 6, (iIndexSt - 1) - (iIndexEnd + 6))
        iIndex = InStr(iIndexSt, NameArray(sNum), "</", vbTextCompare)
        lstAirportStat.ListItems(sCnt).ListSubItems.Add , , Mid(NameArray(sNum), iIndexSt + 1, (iIndex) - (iIndexSt + 1))
      Else
        iIndex = InStr(iIndexSt, NameArray(sNum), "</", vbTextCompare)
        If x = 0 Then
          lstAirportStat.ListItems.Add , , Mid(NameArray(sNum), iIndexSt + 4, (iIndex) - (iIndexSt + 4))
          sCnt = sCnt + 1
        Else
          lstAirportStat.ListItems(sCnt).ListSubItems.Add , , Mid(NameArray(sNum), iIndexSt + 4, (iIndex) - (iIndexSt + 4))
        End If
      End If
    Next
    For x = 0 To 1
      iIndexEnd = InStr(iIndex, NameArray(sNum), "class=", vbTextCompare)
      iIndexSt = InStr(iIndexEnd, NameArray(sNum), ">", vbTextCompare)
      iIndex = InStr(iIndexSt, NameArray(sNum), "</", vbTextCompare)
      lstAirportStat.ListItems(sCnt).ListSubItems.Add , , Replace(Mid(NameArray(sNum), iIndexSt + 1, (iIndex) - (iIndexSt + 1)), "<br />", "")
    Next
    iIndex = 1
  Next
  iIndexEnd = InStr(1, RichTextBox1.Text, "col3_content", vbTextCompare)
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<p>", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  sInfo = Mid(RichTextBox1.Text, iIndexSt + 3, (iIndex) - (iIndexSt + 3))
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "<p>", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  If InStr(1, Mid(RichTextBox1.Text, iIndexSt + 3, (iIndex) - (iIndexSt + 3)), "temperature", vbTextCompare) <> 0 Then
    sInfo = sInfo & vbCrLf & Mid(RichTextBox1.Text, iIndexSt + 3, (iIndex) - (iIndexSt + 3))
  Else
    sInfo = sInfo & vbCrLf & "Temperature Unknown"
  End If
  lblInfo = Replace(sInfo, "&deg;", Chr(176))
  MousePointer = 0
  Erase NameArray
End Sub

Private Sub GetWebpage(sWebPage)
  RichTextBox1.Text = ""
  RichTextBox1.Text = Inet1.OpenURL(sWebPage)
End Sub

Private Sub lstAirportCountry_Click()
  cmdAirLine.Enabled = True
End Sub

Private Sub lstAirportCountry_DblClick()
  If cmdAirLine.Visible And cmdStateAir.Visible = False Then
    cmdAirLine_Click
  Else
    cmdStateAir_Click
  End If
End Sub

Private Sub lstAirportStat_Click()
  cmdAirLine.Enabled = True
End Sub

Private Sub GetAirlineInfo(sUrl As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex, iIndex1 As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sNum As Integer
  Dim sStringToParse As String
  Dim sTxtInfo1, sTxtInfo2 As String
  
  MousePointer = 11
  
  sPageName = "http://www.airwise.com" & sUrl
  GetWebpage sPageName
  sStartPos = "name=""content"""
  
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  If iIndexSt = 0 Then
    MsgBox "No Website Available for " & lstAirportStat.SelectedItem.ListSubItems(2).Text, vbInformation, "Weather Of The World"
    MousePointer = 0
    Exit Sub
  End If
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "<h1>", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
  frmAirInfo.lblInfo.Caption = Mid(RichTextBox1.Text, iIndex + 4, (iIndexEnd) - (iIndex + 4))
  
  Do
    iIndex = InStr(iIndexEnd, RichTextBox1.Text, "<div class=", vbTextCompare)
    If iIndex = 0 Then
      MsgBox "No Website Available for " & lstAirportStat.SelectedItem.ListSubItems(2).Text, vbInformation, "Weather Of The World"
      MousePointer = 0
      Unload frmAirInfo
      Exit Sub
    End If
    iIndex1 = InStr(iIndex, RichTextBox1.Text, "</div", vbTextCompare)
    sStringToParse = Mid(RichTextBox1.Text, iIndex, (iIndex1) - (iIndex))
    If InStr(1, sStringToParse, "href=", vbTextCompare) <> 0 Then
      iIndex = InStr(1, sStringToParse, ">", vbTextCompare)
      iIndexSt = InStr(iIndex, sStringToParse, "<", vbTextCompare)
      If sTxtInfo1 = "" Then
        sTxtInfo1 = Mid(sStringToParse, iIndex + 2, (iIndexSt) - (iIndex + 2))
      Else
        sTxtInfo1 = sTxtInfo1 & vbCrLf & Mid(sStringToParse, iIndex + 1, (iIndexSt) - (iIndex + 1))
      End If
      iIndex = InStr(iIndexSt, sStringToParse, ">", vbTextCompare)
      iIndexSt = InStr(iIndex, sStringToParse, "</", vbTextCompare)
      sTxtInfo1 = sTxtInfo1 & " " & Mid(sStringToParse, iIndex + 1, (iIndexSt) - (iIndex + 1))
    Else
      iIndex = InStr(1, sStringToParse, ">", vbTextCompare)
      iIndexSt = InStr(iIndex, sStringToParse, "<", vbTextCompare)
      If sNum <= 5 And InStr(1, sStringToParse, "Address:") = 0 Then
        sTxtInfo1 = sTxtInfo1 & vbCrLf & Replace(Mid(sStringToParse, iIndex + 1), "<br />", "")
      Else
        If sTxtInfo2 = "" Then
          If InStr(1, sStringToParse, "Address:") = 0 Then
            sTxtInfo2 = Replace(Mid(sStringToParse, iIndex + 1), "<br />", "")
          Else
            sTxtInfo2 = Replace(Mid(sStringToParse, iIndex + 2), "<br />", "")
          End If
        Else
          sTxtInfo2 = sTxtInfo2 & vbCrLf & Replace(Mid(sStringToParse, iIndex + 2), "<br />", "")
        End If
      End If
    End If
    
    If InStr(1, Mid(RichTextBox1.Text, iIndex1, 50), "</table>", vbTextCompare) <> 0 Then
      Exit Do
    End If
    iIndexEnd = iIndex1
    sNum = sNum + 1
  Loop
  frmAirInfo.rchTxt1.Text = Replace(sTxtInfo1, Chr(10) & Chr(10), vbCrLf)
  frmAirInfo.rchTxt2.Text = sTxtInfo2
  MousePointer = 0
  frmAirInfo.Show vbModal
End Sub

Private Sub lstAirportStat_DblClick()
  cmdAirLine_Click
End Sub

Private Sub GetAirPortInformation(sUrl As String)
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim x, sCnt As Integer
  Dim sTitle, sPageName As String
  Dim sLatitude, sLongitude As String
  Dim sStringToParse As String
  Dim NameArray() As String
  
  'On Error Resume Next
  MousePointer = 11
  lstAirportStat.ListItems.Clear
  sPageName = "http://worldaerodata.com" & sUrl
  GetWebpage sPageName
  
  iIndex = InStr(1, RichTextBox1.Text, "<h3>", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
  sTitle = Mid(RichTextBox1.Text, iIndex + 4, (iIndexEnd) - (iIndex + 4)) & " General Info"
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "href=", vbTextCompare)
  iIndex = InStr(iIndexSt, RichTextBox1.Text, ",", vbTextCompare)
  sLatitude = Mid(RichTextBox1.Text, iIndexSt + 8, (iIndex - 1) - (iIndexSt + 8))
  iIndexSt = InStr(iIndex, RichTextBox1.Text, ">", vbTextCompare)
  sLongitude = Mid(RichTextBox1.Text, iIndex + 1, (iIndexSt - 1) - (iIndex + 1))
  AnimationLink = "http://www.mappingsupport.com/p/gmap4.php?ll=" & sLatitude & "," & sLongitude & "&z=10&t=m&icon=pgs"
  
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "<tr>", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</table", vbTextCompare)
  sStringToParse = Mid(RichTextBox1.Text, iIndex, (iIndexEnd) - (iIndex))
  
  NameArray() = Split(sStringToParse, "bgcolor=")
  For x = 1 To UBound(NameArray, 1)
    NameArray(x) = Replace(NameArray(x), "<BR>", "   /   ")
    iIndex = InStr(1, NameArray(x), ">", vbTextCompare)
    If x Mod 2 = 0 Then
      If x = UBound(NameArray, 1) Then
       lstAirportStat.ListItems(sCnt).ListSubItems.Add , , StrConv(Mid(NameArray(x), iIndex + 1), vbProperCase)
      Else
        iIndexSt = InStr(iIndex, NameArray(x), "<", vbTextCompare)
        If Len(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1))) > 7 Then
          lstAirportStat.ListItems(sCnt).ListSubItems.Add , , StrConv(Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&deg;", Chr(176)), vbProperCase)
        Else
          lstAirportStat.ListItems(sCnt).ListSubItems.Add , , Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&deg;", Chr(176))
        End If
      End If
    Else
      iIndexSt = InStr(iIndex, NameArray(x), "</", vbTextCompare)
      lstAirportStat.ListItems.Add , , Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1))
      sCnt = sCnt + 1
    End If
  Next
  WebBrowser1.Navigate2 AnimationLink
  frmAportStatus.lblTitle = StrConv(sTitle, vbProperCase)
  MousePointer = 0
  Erase NameArray
End Sub

Private Sub GetUSAStateInfo(sUrl As String)
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim iIndex As Long
  Dim sPageName As String
  Dim sStartPos As String
  Dim sNum As Integer
  Dim sStringToParse As String
  Dim NameArray() As String
  Dim x, sCnt As Integer
  
  MousePointer = 11
  sPageName = "http://worldaerodata.com/US/" & sUrl
 
  GetWebpage sPageName
  sStartPos = "<h2>"
  iApPage = 0
  iIndexSt = InStr(1, RichTextBox1.Text, sStartPos, vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</", vbTextCompare)
  frmAportStatus.lblTitle.Caption = StrConv(Mid(RichTextBox1.Text, iIndexSt + 4, (iIndexEnd) - (iIndexSt + 4)), vbProperCase)
  sCityTitle = StrConv(Mid(RichTextBox1.Text, iIndexSt + 4, (iIndexEnd) - (iIndexSt + 4)), vbProperCase)
  frmAportStatus.lstAirportCountry.ColumnHeaders.Add 5, , , 2000
  frmAportStatus.lstAirportCountry.ColumnHeaders.Add 6, , , 1000
  frmAportStatus.lstAirportCountry.ColumnHeaders(1).Width = 3700
  frmAportStatus.lstAirportCountry.ColumnHeaders(2).Width = 1000
  frmAportStatus.lstAirportCountry.ColumnHeaders(3).Width = 1000
  frmAportStatus.lstAirportCountry.ColumnHeaders(4).Width = 2800
  
  For x = 1 To 6
    iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<b>", vbTextCompare)
    iIndex = InStr(iIndexSt, RichTextBox1.Text, ">", vbTextCompare)
    iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
    frmAportStatus.lstAirportCountry.ColumnHeaders(x).Text = Mid(RichTextBox1.Text, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  Next
 
  iIndexSt = InStr(iIndexEnd, RichTextBox1.Text, "<tr>", vbTextCompare)
  iIndexEnd = InStr(iIndexSt, RichTextBox1.Text, "</table", vbTextCompare)
  sStringToParse = Mid(RichTextBox1.Text, iIndex, (iIndexEnd) - (iIndex))
  NameArray() = Split(sStringToParse, "width=")
  
  For x = 1 To UBound(NameArray, 1)
    iIndex = InStr(1, NameArray(x), ">", vbTextCompare)
    iIndexEnd = InStr(iIndex, NameArray(x), "</", vbTextCompare)
    If InStr(1, NameArray(x), "href=", vbTextCompare) <> 0 Then
      iIndexSt = InStr(iIndex, NameArray(x), "..", vbTextCompare)
      iIndexEnd = InStr(iIndexSt, NameArray(x), ">", vbTextCompare)
      ReDim Preserve AirPortSummery(sCnt)
      AirPortSummery(sCnt) = Mid(NameArray(x), iIndexSt + 2, (iIndexEnd - 1) - (iIndexSt + 2))
      'City Name
      iIndexSt = InStr(iIndexEnd, NameArray(x), "</", vbTextCompare)
      frmAportStatus.lstAirportCountry.ListItems.Add = StrConv(Mid(NameArray(x), iIndexEnd + 1, (iIndexSt) - (iIndexEnd + 1)), vbProperCase)
      sCnt = sCnt + 1
      sNum = 0
    Else
      sNum = sNum + 1
      iIndex = InStr(1, NameArray(x), ">", vbTextCompare)
      iIndexSt = InStr(iIndex, NameArray(x), "</", vbTextCompare)
      If sNum = 3 Or sNum = 4 Then
        frmAportStatus.lstAirportCountry.ListItems(sCnt).ListSubItems.Add , , StrConv(Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&nbsp;", ""), vbProperCase)
      Else
        frmAportStatus.lstAirportCountry.ListItems(sCnt).ListSubItems.Add , , Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&nbsp;", "")
      End If
    End If
  Next
  Erase NameArray
  AQIShowTool = True
  frmAportStatus.lstAirUSAState.Visible = False
  frmAportStatus.cmdAirLine.Visible = True
  frmAportStatus.lstAirportCountry.Visible = True
  frmAportStatus.lblCount.Caption = sCnt & " Airport"
  cmdStateAir.Top = 9000
  cmdStateAir.Visible = True
  cmdState.Enabled = False
  MousePointer = 0
End Sub

Private Sub lstAirUSAState_Click()
  cmdState.Enabled = True
End Sub

Private Sub GetUSAAirport(sUrl As String)
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim x, sCnt As Integer
  Dim sTitle, sPageName As String
  Dim sLatitude, sLongitude As String
  Dim sStringToParse As String
  Dim NameArray() As String
  Dim sConvLong As Boolean
  Dim sConvLat As Boolean
  
  'On Error Resume Next
  MousePointer = 11
  lstAirportStat.ListItems.Clear
  sPageName = "http://worldaerodata.com" & sUrl
  GetWebpage sPageName
  sLongitude = ""
  sLatitude = ""
  iIndex = InStr(1, RichTextBox1.Text, "<h3>", vbTextCompare)
  iIndexSt = InStr(iIndex, RichTextBox1.Text, "</", vbTextCompare)
  sTitle = Mid(RichTextBox1.Text, iIndex + 4, (iIndexSt) - (iIndex + 4)) & " General Info"
  
  iIndex = InStr(iIndexSt, RichTextBox1.Text, "<tr>", vbTextCompare)
  iIndexEnd = InStr(iIndex, RichTextBox1.Text, "</table", vbTextCompare)
  sStringToParse = Mid(RichTextBox1.Text, iIndex, (iIndexEnd) - (iIndex))
  NameArray() = Split(sStringToParse, "bgcolor=")
  For x = 1 To UBound(NameArray, 1)
    NameArray(x) = Replace(Replace(NameArray(x), Chr(10), ""), "<BR>", "  /  ")
    iIndex = InStr(1, NameArray(x), ">", vbTextCompare)
    If x Mod 2 = 0 Then
      If x = UBound(NameArray, 1) Then
       lstAirportStat.ListItems(sCnt).ListSubItems.Add , , StrConv(Mid(NameArray(x), iIndex + 1), vbProperCase)
      Else
        iIndexSt = InStr(iIndex, NameArray(x), "<", vbTextCompare)
        If x = 4 Then
          lstAirportStat.ListItems(sCnt).ListSubItems.Add , , StrConv(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), vbProperCase)
        Else
          lstAirportStat.ListItems(sCnt).ListSubItems.Add , , Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&deg;", Chr(176))
        End If
        If sConvLong = True Then
          If InStr(2, Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&deg;", Chr(176)), "-", vbTextCompare) <> 0 And _
            InStr(1, NameArray(x), "  /  ", vbTextCompare) = 0 Then
            sLongitude = ConvertDegree(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)))
          Else
            If InStr(1, NameArray(x), "  /  ", vbTextCompare) <> 0 Then
              sLongitude = Mid(NameArray(x), iIndex + 1, InStr(1, NameArray(x), "  /", vbTextCompare) - 10)
            Else
              sLongitude = Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1))
            End If
            
          End If
          sConvLong = False
        End If
        If sConvLat = True Then
          If InStr(2, Replace(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "&deg;", Chr(176)), "-", vbTextCompare) <> 0 And _
            InStr(1, NameArray(x), "  /  ", vbTextCompare) = 0 Then
            sLatitude = ConvertDegree(Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)))
          Else
            If InStr(1, NameArray(x), "  /  ", vbTextCompare) <> 0 Then
              sLatitude = Mid(NameArray(x), iIndex + 1, InStr(1, NameArray(x), "  /", vbTextCompare) - 10)
            Else
              sLatitude = Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1))
            End If
          End If
          sConvLat = False
        End If
      End If
    Else
      iIndexSt = InStr(iIndex, NameArray(x), "</", vbTextCompare)
      lstAirportStat.ListItems.Add , , Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1))
      If InStr(1, Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "Latitude", vbTextCompare) <> 0 Then
        sConvLat = True
      ElseIf InStr(1, Mid(NameArray(x), iIndex + 1, (iIndexSt) - (iIndex + 1)), "Longitude", vbTextCompare) <> 0 Then
        sConvLong = True
      End If
      sCnt = sCnt + 1
    End If
  Next
  AnimationLink = "http://www.mappingsupport.com/p/gmap4.php?ll=" & sLatitude & "," & sLongitude & "&z=10&t=m&icon=pgs"
  WebBrowser1.Navigate2 AnimationLink
  frmAportStatus.lblTitle = StrConv(sTitle, vbProperCase)
  frmAportStatus.lstAirportStat.ColumnHeaders(2).Width = 7000
  cmdStateAir.Caption = "Airport Info"
  Erase NameArray
  MousePointer = 0
End Sub

Private Function ConvertDegree(sItemToCon As String) As String
  Dim iMin As Single
  Dim iSec As Single
  Dim iDeg As Single
  Dim iIndex As Long
  Dim iIndexEnd As Long
  Dim iIndexSt As Long
  Dim sNewDeg As String
  Dim bNev As Boolean
  
  bNev = False
  sItemToCon = Replace(sItemToCon, Chr(10), "")
  sNewDeg = Mid(sItemToCon, 1, Len(sItemToCon) - 1)
  iIndex = InStr(1, sItemToCon, "-", vbTextCompare)
  iDeg = Mid(sNewDeg, 1, iIndex - 1)
  iIndexEnd = InStr(iIndex + 1, sNewDeg, "-", vbTextCompare)
  iMin = Mid(sNewDeg, iIndex + 1, (iIndexEnd) - (iIndex + 1))
  iSec = Mid(sNewDeg, (iIndexEnd + 1))
  If Right(Trim(sItemToCon), 1) = "W" Or Right(sItemToCon, 1) = "S" Then
    bNev = True
  End If
  If bNev Then
    ConvertDegree = "-" & iDeg + (iMin / 60) + (iSec / 3600)
  Else
    ConvertDegree = iDeg + (iMin / 60) + (iSec / 3600)
  End If
End Function

Private Sub lstAirUSAState_DblClick()
  cmdState_Click
End Sub
