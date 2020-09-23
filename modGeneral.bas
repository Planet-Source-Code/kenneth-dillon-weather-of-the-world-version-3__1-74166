Attribute VB_Name = "modGeneral"
Option Explicit
Public sCityTitle As String
Public Arrivalink As String
Public ApArival As Boolean
Public bPicError As Boolean
Public bNoAQIndex As Boolean
Public AQIbgCol(5) As String
Public AQITxt(5) As String
Public AQITitle(5) As String
Public AQIShowTool As Boolean
Public AQICanShowTool As Boolean
Public AQIMonitorShowTool As Boolean
Public AQICityMapArray() As String
Public AQISummeryMap As Boolean
Public AQIPicArray() As String
Public AirPortSummery() As String
Public AirPortUSAState() As String
Public AQICanPicArray() As String
Public bMapView As Boolean
Public sStatusText As String
Public slargeMapLink1 As String
Public slargeMapLink2 As String
Public iApPage As Integer
Public sfndResult As Integer
Public CountriesArray() As String
Public HolDateSelect As String
Public isTallest As Boolean
Public Nozip As Boolean
Public intMH As Integer 'MaxHeight of imagebox
Public intMW As Integer 'MaxWidth of image box
Public OCX() As Byte
Public bGPS As Boolean
Public sStatState As String
Public sStatArea As String
Public sStatCountry As String
Public sStatRegion As String
Public sStatCounty As String
Public PlayRegAnimation As Boolean
Public PlayAnimation As Boolean
Public AnimationLink As String
Public Animation As Boolean
Public sMapPicture As String
Public sFlagPicture As String
Public picTureName As String
Public scntName As String
Public iMinCount As Integer
Public sFrmName As String
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const CB_FINDSTRINGEXACT = &H158
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function InternetCheckConnection Lib "wininet.dll" Alias "InternetCheckConnectionA" (ByVal lpszUrl As String, ByVal dwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, _
        ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
        'Constants
'Const LB_FINDSTRINGEXACT = &H1A2    'To locate exact match

'Declares
Public Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" _
    (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
Public Declare Function SendMessageAsString Lib "user32" Alias "SendMessageA" _
  (ByVal hwnd As Long, ByVal wMsg As Long, _
  ByVal wParam As Long, _
  ByVal lParam As String) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Sub WAIT(ByVal lMilliSec As Long)
    WaitForSingleObject GetCurrentProcess, lMilliSec
End Sub

Public Function FindStringinListControl(ListControl As Object, _
  ByVal SearchText As String) As Long

  '**************************************
  'Input:
  'ListControl: List or ComboBox Object
  'SearchText: String to Search For

  'Returns: ListIndex of Item if found
  'or -1 if not found
  '***************************************
  
  Dim lHwnd As Long
  Dim lMsg As Long

  'On Error Resume Next
  lHwnd = ListControl.hwnd

  If TypeOf ListControl Is ListBox Then
    lMsg = LB_FINDSTRINGEXACT
  ElseIf TypeOf ListControl Is ComboBox Then
    lMsg = CB_FINDSTRINGEXACT
  Else
    FindStringinListControl = -1
    Exit Function
  End If
  FindStringinListControl = SendMessageAsString(lHwnd, lMsg, -1, SearchText)
End Function

Public Function FileExists(FileName As String) As Boolean
  FileExists = (Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive) <> "")
End Function

Public Function SystemDirectory() As String
  Dim RSTR As String
  Dim RLEN As Long

  RSTR = String(255, 0)
  RLEN = GetSystemDirectory(RSTR, Len(RSTR))
  If RLEN < Len(RSTR) Then
    RSTR = Left(RSTR, RLEN)
    If Right(RSTR, 1) = "\" Then
      SystemDirectory = Left(RSTR, Len(RSTR) - 1)
    Else
      SystemDirectory = RSTR
    End If
  Else
    SystemDirectory = ""
  End If
End Function

Public Sub SetPictureBox(NameOfForm As Form, picTureName As String, imgIndex As Integer)
  intMH = NameOfForm.picMain.ScaleHeight '- 20 'in pixels
  intMW = NameOfForm.picMain.ScaleWidth '- 20 'in pixels
  'send it to the LoadAnImage Sub
  Call LoadAnImage(NameOfForm, picTureName, imgIndex, NameOfForm.picMain, NameOfForm.picHidden, NameOfForm.imgPicture, intMH, intMW)
End Sub

Public Sub LoadAnImage(picFormName As Form, picName As String, PicIndex As Integer, picBoxMain As PictureBox, picBoxHidden As PictureBox, imageBoxDisplay As Image, ByVal MaxHeight As Integer, ByVal maxWidth As Integer)
  On Error GoTo LoadAnImageErr
  Dim HighRatio As Single
  Dim WideRatio As Single
  Dim intMaxPicHeight As Integer
  Dim intMaxPicWidth As Integer
  Dim intActualHeight As Integer
  Dim intActualWidth As Integer
  
  intActualHeight = 0
  intActualWidth = 0
  intMaxPicHeight = 0
  intMaxPicWidth = 0
  HighRatio = 0
  WideRatio = 0
  'Get the picture name to load into the hidden picturebox
  'set up the max dimension variables
  intMaxPicHeight = MaxHeight
  intMaxPicWidth = maxWidth
 
  'Load the picture into the hidden picturebox
  'with its Autosize set to True.
  If Len(Trim(picName)) <> 0 Then
    picBoxHidden.Picture = LoadPicture(picName)
    picBoxMain.Visible = False
  Else
    Set picBoxHidden.Picture = picFormName.ImageList2.ListImages(PicIndex).Picture
  End If
  'Get the pic size in pixels
  intActualHeight = CInt(picBoxHidden.ScaleHeight)
  intActualWidth = CInt(picBoxHidden.ScaleWidth)
  'Form a ratio of original height to width - eg 800 x 600 pixels image
  WideRatio = picBoxHidden.ScaleHeight / picBoxHidden.ScaleWidth '600/800
  HighRatio = picBoxHidden.ScaleWidth / picBoxHidden.ScaleHeight '800/600
  'Make the image box invisible until the image is loaded
  imageBoxDisplay.Visible = False
  'Check for Portrait or Landscape image
  If intActualHeight >= intActualWidth Then
    'must be higher than wide - ie portrait
    'Check for smaller image than max allows
    If intActualHeight <= intMaxPicHeight Then
      imageBoxDisplay.Height = intActualHeight
      imageBoxDisplay.Width = intActualWidth
    Else
      imageBoxDisplay.Height = intMaxPicHeight
      imageBoxDisplay.Width = intMaxPicHeight * HighRatio
    End If
  Else
    'must be wider than high - ie landscape
    If intActualWidth <= intMaxPicWidth Then
      imageBoxDisplay.Width = intActualWidth
      imageBoxDisplay.Height = intActualHeight
    Else
      imageBoxDisplay.Width = intMaxPicWidth
      imageBoxDisplay.Height = intMaxPicWidth * WideRatio
      'again make sure the height is not more than the max allows
      If imageBoxDisplay.Height > intMaxPicHeight Then
        'Resize it
        imageBoxDisplay.Height = intMaxPicHeight
        imageBoxDisplay.Width = intMaxPicHeight * HighRatio
      End If
    End If
  End If
  'Center the image within its container picturebox.
  'Load the graphic into the image control.
  imageBoxDisplay.Picture = picBoxHidden.Picture
  If Len(Trim(picName)) <> 0 Then
    imageBoxDisplay.Left = (picFormName.fmMap.Width / 2) - (imageBoxDisplay.Width / 2)
    imageBoxDisplay.Top = (picFormName.fmMap.Height / 2) - (imageBoxDisplay.Height / 2) + 100
  End If
  'Show the image.
  imageBoxDisplay.Visible = True
ExitLoadAnImage:
  Exit Sub
LoadAnImageErr:
  MsgBox Err.Description
  Resume ExitLoadAnImage
End Sub

Public Sub SizePic(picName As String)
  Dim PicRatio As Single
  Dim BoxWidth As Integer
  Dim BoxHeight As Integer
  
  On Error GoTo errorHandler
  'load first pic in in picSource, get ratio,
  'size imgLgCountry, send size, empty box1
  If Right(picName, 3) = "png" Then
    Dim Token As Long
    Token = InitGDIPlus
    frmCountry.picSource = LoadPictureGDIPlus(picName, , True)
    FreeGDIPlus Token
  Else
    frmCountry.picSource.Picture = LoadPicture(picName)
  End If
  PicRatio = frmCountry.picSource.Width / frmCountry.picSource.Height
  
  If PicRatio > 1.33 Then 'pic is landscape
    BoxWidth = Screen.Width / 2.3
    BoxHeight = (Screen.Width / PicRatio) / 2.3
  End If

  If PicRatio < 1.33 Then
    BoxHeight = Screen.Height / 2.3 'pic is portrait
    BoxWidth = (Screen.Height * PicRatio) / 2.3
  End If

  If PicRatio = 1.33 Then 'pic is square
    BoxHeight = Screen.Height / 2.3
    BoxWidth = Screen.Width / 2.3
  End If
  
  Call ShowPic(BoxWidth, BoxHeight, picName)
  frmCountry.imgLgCountry.Visible = True
  Exit Sub
errorHandler:
  If Err.Number = 481 Then
    MsgBox "Unable to show picture, please try again later", vbInformation, "Weather Of The World"
    bPicError = True
  End If
End Sub

Public Sub ShowPic(BoxWidth As Integer, BoxHeight As Integer, picName As String)
  Dim x As Integer
  
  'empty box2, size box2,load imgLgCountry from  box1
  frmCountry.imgLgCountry.Visible = False
  frmCountry.imgLgCountry.Height = BoxHeight
  frmCountry.imgLgCountry.Width = BoxWidth
  
  If Right(picName, 3) = "png" Then
    Dim Token As Long
    Token = InitGDIPlus
    frmCountry.imgLgCountry = LoadPictureGDIPlus(picName, , True)
    FreeGDIPlus Token
  Else
    frmCountry.imgLgCountry.Picture = LoadPicture(picName)
  End If
  
  frmCountry.picSource.Picture = LoadPicture()
  frmCountry.imgLgCountry.Top = 0
  frmCountry.imgLgCountry.Left = 5
  
  If bMapView Or AQICanShowTool Or AQIMonitorShowTool Then
    frmCountry.Height = frmCountry.imgLgCountry.Height + 2000
  ElseIf Len(sStatusText) = 0 Then
    frmCountry.Height = frmCountry.imgLgCountry.Height + 360
  Else
    frmCountry.Height = frmCountry.imgLgCountry.Height + 3600
  End If
  frmCountry.Width = frmCountry.imgLgCountry.Width + 110
  frmCountry.Top = frmWeatherMain.Top + (frmWeatherMain.Height / 2) - (frmCountry.Height / 2)
  If frmCountry.Top < 0 Then
    frmCountry.Top = frmWeatherMain.Top
  End If
  frmCountry.Left = frmWeatherMain.Left + (frmWeatherMain.Width / 2) - (frmCountry.Width / 2)
  If PlayAnimation Then
    frmCountry.cmdAnimate.Caption = "Play Animation"
    frmCountry.cmdAnimate.Visible = True
  End If
  frmCountry.cmdAnimate.Top = frmCountry.imgLgCountry.Height - 400 ' frmCountry.Height - 850
  frmCountry.cmdAnimate.Left = (frmCountry.Width / 2) - (frmCountry.cmdAnimate.Width / 2)
  If bMapView Then
    For x = 0 To UBound(AQIPicArray, 1)
      frmCountry.cmdAQITabs(x).Visible = True
      Set frmCountry.cmdAQITabs(x).MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
      If x < 3 Then
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 400
      Else
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 1000
      End If
    Next
    For x = 0 To 5
      frmCountry.AQIPic(x).Visible = True
      If x < 3 Then
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 320
      Else
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 900
      End If
    Next
    Set frmCountry.cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    frmCountry.cmdExit.Visible = True
    frmCountry.cmdExit.Top = frmCountry.imgLgCountry.Height + 1000
  End If
  If AQICanShowTool Then
    For x = 0 To UBound(AQICanPicArray, 1)
      frmCountry.cmdAQITabs(x).Visible = True
      Set frmCountry.cmdAQITabs(x).MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
      If x < 3 Then
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 400
      Else
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 1000
      End If
    Next
    For x = 0 To 5
      frmCountry.AQIPic(x).Visible = True
      If x < 3 Then
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 320
      Else
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 900
      End If
    Next
    Set frmCountry.cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    frmCountry.cmdExit.Visible = True
    frmCountry.cmdExit.Top = frmCountry.imgLgCountry.Height + 1000
    frmCountry.cmdAQITabs(0).Caption = "Current AQI"
    frmCountry.cmdAQITabs(1).Caption = "AQI Animation"
  End If
  
  If AQIMonitorShowTool Then
    For x = 0 To UBound(AQICanPicArray, 1)
      frmCountry.cmdAQITabs(x).Visible = True
      Set frmCountry.cmdAQITabs(x).MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
      If x < 3 Then
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 400
      Else
        frmCountry.cmdAQITabs(x).Top = frmCountry.imgLgCountry.Height + 1000
      End If
    Next
    For x = 0 To 5
      frmCountry.AQIPic(x).Visible = True
      If x < 3 Then
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 320
      Else
        frmCountry.AQIPic(x).Top = frmCountry.imgLgCountry.Height + 900
      End If
    Next
    Set frmCountry.cmdExit.MouseIcon = frmWeatherMain.ImageList1.ListImages(3).Picture
    frmCountry.cmdExit.Visible = True
    frmCountry.cmdAQITabs(0).Caption = "Current Ozone"
    frmCountry.cmdAQITabs(1).Caption = "Current PM2.5"
    frmCountry.cmdAQITabs(2).Caption = "Ozone Animate"
    frmCountry.cmdAQITabs(3).Caption = "PM2.5 Aninate"
    frmCountry.cmdAQITabs(4).Caption = "Peak Ozone"
    frmCountry.cmdAQITabs(5).Caption = "Peak PM2.5"
    frmCountry.cmdAQITabs(3).Left = 120
    frmCountry.cmdAQITabs(4).Left = 1480
    frmCountry.cmdAQITabs(5).Left = 2840
    frmCountry.cmdExit.Left = (frmCountry.Width / 2 - frmCountry.cmdExit.Width / 2)
    frmCountry.cmdExit.Top = frmCountry.imgLgCountry.Height - 375
    frmCountry.Caption = sFrmName
  End If
  If Len(sStatusText) = 0 Then Exit Sub
  frmCountry.fraInfo.Width = frmCountry.Width - 340
  frmCountry.fraInfo.Height = (frmCountry.Height - frmCountry.imgLgCountry.Height) - 500
  frmCountry.fraInfo.Top = frmCountry.imgLgCountry.Height + 50
  frmCountry.rchTxtInfo.Width = frmCountry.Width - 600
  frmCountry.rchTxtInfo.Height = (frmCountry.Height - frmCountry.imgLgCountry.Height) - 950
End Sub

Public Sub LoadAQITips()
  AQIbgCol(0) = 8453888
  AQIbgCol(1) = 8454143
  AQIbgCol(2) = 4227327
  AQIbgCol(3) = 255
  AQIbgCol(4) = 6370475
  AQIbgCol(5) = 1911674
  AQITitle(0) = "AQI - Good (0 - 50)"
  AQITitle(1) = "AQI - Moderate (51 - 100)"
  AQITitle(2) = "AQI - Unhealthy for Sensitive Groups (101 - 150)"
  AQITitle(3) = "AQI - Unhealthy (151 - 200)"
  AQITitle(4) = "AQI - Very Unhealthy (201 - 300)"
  AQITitle(5) = "AQI - Hazardous (301 - 500)"
  
  AQITxt(0) = "Air quality is considered satisfactory, and air pollution poses little or no risk."
  AQITxt(1) = "Air quality is acceptable; however, for some pollutants there may be a moderate" & vbLf & _
              "health concern for a very small number of people. For example, people who are unusually" & vbLf & _
              "sensitive to ozone may experience respiratory symptoms."
  AQITxt(2) = "Although general public is not likely to be affected at this AQI range, people with lung disease, " & vbLf & _
              "older adults and children are at a greater risk from exposure to ozone, whereas persons with heart and lung" & vbLf & _
              "disease, older adults and children are at greater risk from the presence of particles in the air."
  AQITxt(3) = "Everyone may begin to experience some adverse health effects, and members of the" & vbLf & _
              "sensitive groups may experience more serious effects."
  AQITxt(4) = "This would trigger a health alert signifying that everyone" & vbLf & _
              "may experience more serious health effects."
  AQITxt(5) = "This would trigger a health warnings of emergency conditions." & vbLf & _
              "The entire population is more likely to be affected."
End Sub
