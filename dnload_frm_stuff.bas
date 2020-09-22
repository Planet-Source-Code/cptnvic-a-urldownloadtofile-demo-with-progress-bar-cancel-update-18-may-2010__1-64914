Attribute VB_Name = "Module1"
Option Explicit

'--> declare the api to draw the form border
Public Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'--> declare api to round the form and clip for transparency
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, ByVal RectY1 As Long, ByVal RectX2 As Long, ByVal RectY2 As Long, ByVal EllipseWidth As Long, ByVal EllipseHeight As Long) As Long

'--> declare api for draging the form around (mousedown anywhere on form ... except controls)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'declare the SystemParametersInfo api
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_GETWORKAREA = 48 'Desktop Area with task bar consideration.
'Type structure to hold results of query
Public Type User_Vis_Screen_Rect
    Left As Long
    Top As Long
    Right As Long   'Width = Right - Left
    Bottom As Long  'Height = Bottom - Top
End Type
'assign the type
Public ScreenDimensions As User_Vis_Screen_Rect 'used to keep the actual screen size results (in pixels)
Public GetScreenData As Long ' API call requires a long number
Public StoreDimensions As User_Vis_Screen_Rect  'A good place to store results for later work

Public Sub DrawGradient(Thing As Object, R As Integer, G As Integer, B As Integer, Top2Bot As Boolean, X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    'this sub draws the gradient on the 'download' form
    'it is modified from my msgbox replacement submission
    Dim RedStrt, GrnStrt, BluStrt, Red, Grn, Blu, UseRed, UseGreen, UseBlue, Trips, Dist, Y
    
    If Top2Bot Then
        RedStrt = 255: GrnStrt = 255: BluStrt = 255
        Red = R: Grn = G: Blu = B
    Else
        'swap values
        RedStrt = R: GrnStrt = G: BluStrt = B
        Red = 255: Grn = 255: Blu = 255
    End If
    Trips = Y2 - Y1
    If Trips < 1 Then Exit Sub 'prevent error... skip the gradient
    Dist = (Y2 - Y1) / 255
    For Y = 0 To Trips
        UseRed = (RedStrt / 255) * (255 - (Y / Dist)) + (Red / 255) * (Y / Dist)
        UseGreen = (GrnStrt / 255) * (255 - (Y / Dist)) + (Grn / 255) * (Y / Dist)
        UseBlue = (BluStrt / 255) * (255 - (Y / Dist)) + (Blu / 255) * (Y / Dist)
        Thing.Line (X1, Y1 + Y)-(X2, Y1 + Y), RGB(UseRed, UseGreen, UseBlue)
    Next
End Sub
Public Sub GetScreenInfo()
    'this sub gets the useable screen info for locating the download form
    'Call the SystemParametersInfo API
    GetScreenData = SystemParametersInfo(SPI_GETWORKAREA, vbNull, StoreDimensions, 0)
    'store results
    If GetScreenData Then
        'the API call was successful... returns dimensions in pixel terms
        ScreenDimensions.Left = StoreDimensions.Left
        ScreenDimensions.Right = StoreDimensions.Right
        ScreenDimensions.Top = StoreDimensions.Top
        ScreenDimensions.Bottom = StoreDimensions.Bottom
        'note: on my 800x600 monitor w/ single height taskbar at bottom...
        'ScreenDimensions.Left = 0
        'ScreenDimensions.Right = 800
        'ScreenDimensions.Top = 0
        'ScreenDimensions.Bottom = 572
        'therefore:
        'Total Available Width = ScreenDimensions.Right - ScreenDimensions.Left
        'Total Available Height = ScreenDimensions.Bottom - ScreenDimensions.Top
    Else
        'API call failed
        'try less sophisticated way
        ScreenDimensions.Left = 0
        ScreenDimensions.Right = Int(Screen.Width / Screen.TwipsPerPixelX) 'total screen width in pixels
        ScreenDimensions.Top = 0
        ScreenDimensions.Bottom = Int(Screen.Height / Screen.TwipsPerPixelY)
    End If
End Sub
