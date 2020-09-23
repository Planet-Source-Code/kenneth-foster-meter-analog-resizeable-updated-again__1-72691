VERSION 5.00
Begin VB.UserControl MeterGDIPlus 
   ClientHeight    =   1695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   ScaleHeight     =   113
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   0
      ScaleHeight     =   56
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   0
      Top             =   0
      Width           =   1785
   End
End
Attribute VB_Name = "MeterGDIPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'NOTE: May cause crash in IDE, so save work frequently,
'But okay once compiled in exe

'Ken Foster Nov 2009
'original code by Fernando Macedo (Analog Meter GDIP),but now heavily modified
Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Const GdiPlusVersion As Long = 1
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private GdipToken As Long

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As SmoothingMode) As Long

' Quality mode constants
Private Enum QualityMode
   QualityModeInvalid = -1
   QualityModeDefault = 0
   QualityModeLow = 1       ' Best performance
   QualityModeHigh = 2       ' Best rendering quality
End Enum

Private Enum SmoothingMode
   SmoothingModeInvalid = -1
   SmoothingModeDefault = 0
   SmoothingModeHighSpeed = 1
   SmoothingModeHighQuality = 2
   SmoothingModeNone = 3
   SmoothingModeAntiAlias = 4
End Enum

Private Enum GpUnit
   UnitWorld = 0
   ' World coordinate (non-physical unit)
   UnitDisplay = 1
   ' Variable -- for PageTransform only
   UnitPixel = 2
   ' Each unit is one device pixel.
   UnitPoint = 3
   ' Each unit is a printer's point, or 1/72 inch.
   UnitInch = 4
   ' Each unit is 1 inch.
   UnitDocup1nt = 5
   ' Each unit is 1/300 inch.
   UnitMillip1ter = 6
   ' Each unit is 1 millip1ter.
End Enum

Public Enum eMax
  [100] = 1
  [1000] = 2
End Enum


Public Enum Colors
   Aqua = &HFF00FFFF
   Black = &HFF000000
   Blue = &HFF0000FF
   BlueViolet = &HFF8A2BE2
   Brown = &HFFA52A2A
   Chocolate = &HFFD2691E
   Crimson = &HFFDC143C
   Gold = &HFFFFD700
   Magenta = &HFFFF00FF
   Maroon = &HFF800000
   Olive = &HFF808000
   Orange = &HFFFFA500
   Purple = &HFF800080
   Red = &HFFFF0000
   Silver = &HFFC0C0C0
   Violet = &HFFEE82EE
   White = &HFFFFFFFF
   Yellow = &HFFFFFF00
End Enum

Const m_def_Value = 0
Const m_def_Max = 1
Const m_def_NeedleCol = Red
Const m_def_BackCol = vbWhite
Const m_def_Gradient = False
Const m_def_GradStartCol = vbRed
Const m_def_GradEndCol = vbWhite
Const m_def_GradStyle = 0

Dim m_Gradient As Boolean
Dim m_GradStartCol As OLE_COLOR
Dim m_GradEndCol As OLE_COLOR
Dim m_GradStyle As Integer
Dim m_NeedleCol As Colors
Dim m_BackCol As OLE_COLOR
Dim m_Value As Long
Dim m_Max As Long

Public Enum egGradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Colors, ByVal Width As Single, ByVal unit As GpUnit, ByRef pen As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal color As Colors, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As Long
Private Declare Function GdipDrawLine Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Long
Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As Long

Private Function ShutdownGDIPlus() As Long
    ShutdownGDIPlus = GdiplusShutdown(GdipToken)
End Function

Private Function StartUpGDIPlus(ByVal GdipVersion As Long) As Long
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = GdipVersion
    StartUpGDIPlus = GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Function
Private Sub UserControl_Initialize()
   Call StartUpGDIPlus(GdiPlusVersion)
   m_Value = m_def_Value
   m_Max = m_def_Max
   m_BackCol = m_def_BackCol
   m_Gradient = m_def_Gradient
   m_GradStartCol = m_def_GradStartCol
   m_GradEndCol = m_def_GradEndCol
   m_GradStyle = m_def_GradStyle
   m_NeedleCol = m_def_NeedleCol
End Sub

Private Sub DrawMeter()
 Dim x_Count As Integer
    Dim m_Externo As Integer
    Dim m_Interno As Integer
    Dim m_Graphics As Long
    Dim m_Pen As Long
    Dim m_Brush As Long
    Dim m_Count As Single
    Dim m_Sin As Integer
    Dim m_Cos As Integer
    Const PI = 3.14159265358979
    Const m_Number = PI / 180
    Dim m_X As Integer
    Dim m_Y As Integer
    Dim m_Radius As Single
    Dim mult As Single
   
    p1.Width = 119
    p1.Height = 58
  
    p1.Cls
    Call GdipCreateFromHDC(p1.hDC, m_Graphics)
    Call GdipSetSmoothingMode(m_Graphics, SmoothingModeAntiAlias)
       'arc
       m_X = p1.ScaleWidth / 2 - 6
       m_Y = p1.ScaleHeight - 3
       Call GdipCreatePen1(Black, 1, UnitPixel, m_Pen)
       Call GdipDrawArc(m_Graphics, m_Pen, p1.ScaleWidth / 2 - 46, p1.ScaleHeight / 2 - 15, 80, 80, -180, 180)
       'short tick marks
       m_Interno = 34
       m_Externo = 39
       
       For m_Count = 180 To 360 Step 9
          Call GdipCreatePen1(Black, 1, UnitPixel, m_Pen)
          Call GdipDrawLine(m_Graphics, m_Pen, m_X + m_Interno * Cos(PI / 180 * m_Count), m_Y + m_Interno * Sin(PI / 180 * m_Count), m_Externo * Cos(PI / 180 * m_Count) + m_X, m_Externo * Sin(PI / 180 * m_Count) + m_Y)
       Next m_Count
       p1.ForeColor = RGB(0, 0, 0)
       p1.Font = "Arial"
       p1.FontSize = 7
       p1.FontBold = True
       'Text
       For m_Count = 180 To 360 Step 18
          If m_Max <> "2" Then
             p1.CurrentX = m_X + 50 * Cos(PI / 180 * m_Count) - p1.TextWidth(x_Count) / 2 - 1
             p1.CurrentY = m_Y + 45 * Sin(PI / 180 * m_Count) - p1.TextHeight(x_Count) / 2 - 2
          Else
             p1.CurrentX = m_X + 51 * Cos(PI / 180 * m_Count) - p1.TextWidth(x_Count) / 2 + 2
             p1.CurrentY = m_Y + 45 * Sin(PI / 180 * m_Count) - p1.TextHeight(x_Count) / 2
             p1.FontSize = 6
          End If
       
       Select Case m_Max
       Case "1"
          p1.Print x_Count
          x_Count = x_Count + 10
          mult = 1.2
       Case "2"
          p1.Print x_Count
          x_Count = x_Count + 100
          mult = 0.12
       End Select
   
       Next m_Count
       
       'long tick marks
       m_Interno = 30
       For m_Count = 180 To 360 Step 18
          x_Count = x_Count + 1
          Call GdipCreatePen1(Black, 1, UnitPixel, m_Pen)
          Call GdipDrawLine(m_Graphics, m_Pen, m_X + m_Interno * Cos(PI / 180 * m_Count), m_Y + m_Interno * Sin(PI / 180 * m_Count), m_Externo * Cos(PI / 180 * m_Count) + m_X, m_Externo * Sin(PI / 180 * m_Count) + m_Y)
          p1.CurrentX = m_X + 85 * Cos(PI / 180 * m_Count) - p1.TextWidth(x_Count) / 2 - 2
          p1.CurrentY = m_Y + 85 * Sin(PI / 180 * m_Count) - p1.TextHeight(x_Count) / 2
       Next m_Count
       
       m_Radius = 40
       m_Sin = m_X + m_Radius * Sin((315 - (((45 / 180) * mult * Value) + 7.5) * 6) * m_Number)
       m_Cos = m_Y + m_Radius * Cos((315 - (((45 / 180) * mult * Value) + 7.5) * 6) * m_Number)
      
       Call GdipCreatePen1(Black, 4, UnitPixel, m_Pen)
       'needle
       Call GdipCreatePen1(m_NeedleCol, 1, UnitPixel, m_Pen)
       Call GdipDrawLine(m_Graphics, m_Pen, m_X, m_Y, m_Sin, m_Cos)
       'center circle
       Call GdipCreateSolidFill(Black, m_Brush)
       Call GdipFillEllipse(m_Graphics, m_Brush, m_X - 11, m_Y - 8, 21, 21)
    
       Call GdipDeleteBrush(m_Brush)
       Call GdipDeletePen(m_Pen)
       Call GdipDeleteGraphics(m_Graphics)
End Sub

Private Sub UserControl_InitProperties()
    NeedleCol = Red
    BackCol = m_BackCol
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = 119 * 15
   UserControl.Height = 58 * 15
   If m_Gradient = True Then
     PaintGradient p1.hDC, 0, 0, p1.ScaleWidth, p1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     p1.Picture = p1.Image
  Else
     p1.Picture = LoadPicture()
     p1.BackColor = m_BackCol
  End If
   DrawMeter
End Sub

Private Sub UserControl_Terminate()
   Call ShutdownGDIPlus
End Sub

Public Property Get Value() As Long
   Value = m_Value
End Property

Public Property Get NeedleCol() As Colors
   NeedleCol = m_NeedleCol
End Property

Public Property Let NeedleCol(NewNeedleCol As Colors)
  m_NeedleCol = NewNeedleCol
  PropertyChanged "NeedleCol"
  DrawMeter
End Property
Public Property Get BackCol() As OLE_COLOR
   BackCol = m_BackCol
End Property

Public Property Let BackCol(NewBackCol As OLE_COLOR)
  m_BackCol = NewBackCol
  PropertyChanged "BackCol"
  p1.BackColor = m_BackCol
  DrawMeter
End Property

Public Property Let Value(NewValue As Long)
  m_Value = NewValue
  PropertyChanged "Value"
  DrawMeter
End Property
Public Property Get Max() As eMax
   Max = m_Max
End Property

Public Property Let Max(NewMax As eMax)
  m_Max = NewMax
  PropertyChanged "Max"
  DrawMeter
End Property

Public Property Get Gradient() As Boolean
   Gradient = m_Gradient
End Property

Public Property Let Gradient(NewGradient As Boolean)
  m_Gradient = NewGradient
  PropertyChanged "Gradient"
  If m_Gradient = True Then
     PaintGradient p1.hDC, 0, 0, p1.ScaleWidth, p1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     p1.Picture = p1.Image
  Else
     p1.Picture = LoadPicture()
     p1.BackColor = m_BackCol
  End If
  DrawMeter
End Property
Public Property Get GradStartCol() As OLE_COLOR
   GradStartCol = m_GradStartCol
End Property

Public Property Let GradStartCol(NewGradStartCol As OLE_COLOR)
  m_GradStartCol = NewGradStartCol
  PropertyChanged "GradStartCol"
  If m_Gradient = True Then
     PaintGradient p1.hDC, 0, 0, p1.ScaleWidth, p1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     p1.Picture = p1.Image
     DrawMeter
  End If
End Property
Public Property Get GradEndCol() As OLE_COLOR
   GradEndCol = m_GradEndCol
End Property

Public Property Let GradEndCol(NewGradEndCol As OLE_COLOR)
  m_GradEndCol = NewGradEndCol
  PropertyChanged "GradEndCol"
  If m_Gradient = True Then
     PaintGradient p1.hDC, 0, 0, p1.ScaleWidth, p1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     p1.Picture = p1.Image
     DrawMeter
  End If
End Property

Public Property Get GradStyle() As egGradientDirectionCts
   GradStyle = m_GradStyle
End Property

Public Property Let GradStyle(NewGradStyle As egGradientDirectionCts)
  m_GradStyle = NewGradStyle
  PropertyChanged "GradStyle"
   If m_Gradient = True Then
     PaintGradient p1.hDC, 0, 0, p1.ScaleWidth, p1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     p1.Picture = p1.Image
     DrawMeter
  End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      Value = .ReadProperty("Value", m_def_Value)
      Max = .ReadProperty("Max", m_def_Max)
      NeedleCol = .ReadProperty("NeedleCol", m_def_NeedleCol)
      BackCol = .ReadProperty("BackCol", m_def_BackCol)
      Gradient = .ReadProperty("Gradient", m_def_Gradient)
      GradStartCol = .ReadProperty("GradStartCol", m_def_GradStartCol)
      GradEndCol = .ReadProperty("GradEndCol", m_def_GradEndCol)
      GradStyle = .ReadProperty("GradStyle", m_def_GradStyle)
   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "Max", m_Max, m_def_Max
      .WriteProperty "NeedleCol", m_NeedleCol, m_def_NeedleCol
      .WriteProperty "BackCol", m_BackCol, m_def_BackCol
      .WriteProperty "Gradient", m_Gradient, m_def_Gradient
      .WriteProperty "GradStartCol", m_GradStartCol, m_def_GradStartCol
      .WriteProperty "GradEndCol", m_GradEndCol, m_def_GradEndCol
      .WriteProperty "GradStyle", m_GradStyle, m_def_GradStyle
   End With
End Sub

Private Sub GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
  R = LngCol Mod 256    'Red
  G = (LngCol And vbGreen) / 256 'Green
  B = (LngCol And vbBlue) / 65536 'Blue
End Sub

' Author of this gradient code is Carles P.V. - 2005
Public Sub PaintGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal GradientDirection As egGradientDirectionCts _
                         )

  Dim uBIH    As BITMAPINFOHEADER
  Dim lBits() As Long
  Dim lGrad() As Long
  
  Dim R1      As Long
  Dim G1      As Long
  Dim B1      As Long
  Dim R2      As Long
  Dim G2      As Long
  Dim B2      As Long
  Dim dR      As Long
  Dim dG      As Long
  Dim dB      As Long
  
  Dim Scan    As Long
  Dim i       As Long
  Dim iEnd    As Long
  Dim iOffset As Long
  Dim j       As Long
  Dim jEnd    As Long
  Dim iGrad   As Long
  
    '-- A minor check
    If (Width < 1 Or Height < 1) Then Exit Sub
    
    '-- Decompose colors
    Color1 = Color1 And &HFFFFFF
    R1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    G1 = Color1 Mod &H100&
    Color1 = Color1 \ &H100&
    B1 = Color1 Mod &H100&
    Color2 = Color2 And &HFFFFFF
    R2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    G2 = Color2 Mod &H100&
    Color2 = Color2 \ &H100&
    B2 = Color2 Mod &H100&
    
    '-- Get color distances
    dR = R2 - R1
    dG = G2 - G1
    dB = B2 - B1
    
    '-- Size gradient-colors array
    Select Case GradientDirection
        Case [gdHorizontal]
            ReDim lGrad(0 To Width - 1)
        Case [gdVertical]
            ReDim lGrad(0 To Height - 1)
        Case Else
            ReDim lGrad(0 To Width + Height - 2)
    End Select
    
    '-- Calculate gradient-colors
    iEnd = UBound(lGrad())
    If (iEnd = 0) Then
        '-- Special case (1-pixel wide gradient)
        lGrad(0) = (B1 \ 2 + B2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
        For i = 0 To iEnd
            lGrad(i) = B1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
        Next i
    End If
    
    '-- Size DIB array
    ReDim lBits(Width * Height - 1) As Long
    iEnd = Width - 1
    jEnd = Height - 1
    Scan = Width
    
    '-- Render gradient DIB
    Select Case GradientDirection
        
        Case [gdHorizontal]
        
            For j = 0 To jEnd
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(i - iOffset)
                Next i
                iOffset = iOffset + Scan
            Next j
        
        Case [gdVertical]
        
            For j = jEnd To 0 Step -1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(j)
                Next i
                iOffset = iOffset + Scan
            Next j
            
        Case [gdDownwardDiagonal]
            
            iOffset = jEnd * Scan
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset - Scan
                iGrad = j
            Next j
            
        Case [gdUpwardDiagonal]
            
            iOffset = 0
            For j = 1 To jEnd + 1
                For i = iOffset To iEnd + iOffset
                    lBits(i) = lGrad(iGrad)
                    iGrad = iGrad + 1
                Next i
                iOffset = iOffset + Scan
                iGrad = j
            Next j
    End Select
    
    '-- Define DIB header
    With uBIH
        .biSize = 40
        .biPlanes = 1
        .biBitCount = 32
        .biWidth = Width
        .biHeight = Height
    End With
    
    '-- Paint it!
    Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
End Sub

