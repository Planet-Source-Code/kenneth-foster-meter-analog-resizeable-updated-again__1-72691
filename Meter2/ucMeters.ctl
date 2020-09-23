VERSION 5.00
Begin VB.UserControl ucMeters 
   ClientHeight    =   1155
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1680
   ScaleHeight     =   77
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   990
      ScaleHeight     =   1365
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   1395
      Width           =   1755
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   72
      TabIndex        =   0
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Meter 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   750
      Width           =   1110
   End
End
Attribute VB_Name = "ucMeters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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

Public Enum GradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

Public Enum eMeterType
   Small = 0
   Med = 1
   Large = 2
End Enum

Const m_def_MeterType = 2
Const m_def_Value = 0
Const m_def_LabelColor = vbBlack
Const m_def_NeedleColor = vbRed
Const m_def_Caption = ""
Const m_def_CaptionColor = vbWhite
Const m_def_TickColor = vbBlack
Const m_def_Gradient = False
Const m_def_GradStartCol = vbRed
Const m_def_GradEndCol = vbWhite
Const m_def_GradStyle = 0

Dim m_Gradient As Boolean
Dim m_GradStartCol As OLE_COLOR
Dim m_GradEndCol As OLE_COLOR
Dim m_GradStyle As Integer
Dim m_TickColor As OLE_COLOR
Dim m_CaptionColor As OLE_COLOR
Dim m_Caption As String
Dim m_Value As Integer
Dim m_LabelColor As OLE_COLOR
Dim m_NeedleColor As OLE_COLOR
Dim m_MeterType As Integer

Dim sl As Integer
Dim R1 As Double
Dim R2 As Double
Const vbPI = 3.141592654
Const Deg2Rad = vbPI / 180 'Degrees to Radians
Dim cx As Integer, cy As Integer, r As Single
Dim s As Double
Dim Emax As Double
Dim EMin As Double
Dim Mmax As Double
Dim Mmin As Double
Dim Mdeg As Integer
Dim sin_ As Double
Dim cos_ As Double
Dim temp As Double

Private Sub DrawDial()   'Original code by Max Seim

If MeterType = 0 Then
   pic1.Height = 38
   pic1.Width = pic1.Height * 1.7
   Mmin = 19.8
   Mmax = 40
   EMin = 0
   Emax = 100
End If

If MeterType = 1 Then
   pic1.Height = 50
   pic1.Width = pic1.Height * 1.7
   Mmin = 21.3
   Mmax = 38.7
   EMin = 0
   Emax = 100
End If

If MeterType = 2 Then
   pic1.Height = 70
   pic1.Width = pic1.Height * 1.5
   Mmin = 22.7
   Mmax = 37.4
   Emax = 100
End If
  Label1.Top = pic1.Top + pic1.Height - 2
  Label1.Left = pic1.Left
  Label1.Width = pic1.Width
  Mdeg = 0 ' degrees (0-360)
  cx = pic1.Width / 2 - 2
  cy = pic1.Height
  temp = pic1.CurrentY
  r = IIf(cx > cy, cy, cx) + pic1.Height / 2
  sl = r * 0.1 - temp + 35 'length of meter hand

  ' Scale the Engineering Units
  R1 = Emax - EMin
  R2 = Mmax - Mmin
  s = ((R2 / R1) * m_Value) + Mmin
  
  ' Draw the dial hand
  sin_ = Sin((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cx
  cos_ = Cos((Mdeg - s * 6) * Deg2Rad) * (r - sl) + cy
  pic1.ForeColor = m_NeedleColor
  pic1.DrawWidth = 2
  pic1.Cls
  pic1.Line (cx, cy)-(sin_, cos_)
  'center circle
  pic1.FillColor = m_LabelColor
  pic1.Circle (cx, cy), 8, m_LabelColor
  If MeterType <> 0 Then
  'draw screws
  pic1.Circle (10, pic1.ScaleHeight - 10), 3, vbBlack        'left screw
  pic1.Line (7, pic1.ScaleHeight - 13)-(13, pic1.ScaleHeight - 7), &H404040
  pic1.Circle (pic1.ScaleWidth - 10, pic1.ScaleHeight - 10), 3, vbBlack        'right screw
  pic1.Line (pic1.ScaleWidth - 12, pic1.ScaleHeight - 12)-(pic1.ScaleWidth - 7, pic1.ScaleHeight - 7), &H404040
  End If
  'tick marks
  '0
  pic1.PSet (5, 22), m_TickColor
  '25
  pic1.PSet (pic1.ScaleWidth / 4, 10), m_TickColor
  '50
  pic1.PSet (pic1.ScaleWidth / 2, 5), m_TickColor
  '75
  pic1.PSet (pic1.ScaleWidth / 2 + pic1.ScaleWidth / 4 - 2, 10), m_TickColor
  '100
  pic1.PSet (pic1.ScaleWidth - 7, 22), m_TickColor
  
  'print tick labels
  pic1.ForeColor = m_TickColor
  pic1.CurrentX = 4
  pic1.CurrentY = 25
  pic1.Print "0"
  pic1.CurrentX = pic1.ScaleWidth / 4 - 6
  pic1.CurrentY = 14
  pic1.Print "25"
  pic1.CurrentX = pic1.ScaleWidth / 2 - 6
  pic1.CurrentY = 9
  pic1.Print "50"
  pic1.CurrentX = pic1.ScaleWidth / 4 + pic1.ScaleWidth / 2 - 8
  pic1.CurrentY = 14
  pic1.Print "75"
  pic1.CurrentX = pic1.ScaleWidth - 20
  pic1.CurrentY = 26
  pic1.Print "100"
  pic1.CurrentX = 20
  pic1.CurrentY = 20
End Sub

Private Sub UserControl_Initialize()
  m_MeterType = m_def_MeterType
  m_Value = m_def_Value
  m_LabelColor = m_def_LabelColor
  m_NeedleColor = m_def_NeedleColor
  m_Caption = m_def_Caption
  m_CaptionColor = m_def_CaptionColor
  m_TickColor = m_def_TickColor
  m_Gradient = m_def_Gradient
  m_GradStartCol = m_def_GradStartCol
  m_GradEndCol = m_def_GradEndCol
  m_GradStyle = m_def_GradStyle
End Sub

Private Sub UserControl_InitProperties()
   Caption = Extender.Name
   MeterType = 2
   pic1.CurrentY = 9
End Sub

Private Sub UserControl_Resize()
   Select Case MeterType
   Case 0
      UserControl.Width = 70 * 14
      UserControl.Height = 60 * 14
   Case 1
      UserControl.Width = 93 * 14
      UserControl.Height = 72 * 14
   Case 2
      UserControl.Width = 112 * 14
      UserControl.Height = 94 * 14
   End Select
   DrawDial
End Sub

Public Property Get MeterType() As eMeterType
   MeterType = m_MeterType
End Property

Public Property Let MeterType(NewMeterType As eMeterType)
  m_MeterType = NewMeterType
  PropertyChanged "MeterType"
  UserControl_Resize
End Property

Public Property Get Value() As Integer
   Value = m_Value
End Property

Public Property Let Value(NewValue As Integer)
  m_Value = NewValue
  PropertyChanged "Value"
  DrawDial
End Property
Public Property Get LabelColor() As OLE_COLOR
   LabelColor = m_LabelColor
End Property

Public Property Let LabelColor(NewLabelColor As OLE_COLOR)
  m_LabelColor = NewLabelColor
  PropertyChanged "LabelColor"
  Label1.BackColor = m_LabelColor
  DrawDial
End Property
Public Property Get NeedleColor() As OLE_COLOR
   NeedleColor = m_NeedleColor
End Property

Public Property Let NeedleColor(NewNeedleColor As OLE_COLOR)
  m_NeedleColor = NewNeedleColor
  PropertyChanged "NeedleColor"
  DrawDial
End Property

Public Property Get TickColor() As OLE_COLOR
   TickColor = m_TickColor
End Property

Public Property Let TickColor(NewTickColor As OLE_COLOR)
  m_TickColor = NewTickColor
  PropertyChanged "TickColor"
  DrawDial
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption"
  Label1.Caption = m_Caption
  DrawDial
End Property

Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(NewCaptionColor As OLE_COLOR)
  m_CaptionColor = NewCaptionColor
  PropertyChanged "CaptionColor"
  Label1.ForeColor = m_CaptionColor
  DrawDial
End Property

Public Property Get Gradient() As Boolean
   Gradient = m_Gradient
End Property

Public Property Let Gradient(NewGradient As Boolean)
  m_Gradient = NewGradient
  PropertyChanged "Gradient"
  If m_Gradient = True Then
     PaintGradient Picture1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = Picture1.Image
  Else
     pic1.Picture = LoadPicture()
     pic1.BackColor = vbWhite
  End If
  DrawDial
End Property
Public Property Get GradStartCol() As OLE_COLOR
   GradStartCol = m_GradStartCol
End Property

Public Property Let GradStartCol(NewGradStartCol As OLE_COLOR)
  m_GradStartCol = NewGradStartCol
  PropertyChanged "GradStartCol"
  If m_Gradient = True Then
     PaintGradient Picture1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = Picture1.Image
     DrawDial
  End If
End Property
Public Property Get GradEndCol() As OLE_COLOR
   GradEndCol = m_GradEndCol
End Property

Public Property Let GradEndCol(NewGradEndCol As OLE_COLOR)
  m_GradEndCol = NewGradEndCol
  PropertyChanged "GradEndCol"
  If m_Gradient = True Then
     PaintGradient Picture1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = Picture1.Image
     DrawDial
  End If
End Property
Public Property Get GradStyle() As GradientDirectionCts
   GradStyle = m_GradStyle
End Property

Public Property Let GradStyle(NewGradStyle As GradientDirectionCts)
  m_GradStyle = NewGradStyle
  PropertyChanged "GradStyle"
   If m_Gradient = True Then
     PaintGradient Picture1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = Picture1.Image
     DrawDial
  End If
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      MeterType = .ReadProperty("MeterType", m_def_MeterType)
      Value = .ReadProperty("Value", m_def_Value)
      LabelColor = .ReadProperty("LabelColor", m_def_LabelColor)
      NeedleColor = .ReadProperty("NeedleColor", m_def_NeedleColor)
      Caption = .ReadProperty("Caption", m_def_Caption)
      CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
      TickColor = .ReadProperty("TickColor", m_def_TickColor)
      Gradient = .ReadProperty("Gradient", m_def_Gradient)
      GradStartCol = .ReadProperty("GradStartCol", m_def_GradStartCol)
      GradEndCol = .ReadProperty("GradEndCol", m_def_GradEndCol)
      GradStyle = .ReadProperty("GradStyle", m_def_GradStyle)

   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "MeterType", m_MeterType, m_def_MeterType
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "LabelColor", m_LabelColor, m_def_LabelColor
      .WriteProperty "NeedleColor", m_NeedleColor, m_def_NeedleColor
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "TickColor", m_TickColor, m_def_TickColor
      .WriteProperty "Gradient", m_Gradient, m_def_Gradient
      .WriteProperty "GradStartCol", m_GradStartCol, m_def_GradStartCol
      .WriteProperty "GradEndCol", m_GradEndCol, m_def_GradEndCol
      .WriteProperty "GradStyle", m_GradStyle, m_def_GradStyle

   End With
End Sub

Public Sub PaintGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal GradientDirection As GradientDirectionCts _
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

