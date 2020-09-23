VERSION 5.00
Begin VB.UserControl ucMeter 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2205
   DrawWidth       =   6
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   2160
   ScaleWidth      =   2205
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   30
      ScaleHeight     =   67
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   0
      Top             =   15
      Width           =   1320
   End
   Begin VB.Shape Shape1 
      Height          =   1860
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   -15
      TabIndex        =   1
      Top             =   1575
      Width           =   1875
   End
End
Attribute VB_Name = "ucMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Ken Foster Dec 2009
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

Public Enum eGradientDirectionCts
    [gdHorizontal] = 0
    [gdVertical] = 1
    [gdDownwardDiagonal] = 2
    [gdUpwardDiagonal] = 3
End Enum

Public Enum eTickStyle
   Fourths = 0
   Tenths = 1
End Enum

Const m_def_Value = 0
Const m_def_Caption = "Meter"
Const m_def_LabelColor = vbBlack
Const m_def_TextColor = vbWhite
Const m_def_NeedleColor = vbRed
Const m_def_BackColor = vbWhite
Const m_def_TickColor = vbBlack
Const m_def_DrawScrews = True
Const m_def_TickLabels = True
Const m_def_CaptionFace = ""
Const m_def_Gradient = False
Const m_def_GradStartCol = vbRed
Const m_def_GradEndCol = vbWhite
Const m_def_GradStyle = 0
Const m_def_TickStyle = 0

Dim m_TickStyle As Integer
Dim m_Gradient As Boolean
Dim m_GradStartCol As OLE_COLOR
Dim m_GradEndCol As OLE_COLOR
Dim m_GradStyle As Integer
Dim m_CaptionFace As String
Dim m_DrawScrews As Boolean
Dim m_TickLabels As Boolean
Dim m_TickColor As OLE_COLOR
Dim m_LabelColor As OLE_COLOR
Dim m_TextColor As OLE_COLOR
Dim m_NeedleColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_Caption As String
Dim Max As Integer
Dim m_Value As Single
Dim ley As Single  'needle length

Private Sub UserControl_Initialize()
  Max = 100  'this is the base range. changing it will cause the needle to move weird.
  pic1.BackColor = UserControl.BackColor
  ley = 19
End Sub

Private Sub UserControl_Resize()
  pic1.Left = 20
  pic1.Height = UserControl.Height - Label1.Height - 8
  pic1.Width = UserControl.Width - 36
  Shape1.Width = UserControl.Width - 6
  Shape1.Height = UserControl.Height - 2
  Label1.Width = UserControl.Width
  Label1.Top = pic1.Top + pic1.Height
  Label1.Left = 0
 If m_Gradient = True Then
     PaintGradient pic1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = pic1.Image
  Else
     pic1.Picture = LoadPicture()
     pic1.BackColor = m_BackColor
  End If
  DrawScale
End Sub

Private Sub UserControl_InitProperties()
   m_Caption = Extender.Name
   m_Value = m_def_Value
   m_LabelColor = m_def_LabelColor
   m_TextColor = m_def_TextColor
   m_NeedleColor = m_def_NeedleColor
   m_BackColor = m_def_BackColor
   m_TickColor = m_def_TickColor
   Caption = Extender.Name
   m_DrawScrews = m_def_DrawScrews
   m_TickLabels = m_def_TickLabels
   m_CaptionFace = m_def_CaptionFace
   m_Gradient = m_def_Gradient
   m_GradStartCol = m_def_GradStartCol
   m_GradEndCol = m_def_GradEndCol
   m_GradStyle = m_def_GradStyle
   m_TickStyle = m_def_TickStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      BackColor = .ReadProperty("BackColor", m_def_BackColor)
      Caption = .ReadProperty("Caption", m_def_Caption)
      CaptionFace = .ReadProperty("CaptionFace", m_def_CaptionFace)
      DrawScrews = .ReadProperty("DrawScrews", m_def_DrawScrews)
      LabelColor = .ReadProperty("LabelColor", m_def_LabelColor)
      TextColor = .ReadProperty("TextColor", m_def_TextColor)
      NeedleColor = .ReadProperty("NeedleColor", m_def_NeedleColor)
      Value = .ReadProperty("Value", m_def_Value)
      TickColor = .ReadProperty("TickColor", m_def_TickColor)
      TickLabels = .ReadProperty("TickLabels", m_def_TickLabels)
      TickStyle = .ReadProperty("TickStyle", m_def_TickStyle)
      Gradient = .ReadProperty("Gradient", m_def_Gradient)
      GradStartCol = .ReadProperty("GradStartCol", m_def_GradStartCol)
      GradEndCol = .ReadProperty("GradEndCol", m_def_GradEndCol)
      GradStyle = .ReadProperty("GradStyle", m_def_GradStyle)
   End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "BackColor", m_BackColor, m_def_BackColor
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "CaptionFace", m_CaptionFace, m_def_CaptionFace
      .WriteProperty "DrawScrews", m_DrawScrews, m_def_DrawScrews
      .WriteProperty "LabelColor", m_LabelColor, m_def_LabelColor
      .WriteProperty "TextColor", m_TextColor, m_def_TextColor
      .WriteProperty "NeedleColor", m_NeedleColor, m_def_NeedleColor
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "TickColor", m_TickColor, m_def_TickColor
      .WriteProperty "TickLabels", m_TickLabels, m_def_TickLabels
      .WriteProperty "TickStyle", m_TickStyle, m_def_TickStyle
      .WriteProperty "Gradient", m_Gradient, m_def_Gradient
      .WriteProperty "GradStartCol", m_GradStartCol, m_def_GradStartCol
      .WriteProperty "GradEndCol", m_GradEndCol, m_def_GradEndCol
      .WriteProperty "GradStyle", m_GradStyle, m_def_GradStyle
   End With
End Sub

Private Sub DrawScale()
Dim cx As Integer
Dim cy As Integer
Dim sStart As Integer
Dim sEnd As Integer
Dim lex As Single
Dim Sscale As Single
Dim Sratio As Single
Dim rt As Single

  pic1.Cls
  'meter center
  cx = pic1.ScaleWidth / 2
  cy = pic1.ScaleHeight
  pic1.DrawWidth = 2
  Sratio = pic1.ScaleWidth / Max
  lex = Value * Sratio
  ley = 0.3 * (Abs(Value - (Max / 2))) + 5


If DrawScrews = True Then
   pic1.FillColor = vbBlack
   'draw screws
   pic1.Circle (10, pic1.ScaleHeight - 10), 3, vbBlack        'left screw bottom
   pic1.Line (7, pic1.ScaleHeight - 7)-(13, pic1.ScaleHeight - 13), &H404040
   pic1.Circle (pic1.ScaleWidth - 10, pic1.ScaleHeight - 10), 3, vbBlack        'right screw bottom
   pic1.Line (pic1.ScaleWidth - 12, pic1.ScaleHeight - 12)-(pic1.ScaleWidth - 7, pic1.ScaleHeight - 7), &H404040
   
   pic1.Circle (10, 5), 3, vbBlack        'left screw top
   pic1.Line (7, 5)-(13, 5), &H404040
   pic1.Circle (pic1.ScaleWidth - 10, 5), 3, vbBlack         'right screw top
   pic1.Line (pic1.ScaleWidth - 12, 7)-(pic1.ScaleWidth - 7, 3), &H404040
End If

  'needle
  pic1.Line (cx, cy)-(lex, ley), m_NeedleColor
  'center circle
  pic1.FillColor = m_LabelColor
  pic1.Circle (cx, cy), 10, m_LabelColor
  'print caption face
  pic1.ForeColor = m_TickColor
  pic1.CurrentX = cx - (Len(m_CaptionFace) * 3)
  pic1.CurrentY = cy - 25
  pic1.FontBold = True
  pic1.Print m_CaptionFace
  pic1.FontBold = False
  
  If TickStyle = 0 Then
     'tick marks
     '0
     pic1.PSet (1, 19), m_TickColor
     '25
     pic1.PSet (pic1.ScaleWidth / 4, 10), m_TickColor
     '50
     pic1.PSet (pic1.ScaleWidth / 2, 5), m_TickColor
     '75
     pic1.PSet (pic1.ScaleWidth / 2 + pic1.ScaleWidth / 4 + 2, 10), m_TickColor
    '100
     pic1.PSet (pic1.ScaleWidth - 2, 20), m_TickColor

If TickLabels = True Then
  'draw tick labels
  pic1.ForeColor = m_TickColor
  pic1.CurrentX = 3
  pic1.CurrentY = 18
  pic1.Print "0"
  pic1.CurrentX = pic1.ScaleWidth / 4 - 6
  pic1.CurrentY = 12
  pic1.Print "25"
  pic1.CurrentX = pic1.ScaleWidth / 2 - 6
  pic1.CurrentY = 8
  pic1.Print "50"
  pic1.CurrentX = pic1.ScaleWidth / 4 + pic1.ScaleWidth / 2 - 8
  pic1.CurrentY = 12
  pic1.Print "75"
  pic1.CurrentX = pic1.ScaleWidth - 20
  pic1.CurrentY = 21
  pic1.Print "100"
End If
Else
    rt = pic1.ScaleWidth / 10 + 1
    pic1.PSet (1, 19), m_TickColor
    pic1.PSet (rt - 1, 16), m_TickColor
    pic1.PSet (rt * 2 - 2, 12), m_TickColor
    pic1.PSet (rt * 3 - 3, 9), m_TickColor
    pic1.PSet (rt * 4 - 4, 6), m_TickColor
    pic1.PSet (pic1.ScaleWidth / 2, 4), m_TickColor
    pic1.PSet (pic1.ScaleWidth / 2 + rt, 6), m_TickColor
    pic1.PSet (pic1.ScaleWidth / 2 + rt * 2 - 1, 9), m_TickColor
    pic1.PSet (pic1.ScaleWidth / 2 + rt * 3 - 2, 12), m_TickColor
    pic1.PSet (pic1.ScaleWidth / 2 + rt * 4 - 3, 16), m_TickColor
    pic1.PSet (pic1.ScaleWidth - 2, 20), m_TickColor
    
    If TickLabels = True Then
        pic1.ForeColor = m_TickColor
        pic1.CurrentX = 3
        pic1.CurrentY = 18
        pic1.Print "0"
        pic1.CurrentX = pic1.ScaleWidth / 4 - 12
        pic1.CurrentY = 14
        pic1.CurrentX = pic1.ScaleWidth / 2 - 6
        pic1.CurrentY = 8
        pic1.Print "50"
        pic1.CurrentX = pic1.ScaleWidth / 4 + pic1.ScaleWidth / 2
        pic1.CurrentY = 14
        pic1.CurrentX = pic1.ScaleWidth - 20
        pic1.CurrentY = 21
        pic1.Print "100"
    End If
End If
End Sub

Public Property Get Gradient() As Boolean
   Gradient = m_Gradient
End Property

Public Property Let Gradient(NewGradient As Boolean)
  m_Gradient = NewGradient
  PropertyChanged "Gradient"
  If m_Gradient = True Then
     PaintGradient pic1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = pic1.Image
  Else
     pic1.Picture = LoadPicture()
     pic1.BackColor = m_BackColor
  End If
  DrawScale
End Property
Public Property Get GradStartCol() As OLE_COLOR
   GradStartCol = m_GradStartCol
End Property

Public Property Let GradStartCol(NewGradStartCol As OLE_COLOR)
  m_GradStartCol = NewGradStartCol
  PropertyChanged "GradStartCol"
  If m_Gradient = True Then
     PaintGradient pic1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = pic1.Image
     DrawScale
  End If
End Property
Public Property Get GradEndCol() As OLE_COLOR
   GradEndCol = m_GradEndCol
End Property

Public Property Let GradEndCol(NewGradEndCol As OLE_COLOR)
  m_GradEndCol = NewGradEndCol
  PropertyChanged "GradEndCol"
  If m_Gradient = True Then
     PaintGradient pic1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = pic1.Image
     DrawScale
  End If
End Property

Public Property Get GradStyle() As eGradientDirectionCts
   GradStyle = m_GradStyle
End Property

Public Property Let GradStyle(NewGradStyle As eGradientDirectionCts)
  m_GradStyle = NewGradStyle
  PropertyChanged "GradStyle"
   If m_Gradient = True Then
     PaintGradient pic1.hDC, 0, 0, pic1.ScaleWidth, pic1.ScaleHeight, GradStartCol, GradEndCol, GradStyle
     pic1.Picture = pic1.Image
     DrawScale
  End If
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = m_BackColor
End Property

Public Property Let BackColor(NewBackColor As OLE_COLOR)
  m_BackColor = NewBackColor
  UserControl.BackColor = m_BackColor
  pic1.BackColor = UserControl.BackColor
  PropertyChanged "BackColor"
  DrawScale
End Property

Public Property Get Caption() As String
   Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
  m_Caption = NewCaption
  Label1.Caption = m_Caption
  PropertyChanged "Caption"
  DrawScale
End Property

Public Property Get CaptionFace() As String
   CaptionFace = m_CaptionFace
End Property

Public Property Let CaptionFace(NewCaptionFace As String)
  m_CaptionFace = NewCaptionFace
  PropertyChanged "CaptionFace"
  DrawScale
End Property

Public Property Get DrawScrews() As Boolean
   DrawScrews = m_DrawScrews
End Property

Public Property Let DrawScrews(NewDrawScrews As Boolean)
  m_DrawScrews = NewDrawScrews
  PropertyChanged "DrawScrews"
  DrawScale
End Property

Public Property Get LabelColor() As OLE_COLOR
   LabelColor = m_LabelColor
End Property

Public Property Let LabelColor(NewLabelColor As OLE_COLOR)
  m_LabelColor = NewLabelColor
  Label1.BackColor = m_LabelColor
  PropertyChanged "LabelColor"
  DrawScale
End Property

Public Property Get NeedleColor() As OLE_COLOR
   NeedleColor = m_NeedleColor
End Property

Public Property Let NeedleColor(NewNeedleColor As OLE_COLOR)
  m_NeedleColor = NewNeedleColor
  PropertyChanged "NeedleColor"
  DrawScale
End Property

Public Property Get TextColor() As OLE_COLOR
   TextColor = m_TextColor
End Property

Public Property Let TextColor(NewTextColor As OLE_COLOR)
  m_TextColor = NewTextColor
  Label1.ForeColor = m_TextColor
  PropertyChanged "TextColor"
  DrawScale
End Property

Public Property Get TickColor() As OLE_COLOR
   TickColor = m_TickColor
End Property

Public Property Let TickColor(NewTickColor As OLE_COLOR)
  m_TickColor = NewTickColor
  PropertyChanged "TickColor"
  DrawScale
End Property

Public Property Get TickLabels() As Boolean
   TickLabels = m_TickLabels
End Property

Public Property Let TickLabels(NewTickLabels As Boolean)
  m_TickLabels = NewTickLabels
  PropertyChanged "TickLabels"
  DrawScale
End Property

Public Property Get TickStyle() As eTickStyle
   TickStyle = m_TickStyle
End Property

Public Property Let TickStyle(NewTickStyle As eTickStyle)
  m_TickStyle = NewTickStyle
  PropertyChanged "TickStyle"
  DrawScale
End Property

Public Property Get Value() As Single
   Value = m_Value
End Property

Public Property Let Value(NewValue As Single)
  m_Value = NewValue
  PropertyChanged "Value"
  DrawScale
End Property
' Author of this gradient code is Carles P.V. - 2005
Public Sub PaintGradient(ByVal hDC As Long, _
                         ByVal x As Long, _
                         ByVal y As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long, _
                         ByVal Color1 As Long, _
                         ByVal Color2 As Long, _
                         ByVal GradientDirection As eGradientDirectionCts _
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


