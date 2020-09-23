VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Meters Demo"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5010
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HS4 
      Height          =   270
      Left            =   2715
      Max             =   100
      TabIndex        =   11
      Top             =   4290
      Width           =   1875
   End
   Begin VB.HScrollBar HS3 
      Height          =   300
      Left            =   2580
      Max             =   100
      TabIndex        =   10
      Top             =   2910
      Width           =   2010
   End
   Begin VB.HScrollBar HS2 
      Height          =   255
      Left            =   300
      Max             =   100
      TabIndex        =   9
      Top             =   3915
      Width           =   1755
   End
   Begin Project1.ucMeter ucMeter2 
      Height          =   1125
      Left            =   2610
      TabIndex        =   8
      Top             =   1530
      Width           =   1920
      _ExtentX        =   3334
      _ExtentY        =   1984
      Caption         =   "Meter 2"
      NeedleColor     =   16711680
      TickStyle       =   1
   End
   Begin Project1.MeterGDIPlus MeterGDIPlus1 
      Height          =   870
      Left            =   2760
      TabIndex        =   7
      Top             =   3375
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   1535
      BackCol         =   12648447
   End
   Begin Project1.ucMeters ucMeters3 
      Height          =   1320
      Left            =   330
      TabIndex        =   3
      Top             =   2190
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2328
      Caption         =   "ucMeters3"
      Gradient        =   -1  'True
      GradStartCol    =   16777215
      GradEndCol      =   8421631
      GradStyle       =   1
   End
   Begin Project1.ucMeters ucMeters2 
      Height          =   1005
      Left            =   315
      TabIndex        =   2
      Top             =   1035
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1773
      MeterType       =   1
      NeedleColor     =   4210752
      Caption         =   "ucMeters2"
      CaptionColor    =   33023
      Gradient        =   -1  'True
      GradStartCol    =   16777215
      GradEndCol      =   16744576
      GradStyle       =   1
   End
   Begin Project1.ucMeter ucMeter1 
      Height          =   1110
      Left            =   2595
      TabIndex        =   1
      Top             =   165
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   1958
      BackColor       =   8438015
      Caption         =   "Meter 1"
      Gradient        =   -1  'True
      GradStartCol    =   8421631
      GradStyle       =   1
   End
   Begin Project1.ucMeters ucMeters1 
      Height          =   840
      Left            =   330
      TabIndex        =   0
      Top             =   150
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1482
      MeterType       =   0
      Caption         =   "meter"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Two Scale Types"
      Height          =   210
      Left            =   2970
      TabIndex        =   13
      Top             =   2685
      Width           =   1365
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "two scales"
      Height          =   225
      Left            =   4020
      TabIndex        =   12
      Top             =   4575
      Width           =   825
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter 5 (GDI Plus)"
      Height          =   240
      Left            =   2655
      TabIndex        =   6
      Top             =   4575
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter 3    Resizeable"
      Height          =   195
      Left            =   2820
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Meter2  (3 sizes)"
      Height          =   285
      Left            =   525
      TabIndex        =   4
      Top             =   3615
      Width           =   1305
   End
   Begin VB.Shape Shape3 
      Height          =   3165
      Left            =   2475
      Top             =   90
      Width           =   2250
   End
   Begin VB.Shape Shape1 
      Height          =   4245
      Left            =   240
      Top             =   60
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HS2_Change()
   HS2_Scroll
End Sub

Private Sub HS2_Scroll()
   ucMeters1.Value = HS2.Value
   ucMeters2.Value = HS2.Value
   ucMeters3.Value = HS2.Value
End Sub

Private Sub HS3_Change()
   HS3_Scroll
End Sub

Private Sub HS3_Scroll()
   ucMeter1.Value = HS3.Value
   ucMeter2.Value = HS3.Value
End Sub

Private Sub HS4_Change()
   HS4_Scroll
End Sub

Private Sub HS4_Scroll()
   MeterGDIPlus1.Value = HS4.Value
End Sub
