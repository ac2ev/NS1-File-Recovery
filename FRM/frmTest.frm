VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "APdata"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14700
   FillColor       =   &H0000FF00&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   14700
   StartUpPosition =   3  'Windows Default
   Begin NS1_Recovery.GraphlitePro GraphlitePro1 
      Height          =   7920
      Left            =   75
      TabIndex        =   8
      Top             =   510
      Width           =   14580
      _ExtentX        =   25718
      _ExtentY        =   13970
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Stacked"
      Height          =   195
      Index           =   2
      Left            =   5100
      TabIndex        =   7
      Top             =   255
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   375
      Left            =   6405
      TabIndex        =   6
      Top             =   60
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Point"
      Height          =   195
      Index           =   3
      Left            =   4200
      TabIndex        =   5
      Top             =   270
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Line"
      Height          =   195
      Index           =   1
      Left            =   5100
      TabIndex        =   4
      Top             =   0
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bar"
      Height          =   195
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Display Legends"
      Height          =   195
      Left            =   2580
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Plot Points"
      Height          =   195
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   1275
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Noise"
      Height          =   195
      Left            =   10830
      TabIndex        =   10
      Top             =   150
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Signal"
      Height          =   195
      Left            =   9330
      TabIndex        =   9
      Top             =   150
      Width           =   435
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   10335
      Top             =   90
      Width           =   405
   End
   Begin VB.Image Image2 
      Height          =   330
      Left            =   10320
      Top             =   82
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   8850
      Top             =   90
      Width           =   405
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   8835
      Top             =   82
      Width           =   435
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Check1_Click()

GraphlitePro1.PlotPoints = Check1 * -1
GraphlitePro1.Refresh

End Sub

Private Sub Check2_Click()

GraphlitePro1.DisplayLegend = Check2 * -1
GraphlitePro1.Refresh

End Sub

Private Sub cmdPrint_Click()
Dim cdlg As New cdlg
Dim Owner As Long
Dim DisableMargins As Boolean
Dim DisableOrientation As Boolean
Dim DisablePaper As Boolean
Dim DisablePrinter As Boolean
Dim LeftMargin As Long
Dim MinLeftMargin As Long
Dim RightMargin As Long
Dim MinRightMargin As Long
Dim TopMargin As Long
Dim MinTopMargin As Long
Dim BottomMargin As Long
Dim MinBottomMargin As Long
Dim PaperSize As EPaperSize
Dim Orientation As EOrientation
Dim PrintQuality As EPrintQuality
Dim Units As EPageSetupUnits
Dim Printer As Object
Dim flags As Long
Dim Hook As Boolean
Dim EventSink As Object
 cdlg.VBPageSetupDlg , , , , , LeftMargin, , RightMargin, , TopMargin, , BottomMargin, , PaperSize, Orientation
 
    GraphlitePro1.PrintGraph PaperSize, Orientation, LeftMargin, RightMargin, TopMargin, BottomMargin

End Sub

Public Sub Command1_Click()
Dim DataPoints As Integer
Dim n As Integer

'If GraphLitePro1.ChartType = Bar Then
'   DataPoints = 10
'Else
'   DataPoints = 30
'End If

ReDim APData(2, UBound(ns1.APINFO(gphIndex).APData) - 1) As Variant
Dim l As Long
For l = LBound(ns1.APINFO(gphIndex).APData) To UBound(ns1.APINFO(gphIndex).APData) - 1
    APData(0, l) = ns1.APINFO(gphIndex).APData(l).Time.Time
    APData(1, l) = ns1.APINFO(gphIndex).APData(l).Signal 'IIf(ns1.APINFO(gphIndex).APData(l).Signal > 0 Or ns1.APINFO(gphIndex).APData(l).Signal = -32767, -150, ns1.APINFO(gphIndex).APData(l).Signal)
    APData(2, l) = ns1.APINFO(gphIndex).APData(l).Noise 'IIf(ns1.APINFO(gphIndex).APData(l).Noise > 0 Or ns1.APINFO(gphIndex).APData(l).Noise = -32767, -150, ns1.APINFO(gphIndex).APData(l).Noise)
Next l
GraphlitePro1.BackColor = &H0      'black
GraphlitePro1.ForeColor = &HFFFFFF ' white
GraphlitePro1.RegisterData APData
If Shape1.FillColor = vbGreen Then GraphlitePro1.SetSeriesOptions 0, vbGreen, "Signal"
If Shape2.FillColor = vbRed Then GraphlitePro1.SetSeriesOptions 1, vbRed, "Noise"
GraphlitePro1.Title = "Signal/Noise dBm"
GraphlitePro1.LowScale = -100
GraphlitePro1.HighScale = -99
GraphlitePro1.VerticalTickInterval = 10
GraphlitePro1.HorizontalTickFrequency = 60
For n = 0 To 3
   If Option1(n) Then
      GraphlitePro1.ChartType = n
      Exit For
   End If
Next n
GraphlitePro1.Refresh

End Sub
'
'Private Sub Form_Load()
'    Command1_Click
'End Sub



Private Sub Form_Resize()

GraphlitePro1.Width = Me.ScaleWidth - (GraphlitePro1.Left * 2)
GraphlitePro1.Height = Me.ScaleHeight - (GraphlitePro1.top + 120)
GraphlitePro1.Refresh

End Sub



Private Sub Image1_Click()
 Dim cdlg As cdlg
 Dim Color As Long
 Set cdlg = New cdlg
 cdlg.VBChooseColor Color
 If Color <> -1 Then
    GraphlitePro1.SetSeriesOptions 0, Color, "Signal"
    Shape1.FillColor = Color
    Command1_Click
 End If
 Set cdlg = Nothing
 
End Sub

Private Sub Image2_Click()
Dim cdlg As cdlg
 Dim Color As Long
 Set cdlg = New cdlg
 cdlg.VBChooseColor Color
 If Color <> -1 Then
    GraphlitePro1.SetSeriesOptions 1, Color, "Noise"
    Shape2.FillColor = Color
    Command1_Click
 End If
 Set cdlg = Nothing
 
End Sub

Private Sub Option1_Click(Index As Integer)
If Index <> 2 Then
    GraphlitePro1.LowScale = -100
    GraphlitePro1.HighScale = -99
Else
    GraphlitePro1.LowScale = 0
    GraphlitePro1.HighScale = 100
End If
GraphlitePro1.ChartType = Index
GraphlitePro1.Refresh

End Sub
