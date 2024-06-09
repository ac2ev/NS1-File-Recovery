VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Dialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Help"
   ClientHeight    =   7710
   ClientLeft      =   2775
   ClientTop       =   3675
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Tag             =   "0011"
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7515
      Left            =   1410
      TabIndex        =   0
      Tag             =   "0011"
      Top             =   150
      Width           =   10890
      ExtentX         =   19209
      ExtentY         =   13256
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
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
   Begin VB.Image Image1 
      Height          =   555
      Left            =   615
      Picture         =   "Dialog.frx":0000
      Stretch         =   -1  'True
      Tag             =   "1100"
      Top             =   473
      Width           =   540
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   578
      Tag             =   "1100"
      Top             =   443
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   458
      Tag             =   "1100"
      Top             =   323
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   405
      Tag             =   "1100"
      Top             =   270
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   8355
      Left            =   -15
      Tag             =   "0011"
      Top             =   -690
      Width           =   1095
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Resizer As New ControlResizer

Option Explicit

Private Sub Form_Load()
Resizer.InitResizer Me, Me.Width, Me.Height
Resizer.InitResizer Me, Me.Width, Me.Height
   WebBrowser1.Navigate App.Path & "\DOC\Help.htm"
   
End Sub

Private Sub Form_Resize()
Resizer.FormResized Me
Resizer.FormResized Me

End Sub
