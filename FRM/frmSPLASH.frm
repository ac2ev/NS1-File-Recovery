VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4125
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSPLASH.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDont 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't show again"
      Height          =   240
      Left            =   5490
      TabIndex        =   7
      Top             =   3870
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pattern Matching Recovery Copyright © Scruge"
      Height          =   240
      Left            =   1590
      TabIndex        =   9
      Top             =   3150
      Width           =   5385
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " And Conversion Utility"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2040
      TabIndex        =   8
      Top             =   630
      Width           =   1950
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "http:\\"
      Height          =   240
      Left            =   1605
      TabIndex        =   6
      Top             =   2850
      Width           =   5385
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSPLASH.frx":000C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   1455
      TabIndex        =   5
      Top             =   1650
      Width           =   5685
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   ""
      Height          =   240
      Left            =   1605
      TabIndex        =   4
      Top             =   2685
      Width           =   5385
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   ""
      Height          =   240
      Left            =   1575
      TabIndex        =   3
      Top             =   2490
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ns1 File Format and Netstumbler icons Copyright © Marius Milner, 2003-2004"
      Height          =   390
      Left            =   1710
      TabIndex        =   2
      Top             =   3615
      Width           =   5385
   End
   Begin VB.Label lblApp 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NS1 File Recovery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1500
      TabIndex        =   1
      Top             =   75
      Width           =   5385
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   615
      Picture         =   "frmSPLASH.frx":00E7
      Stretch         =   -1  'True
      Top             =   480
      Width           =   540
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   585
      Top             =   450
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   465
      Top             =   330
      Width           =   855
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   960
      Left            =   405
      Top             =   270
      Width           =   960
   End
   Begin VB.Label lblVERSION 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1515
      TabIndex        =   0
      Top             =   810
      Width           =   5205
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   4305
      Left            =   0
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constants for topmost.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Enum ONTOPSETTING
    WINDOW_ONTOP = HWND_TOPMOST
    WINDOW_NOT_ONTOP = HWND_NOTOPMOST
End Enum
Dim creg As cRegistry
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Functionality to Set a window always on top or turn it off.
' Date: March,10 1999 @ 10:18:37
'------------------------------------------------------------
Public Sub SetFormOnTop(formHWND As Long, Optional sSETTING As ONTOPSETTING = WINDOW_ONTOP)
    On Error Resume Next
    Call SetWindowPos(formHWND, sSETTING, 0, 0, 0, 0, flags)
End Sub

Private Sub chkDont_Click()
    creg.SetRegistryValue "Show Splash", chkDont.Value, REG_DWORD, , , , False
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Set creg = New cRegistry
        creg.hKey = HKEY_LOCAL_MACHINE
        creg.KeyPath = "Software\NS1"
        chkDont.Value = creg.GetRegistryValue("Show Splash", vbUnchecked, , , , False)
    If chkDont.Value = vbUnchecked Or dontcheckreg = True Then
        lblVERSION = "BETA " & App.Major & "." & App.Minor & "." & App.Revision
        SetFormOnTop Me.hwnd, WINDOW_ONTOP
        Me.Refresh
    Else
        Unload Me
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Set creg = Nothing
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Me.Visible = False
    DoEvents
    frmMain.Show
End Sub

Private Sub Image1_Click()
  Unload Me
End Sub

Private Sub Label1_Click()
 Unload Me
End Sub

Private Sub Label3_Click()
 Unload Me
End Sub

Private Sub Label4_Click()
 Unload Me
End Sub

Private Sub Label5_Click()
 Unload Me
End Sub

Private Sub lblApp_Click()
 Unload Me
End Sub

Private Sub lblVERSION_Click()
 Unload Me
End Sub
