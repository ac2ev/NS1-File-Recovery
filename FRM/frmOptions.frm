VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   9450
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5220
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   30
      Top             =   9180
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9155
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Google Maps .KML"
      Height          =   7185
      Left            =   135
      TabIndex        =   11
      Top             =   1440
      Width           =   5025
      Begin VB.Frame Frame6 
         Caption         =   "3D Visualization"
         Height          =   645
         Left            =   240
         TabIndex        =   46
         Top             =   5490
         Width           =   3975
         Begin VB.CheckBox chk3D 
            Caption         =   "Include"
            Height          =   255
            Left            =   150
            TabIndex        =   47
            Top             =   270
            Width           =   2250
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Group By"
         Height          =   1005
         Left            =   225
         TabIndex        =   40
         Top             =   285
         Width           =   4170
         Begin VB.OptionButton optGroup 
            Caption         =   "Mode (AP, Peer)"
            Height          =   255
            Index           =   4
            Left            =   2220
            TabIndex        =   45
            Top             =   435
            Width           =   1830
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "Encryption"
            Height          =   255
            Index           =   3
            Left            =   2220
            TabIndex        =   44
            Top             =   195
            Width           =   1830
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "Channel (B/G channels 1-14)"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   675
            Width           =   3780
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "SSID"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   450
            Width           =   1830
         End
         Begin VB.OptionButton optGroup 
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   41
            Top             =   225
            Width           =   1830
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   240
         TabIndex        =   36
         Top             =   1245
         Width           =   4125
         Begin VB.OptionButton optSSID 
            Caption         =   "Use BSSID for description field"
            Height          =   270
            Index           =   0
            Left            =   75
            TabIndex        =   39
            Top             =   240
            Value           =   -1  'True
            Width           =   2595
         End
         Begin VB.OptionButton optSSID 
            Caption         =   "Use SSID for description field"
            Height          =   270
            Index           =   1
            Left            =   75
            TabIndex        =   38
            Top             =   510
            Width           =   2595
         End
         Begin VB.OptionButton optSSID 
            Caption         =   "Leave Blank (Breadcrumb)"
            Height          =   270
            Index           =   2
            Left            =   75
            TabIndex        =   37
            Top             =   780
            Width           =   2595
         End
      End
      Begin VB.Frame Frame3 
         Height          =   810
         Left            =   240
         TabIndex        =   31
         Top             =   6225
         Width           =   2565
         Begin VB.PictureBox picIcons 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   1
            Left            =   105
            ScaleHeight     =   480
            ScaleWidth      =   465
            TabIndex        =   35
            ToolTipText     =   "Ad-Hoc No Encryption"
            Top             =   195
            Width           =   495
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   3
            Left            =   1335
            ScaleHeight     =   480
            ScaleWidth      =   465
            TabIndex        =   34
            ToolTipText     =   "AP No Encryption"
            Top             =   195
            Width           =   495
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   2
            Left            =   720
            ScaleHeight     =   480
            ScaleWidth      =   465
            TabIndex        =   33
            ToolTipText     =   "Ad-Hoc Encryption"
            Top             =   195
            Width           =   495
         End
         Begin VB.PictureBox picIcons 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   510
            Index           =   4
            Left            =   1950
            ScaleHeight     =   480
            ScaleWidth      =   465
            TabIndex        =   32
            ToolTipText     =   "AP Encryption"
            Top             =   195
            Width           =   495
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data to Export"
         Height          =   2910
         Left            =   240
         TabIndex        =   12
         Top             =   2520
         Width           =   4020
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   16
            Left            =   2115
            TabIndex        =   29
            Top             =   2565
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   15
            Left            =   2115
            TabIndex        =   28
            Top             =   2280
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   14
            Left            =   2115
            TabIndex        =   27
            Top             =   1995
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   13
            Left            =   2115
            TabIndex        =   26
            Top             =   1725
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   12
            Left            =   2115
            TabIndex        =   25
            Top             =   1440
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   11
            Left            =   2115
            TabIndex        =   24
            Top             =   1155
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   10
            Left            =   2115
            TabIndex        =   23
            Top             =   885
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   9
            Left            =   2115
            TabIndex        =   22
            Top             =   600
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   8
            Left            =   2115
            TabIndex        =   21
            Top             =   315
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   7
            Left            =   180
            TabIndex        =   20
            Top             =   2280
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   6
            Left            =   180
            TabIndex        =   19
            Top             =   1995
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   5
            Left            =   180
            TabIndex        =   18
            Top             =   1725
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   17
            Top             =   1440
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   3
            Left            =   180
            TabIndex        =   16
            Top             =   1155
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   15
            Top             =   885
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   14
            Top             =   600
            Width           =   1395
         End
         Begin VB.CheckBox chkExport 
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   13
            Top             =   315
            Width           =   1395
         End
      End
   End
   Begin VB.CheckBox chkStrip 
      Caption         =   "Strip out Non Printable Characters from SSID"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   1080
      Width           =   4845
   End
   Begin VB.TextBox txtNoSSID 
      Height          =   270
      Left            =   180
      TabIndex        =   8
      Top             =   645
      Width           =   4875
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   6
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   5
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3975
      TabIndex        =   1
      Top             =   8730
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2745
      TabIndex        =   0
      Top             =   8730
      Width           =   1095
   End
   Begin VB.Label lblNoSSID 
      Caption         =   $"frmOptions.frx":000C
      Height          =   390
      Left            =   180
      TabIndex        =   9
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lIndex As Integer
Dim Temp As Integer
Public IconIndex As Integer

Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload frmThumbs
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Descriptor = Temp
    creg.SetRegistryValue "NoSSID", txtNoSSID.Text, REG_SZ, , , True, False
    creg.SetRegistryValue "NoPrint", chkStrip.Value, REG_DWORD, , , True, False
    creg.SetRegistryValue "KMLDescriptor", Descriptor, REG_DWORD, , , True, False
    creg.SetRegistryValue "ExportItems", ExportItem, REG_binary, , , True, False
    creg.SetRegistryValue "GroupBy", GroupBy, REG_DWORD, , , True, False
    creg.SetRegistryValue "3D", chk3D.Value, REG_DWORD, , , True, False
    Use3D = chk3D.Value
    
    'creg.SetRegistryValue "Version", txtVersion.Text, REG_SZ, , , , False
    ReDim ExportItem(17)
    For lIndex = 0 To 16
     ExportItem(lIndex) = chkExport(lIndex).Value
    Next lIndex
    NoSSID = txtNoSSID.Text
    NoPrint = chkStrip.Value
    
    Unload frmThumbs
    Unload Me
    
End Sub

Private Sub cmdThumbnail_Click()
 frmThumbs.Show 1
End Sub

Private Sub Form_Load()
    'center the forms
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    txtNoSSID.Text = NoSSID
    chkStrip.Value = NoPrint
    On Error Resume Next
    
    For lIndex = 0 To 16
        chkExport(lIndex).Value = ExportItem(lIndex)
    Next lIndex
    optSSID(Descriptor).Value = True
    optGroup(GroupBy).Value = True
    
    chkExport(0).Caption = "SSID"
    chkExport(1).Caption = "BSSID"
    chkExport(2).Caption = "Type"
    chkExport(3).Caption = "Latitude"
    chkExport(4).Caption = "Longitude"
    chkExport(5).Caption = "FirstSeen"
    chkExport(6).Caption = "LastSeen"
    chkExport(7).Caption = "SNR"
    chkExport(8).Caption = "Name"
    chkExport(9).Caption = "Flags"
    chkExport(10).Caption = "Channels"
    chkExport(11).Caption = "Beacon"
    chkExport(12).Caption = "DataRate"
    chkExport(13).Caption = "Channel"
    chkExport(14).Caption = "IP Address"
    chkExport(15).Caption = "IP Mask "
    chkExport(16).Caption = "IP SubNet"
    IconPath = creg.GetRegistryValue("IconPath", App.Path & "\icons")
           
     If IconPath = App.Path & "\icons" And Not FileExists(App.Path & "\icons", vbDirectory) Then
        MkDir App.Path & "\icons"
        If Not FileExists(App.Path & "\icons\100.PNG") Then ExtractRes App.Path & "\icons", 100, "PNG"
        If Not FileExists(App.Path & "\icons\101.PNG") Then ExtractRes App.Path & "\icons", 101, "PNG"
        If Not FileExists(App.Path & "\icons\102.PNG") Then ExtractRes App.Path & "\icons", 102, "PNG"
        If Not FileExists(App.Path & "\icons\103.PNG") Then ExtractRes App.Path & "\icons", 103, "PNG"
        If Not FileExists(App.Path & "\icons\104.PNG") Then ExtractRes App.Path & "\icons", 104, "PNG"
        If Not FileExists(App.Path & "\icons\105.PNG") Then ExtractRes App.Path & "\icons", 105, "PNG"
     End If

     picAdHocNW = creg.GetRegistryValue("picAdHocNW", App.Path & "\icons\100.png", , , , False)
     picAdHocW = creg.GetRegistryValue("picAdHocW", App.Path & "\icons\101.png", , , , False)
     picAPNW = creg.GetRegistryValue("picAPNW", App.Path & "\icons\100.png", , , , False)
     picAPW = creg.GetRegistryValue("picAPW", App.Path & "\icons\101.png", , , , False)
    'txtVersion.TabIndex = creg.GetRegistryValue("Version", "0.5.7", , , , False)
     chk3D.Value = creg.GetRegistryValue("3D", 1, , , , False)
     LoadPicture picIcons(1), picAdHocNW
     LoadPicture picIcons(2), picAdHocW
     LoadPicture picIcons(3), picAPNW
     LoadPicture picIcons(4), picAPW
    
   
End Sub



Private Sub LoadPicture(ByRef pic As PictureBox, Filename As String)
    Dim png As New LoadPNG
    Dim Test As Long
   
   On Error GoTo ErrorHandler
   Me.Refresh
    png.PicBox = pic
    png.SetOwnBkgndColor False
    png.SetAlpha = False
    png.SetTrans = False


Test = png.OpenPNG(Filename)

If png.ErrorNumber <> 0 Then MsgBox "Error loading picture" & png.ErrorNumber
 Exit Sub
'**********************
ErrorHandler:
   MsgBox Err.Description, vbExclamation
   Exit Sub
End Sub


Private Sub Form_Unload(Cancel As Integer)
 Unload frmThumbs
End Sub

Private Sub optGroup_Click(Index As Integer)
    GroupBy = Index
End Sub

Private Sub optSSID_Click(Index As Integer)
 Temp = Index
End Sub


Private Sub picIcons_Click(Index As Integer)
    If Index = 1 Then StatusBar1.Panels(1).Text = "Loading....Ad-Hoc No WEP"
    If Index = 2 Then StatusBar1.Panels(1).Text = "Loading....Ad-Hoc WEP"
    If Index = 3 Then StatusBar1.Panels(1).Text = "Loading....AP No WEP"
    If Index = 4 Then StatusBar1.Panels(1).Text = "Loading....AP WEP"
    
    IconIndex = Index
    frmThumbs.Top = picIcons(Index).Top
    frmThumbs.Left = picIcons(Index).Left + picIcons(Index).Width + 10
    frmThumbs.LastPath = IconPath
    
    frmThumbs.Show 1
    If frmThumbs.PicPath = "" Then Exit Sub
    
    If Index = 1 Then
        creg.SetRegistryValue "picAdHocNW", frmThumbs.PicPath, REG_SZ, , , True, False
        picAdHocNW = frmThumbs.PicPath
        LoadPicture picIcons(1), picAdHocNW
    End If
    
    If Index = 2 Then
        creg.SetRegistryValue "picAdHocW", frmThumbs.PicPath, REG_SZ, , , True, False
        picAdHocW = frmThumbs.PicPath
        LoadPicture picIcons(2), picAdHocW
    End If
    
    If Index = 3 Then
        creg.SetRegistryValue "picAPNW", frmThumbs.PicPath, REG_SZ, , , True, False
        picAPNW = frmThumbs.PicPath
        LoadPicture picIcons(3), picAPNW
    End If
    
    If Index = 4 Then
        creg.SetRegistryValue "picAPW", frmThumbs.PicPath, REG_SZ, , , True, False
        picAPW = frmThumbs.PicPath
        LoadPicture picIcons(4), picAPW
    End If
    IconPath = frmThumbs.LastPath
End Sub

