VERSION 5.00
Begin VB.Form frmFTPAdd 
   Caption         =   "LiveUpdate configuration"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDontSee 
      Caption         =   "Check1"
      Height          =   255
      Left            =   150
      TabIndex        =   17
      Top             =   4725
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Caption         =   "Before LiveUpdate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   22
      Top             =   3375
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse ..."
         Enabled         =   0   'False
         Height          =   350
         Left            =   3000
         TabIndex        =   15
         Top             =   690
         Width           =   1095
      End
      Begin VB.CheckBox ChkExecute 
         Caption         =   "Check1"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtExcecute 
         BackColor       =   &H80000004&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblCheck 
         Caption         =   "E&xecute"
         Height          =   255
         Left            =   400
         TabIndex        =   12
         Top             =   380
         Width           =   2175
      End
   End
   Begin VB.Frame fraDownload 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   105
      TabIndex        =   21
      Top             =   2085
      Width           =   4215
      Begin VB.TextBox txtDestination 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   11
         Text            =   "c:\MyRep\file.xxx"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtDir 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   9
         Text            =   "/pub/users/"
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "&Destination :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "&FTP Directory :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "FTP config"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "NS1 Recovery/Conversion Program"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "ftp.frontiernet.net"
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "user@host.com"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "anonymous"
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "&Name :"
         Height          =   255
         Left            =   75
         TabIndex        =   0
         Top             =   405
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "FTP &Server :"
         Height          =   255
         Left            =   75
         TabIndex        =   2
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "&Password :"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   1485
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "&User :"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   1125
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3030
      TabIndex        =   19
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1710
      TabIndex        =   18
      Top             =   5085
      Width           =   1215
   End
   Begin VB.Label lblDontSee 
      Caption         =   "See &config button next time"
      Height          =   255
      Left            =   465
      TabIndex        =   16
      Top             =   4755
      Width           =   2295
   End
End
Attribute VB_Name = "frmFTPAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkExecute_Click()
If Me.ChkExecute = 1 Then
   Me.cmdBrowse.Enabled = True
   Me.txtExcecute.BackColor = 16777215
   Me.txtExcecute.Locked = False
Else
   Me.cmdBrowse.Enabled = False
   Me.txtExcecute.BackColor = 12632256
   Me.txtExcecute.Locked = True
End If
End Sub

Private Sub cmdBrowse_Click()
Dim strFiles As String
strFiles = GetOpenFile(CurDir, "Choisir un fichier", True, False)
If strFiles <> "" Then
   Me.txtExcecute = strFiles
End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  frmLiveUpdate.creg.SetRegistryValue "USER", txtUser, REG_SZ, , , , False
  FTP_User = txtUser
  
  frmLiveUpdate.creg.SetRegistryValue "HOST", txtHost, REG_SZ, , , , False
  FTP_Server = txtHost
  
  frmLiveUpdate.creg.SetRegistryValue "PASS", Encrypte(txtPass), REG_SZ, , , , False
  FTP_Pass = Encrypte(txtPass)
  
  frmLiveUpdate.creg.SetRegistryValue "Remote DIR", txtDir, REG_SZ, , , , False
  strDir = txtDir
  
  frmLiveUpdate.creg.SetRegistryValue "Local DIR", txtDestination, REG_SZ, , , , False
  strPath = txtDestination
    
  frmLiveUpdate.creg.SetRegistryValue "EXECUTE", ChkExecute, REG_SZ, , , , False
  intExecute = ChkExecute
  
  frmLiveUpdate.creg.SetRegistryValue "EXECFILES", txtExcecute, REG_SZ, , , , False
  strExecute = txtexecute
  
  frmLiveUpdate.creg.SetRegistryValue "SEECONFIG", chkDontSee, REG_DWORD, , , , False
  frmLiveUpdate.cmdConfig.Visible = chkDontSee
  
  
  frmLiveUpdate.Caption = "LiveUpdate " & txtName
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  txtUser = FTP_User
  txtHost = FTP_Server
  txtPass = FTP_Pass
  txtFiles = strDir
  txtDestination = strPath
  ChkExecute = intExecute
  txtExcecute = strExecute
 ' chkDontSee.Value = frmLiveUpdate.chkConfig.Value
End Sub

Private Sub lblCheck_Click()
  If Me.ChkExecute = 1 Then
     Me.ChkExecute = 0
  Else
     Me.ChkExecute = 1
  End If
End Sub

Private Sub lblDontSee_Click()
  If Me.chkDontSee = 1 Then
     Me.chkDontSee = 0
  Else
     Me.chkDontSee = 1
  End If
End Sub
