VERSION 5.00
Begin VB.Form frmLiveUpdate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LiveUpdate"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLiveUpdate 
      Caption         =   "Enable Live Updates"
      Height          =   270
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Each time the program is loaded it will check to see if there is a new version"
      Top             =   3435
      Width           =   1845
   End
   Begin VB.CommandButton cmdConfig 
      Caption         =   "Config ..."
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame frmTransfert 
      Caption         =   "Progress"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Visible         =   0   'False
      Width           =   4455
      Begin VB.Label Lbl_Averages 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label txtPercent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   525
         Width           =   4215
      End
      Begin VB.Label Lbl_FileSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label Percent 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   140
         TabIndex        =   9
         Top             =   560
         Width           =   15
      End
      Begin VB.Label lbl_Time 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   3735
      End
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   360
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   6735
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   3270
      Left            =   120
      Picture         =   "frmUpdate.frx":0CCA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblEnd 
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3200
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Height          =   855
      Left            =   2520
      TabIndex        =   5
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblCaption 
      Caption         =   "Bienvenue à la mise à jour internet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2445
      TabIndex        =   4
      Top             =   195
      Width           =   4215
   End
End
Attribute VB_Name = "frmLiveUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim intStep As Integer
Dim hOpen As Long, hConnection As Long
Dim AppName As String
Dim blnNewUpdate As Boolean
Dim InProgress As Boolean
Public creg As cRegistry

Private Sub chkLiveUpdate_Click()
 creg.SetRegistryValue "LiveUpdate", chkLiveUpdate.Value, REG_DWORD, , "Software\NS1", , False
End Sub

Private Sub cmdBack_Click()
  intStep = intStep - 1
  Step (intStep)
End Sub

Private Sub cmdCancel_Click()
  If InProgress = True Then
     If MsgBox("Do you really want cancel the update ?", vbYesNo + vbDefaultButton2, "LiveUpdate") = vbYes Then
        StopTransfert = True
        frmTransfert.Visible = False
        'Unload Me
     End If
  Else
     Unload Me
  End If
     
End Sub

Private Sub cmdConfig_Click()
  frmFTPAdd.Show vbModal
End Sub

Private Sub cmdNext_Click()

  If Me.cmdNext.Caption = "Finish" Then
     'Program to execute when live update is finished
     If blnNewUpdate = True Then 'If update is downloaded

        intExecute = creg.GetRegistryValue("EXECUTE", 0, , , , False)
        strExecute = creg.GetRegistryValue("EXECFILES", App.Path & "\Updates\ns1Recovery.exe", , , , False)
        
        If intExecute = 1 And Dir(strExecute) <> "" Then
           Call Shell(strExecute, vbNormalFocus)
        End If
        Call Shell("explorer " & strPath, vbNormalFocus)
     End If
     
     Unload Me
  Else
     intStep = intStep + 1
     Step (intStep)
  End If
End Sub

Private Sub Form_Load()
    Set creg = New cRegistry
    creg.hKey = HKEY_LOCAL_MACHINE
    creg.KeyPath = "Software\NS1\Live Update"
    If Not creg.DoesKeyExist Then creg.CreateKey
    myVer = creg.GetRegistryValue("Version", "0.5.7", , , , False)
    chkLiveUpdate.Value = creg.GetRegistryValue("LiveUpdate", False, , "Software\NS1", , False)
    Me.cmdConfig.Visible = creg.GetRegistryValue("AllowConfig", 0, , , , False)
    AppName = "Ns1KFrac"
    intStep = intStep + 1
    Step (intStep)
    frmLiveUpdate.Caption = "LiveUpdate " & AppName
      FTP_Server = creg.GetRegistryValue("Host", "ftp.frontiernet.net", , , , False)
     FTP_User = creg.GetRegistryValue("User", "anonymous", , , , False)
    FTP_Pass = Decrypt(creg.GetRegistryValue("Pass", "}o|Jry}~8myw", , , , False)) 'user@host.com
        strDir = creg.GetRegistryValue("Remote DIR", "/pub/users/", , , , False)
    strFile = creg.GetRegistryValue("File", "ns1Recovery.exe", , , , False)
    strPath = creg.GetRegistryValue("Local DIR", App.Path & "\Updates\", , , , False)
    If Dir(strPath, vbDirectory) = "" Then MkDir (strPath)
End Sub

Private Sub Step(Number As Integer)
Dim blnUpdate As Boolean 'Si il ya eu mise à jour

Select Case Number
Case 1   'Form load
     Me.lblCaption.Caption = "Welcome to the LiveUpdate of " & vbCrLf & AppName
     Me.lblInfo.Caption = "LiveUpdate checks for more recents update available on Internet."
     Me.lblEnd.Caption = "Click Next to continue LiveUpdate"
Case 2
     Me.lblCaption.Caption = "Internet Connection"
     Me.lblInfo = "LiveUpdate goes connected on the Internet server." & vbCrLf & vbCrLf & "If you are not connected, LiveUpdate goes open automatically your Internet connection." & vbCrLf
     Me.lblEnd.Caption = "Press Next to connect."
     
Case 3
     Me.cmdNext.Enabled = False
     Me.cmdBack.Enabled = False
     If FTPConnect = True Then
        Me.cmdNext.Enabled = True
        Me.lblCaption.Caption = "You are conected on the Server"
        Me.lblInfo = "Now LiveUpdate while verify if a most recent update available on Internet."
        Me.lblEnd.Caption = "Press Next to verify update"
     Else
        frmLiveUpdate.lblEnd.Caption = "Impossible to connect"
        Me.lblCaption = "Connection error!"
        Me.lblInfo = ""
        
        Me.cmdNext.Enabled = True
        Me.cmdNext.Caption = "Finish"
        Me.cmdBack.Enabled = False
        Me.cmdCancel.Enabled = False
        Exit Sub
     End If
Case 4
     blnUpdate = SendFiles("LiveUpdate")
     If blnUpdate = False Then
        Me.cmdNext.Caption = "Finish"
        Me.cmdBack.Enabled = False
        Me.cmdCancel.Enabled = False
        Me.lblCaption.Caption = "Thank you for using LiveUpdate"
        Me.lblInfo = "No more recent update was available at this moment."
        Me.lblEnd.Caption = "Press Finish to quit"
     Else
        Me.lblCaption.Caption = "Thank you for using LiveUpdate"
        Me.lblInfo.Caption = "LiveUpdate success!, Please replace current file with downloaded file"
     End If
End Select

If Number > 1 And Number < 4 Then
   Me.cmdBack.Enabled = True
Else
   Me.cmdBack.Enabled = False
End If
End Sub

Public Function FTPConnect() As Boolean
  Screen.MousePointer = 11

  
Dim nFlag As Long
    'MousePointer = vbHourglass
        frmLiveUpdate.lblEnd.Caption = "Connecting..."
        DoEvents
    hOpen = InternetOpen(FTP_UAgent, INTERNET_OPEN_TYPE_DIRECT, _
                         vbNullString, vbNullString, 0)
    
    If hOpen <> 0 Then
      hConnection = InternetConnect(hOpen, FTP_Server, _
                                    INTERNET_INVALID_PORT_NUMBER, _
                                    FTP_User, _
                                    FTP_Pass, _
                                    INTERNET_SERVICE_FTP, nFlag, 0)
   
      If hConnection <> 0 Then
         frmLiveUpdate.lblEnd.Caption = "Connected"
         FTPConnect = True
      Else
         FTPConnect = False
      End If
    Else
       FTPConnect = False
    End If
    Screen.MousePointer = 0
End Function

Private Sub Form_Unload(Cancel As Integer)
  Call InternetCloseHandle(hConnection)
  Call InternetCloseHandle(hOpen)
End Sub

Public Function SendFiles(vName As String) As Boolean
Dim x            As Integer

Dim SizeFile     As Long
Dim pData        As WIN32_FIND_DATA
Dim hFile        As Long
Dim hRet         As Long
Dim lTime        As FILETIME
Dim sTime        As SYSTEMTIME
Dim strFTPTime   As String
Dim strLocalTime As String
Dim intNbrPasse  As Integer
Dim BlnResult    As Boolean
Me.Timer.Enabled = True
intNbrPasse = 0
Me.lblEnd.Caption = "Check update...."
'For x = 1 To 99

    
    If strFile <> "" Then
       Call ResetPB   'Reset la progress bar
       pData.cFileName = String(MAX_PATH, 0)
       hFile = FtpFindFirstFile(hConnection, Trim(strDir & strFile), pData, 0, 0)
       If hFile = 0 Then GoTo ExitRoutine
       hRet = InternetFindNextFile(hFile, pData)
       SizeFile = pData.nFileSizeLow
       glbSize = SizeFile
       lTime = pData.ftLastWriteTime
       
       l = FileTimeToSystemTime(lTime, sTime)
       strFTPTime = GetFileDateString(pData.ftLastWriteTime) 'Date Heure du fichier FTP
       strLocalTime = RetFileDate(App.Path & "\" & strFile)   'Date Heure du fichier Local
       If strFTPTime <= strLocalTime Then
          Me.lblEnd.Caption = "No update available" & strFile
       Else
          frmTransfert.Visible = True
          intNbrPasse = intNbrPasse + 1
          BlnResult = GetFiles(strDir & strFile, strPath & strFile, SizeFile, 1)    'Si le transfert à réussi
                    
       End If
       strFiles = ""
       Call InternetCloseHandle(hFile)
       Call InternetCloseHandle(hRet)
       hFile = 0
       hRet = 0
    Else
       Me.cmdCancel.Caption = "Close"
       Me.Timer.Enabled = False
       Me.lbl_Time.Caption = ""
    End If
'Next x
ExitRoutine:

If intNbrPasse = 0 Then
   Me.lblEnd.Caption = "You have most recent version"
   Me.cmdBack.Enabled = False
   Me.cmdCancel.Enabled = False
   Me.cmdNext.Caption = "Close"
   SendFiles = False
Else
   If BlnResult = True Then
      Me.lblEnd.Caption = "Update Complete " & intNbrPasse & " file(s) updated !"
      blnNewUpdate = True
      Me.cmdBack.Enabled = False
      Me.cmdCancel.Enabled = False
      Me.cmdNext.Caption = "Finish"
      SendFiles = True
   Else
      
   End If
End If

End Function

Private Function ResetPB()
'Reset la Progress Bar
  Me.Percent.Width = 15
  Me.txtPercent.Caption = ""

  Me.Lbl_FileSize.Caption = ""
  Me.Lbl_Averages.Caption = ""
  Me.lbl_Time.Caption = ""
End Function

Public Function GetFiles(strFile As String, strNewFile As String, lngFileSize As Long, vMode As Integer) As Boolean
   Dim hFile                  As Long
   Dim sBuffer                As String
   Dim sReadBuffer            As String * 4096 'par tranche de 4k
   Dim lNumberOfBytesRead     As Long
   Dim bDoLoop                As Boolean
   Dim Sum                    As Long
   Dim x                      As Integer
   GetFiles = True
   
   If vMode = 0 Then  'Mode de transfert des données
       Transfer = FTP_TRANSFER_TYPE_ASCII
   Else
       Transfer = FTP_TRANSFER_TYPE_BINARY
   End If
   InProgress = True
   hFile = FtpOpenFile(hConnection, Trim(strFile), GENERIC_READ, Transfer, 0)
   Open strNewFile For Binary Access Write As #2
   
   bDoLoop = True
   StopTransfert = False
   
   While bDoLoop
      DoEvents
      If StopTransfert = True Then
         Close #2
         Kill strNewFile
         
         For x = 1 To 10000
             DoEvents
         Next x
         GetFiles = False
         Call ResetPB
         GoTo StopGetFiles
      End If
      
      sReadBuffer = vbNullChar
      bDoLoop = InternetReadFile(hFile, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
      sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
      If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
      Sum = Sum + lNumberOfBytesRead
      Call ProgressBar(lngFileSize, Str(Sum), strFile)
      Put #2, , sBuffer
      sBuffer = ""
   Wend
         
StopGetFiles:
   Close #2
   InternetCloseHandle (hFile)
   
End Function


Private Sub Timer_Timer()
  Dim Nbrk As Long
  Nbrk = DoneBytes - OldBytes
  If Nbrk > 0 Then
     Lbl_Averages.Refresh
     lbl_Time.Refresh
     Lbl_Averages.Caption = "Average : " & Format(Nbrk / 1024, "###0.0") & " / Kbps"
     lbl_Time.Caption = ConvSeconde(((glbSize - DoneBytes) / (Nbrk / 1024) / 1024))
  End If
  OldBytes = DoneBytes
End Sub

Public Function ProgressBar(Size, Done, Files)
'Affiche la bar de progression
  If Done = 0 Then Exit Function
  Dim iSendPercent As Integer
  Dim x            As Integer

  iSendPercent = (Done / Size) * 100
  If iSendPercent >= 50 Then
     frmLiveUpdate.txtPercent.ForeColor = 16777215
  Else
     frmLiveUpdate.txtPercent.ForeColor = 0
  End If

  DoneBytes = Done
  frmLiveUpdate.frmTransfert.Caption = "Transfer " & Trim(Files)
  frmLiveUpdate.Percent.Width = 41.5 * iSendPercent
  frmLiveUpdate.Percent.Caption = iSendPercent & " %"
  frmLiveUpdate.Percent.Refresh
  frmLiveUpdate.Lbl_FileSize.Caption = Format(Done / 1000, "###0.0") & "Kb / " & Format(Size / 1000, "###0.0") & " Kb"
  DoEvents
End Function
