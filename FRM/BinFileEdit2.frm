VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBinFileEdit2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary file edit"
   ClientHeight    =   7365
   ClientLeft      =   1605
   ClientTop       =   1530
   ClientWidth     =   10425
   Icon            =   "BinFileEdit2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   695
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picHeader 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10365
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   10425
      Begin VB.CommandButton cmdDummy 
         Height          =   375
         Index           =   3
         Left            =   7920
         TabIndex        =   37
         Top             =   0
         Width           =   45
      End
      Begin VB.CommandButton cmdDummy 
         Height          =   375
         Index           =   2
         Left            =   6750
         TabIndex        =   36
         Top             =   -30
         Width           =   45
      End
      Begin VB.CommandButton cmdDummy 
         Height          =   375
         Index           =   1
         Left            =   2340
         TabIndex        =   35
         Top             =   -30
         Width           =   45
      End
      Begin VB.CommandButton cmdDummy 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   34
         Top             =   0
         Width           =   45
      End
      Begin VB.CommandButton cmdRight 
         Height          =   345
         Left            =   9480
         Picture         =   "BinFileEdit2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "Right one character"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdLeft 
         Height          =   345
         Left            =   9180
         Picture         =   "BinFileEdit2.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Left one character"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdPrintPage 
         Height          =   345
         Left            =   690
         Picture         =   "BinFileEdit2.frx":090E
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Print current page"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdSave 
         Height          =   345
         Left            =   360
         Picture         =   "BinFileEdit2.frx":0A58
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   315
      End
      Begin VB.Frame fraEditContainer 
         Height          =   435
         Left            =   1740
         TabIndex        =   29
         Top             =   -90
         Width           =   345
         Begin VB.Image imgEdit 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   30
            Picture         =   "BinFileEdit2.frx":0BA2
            Top             =   120
            Width           =   270
         End
      End
      Begin VB.CommandButton cmdGoTo 
         Height          =   345
         Left            =   7590
         Picture         =   "BinFileEdit2.frx":0EE4
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Go to"
         Top             =   0
         Width           =   315
      End
      Begin VB.TextBox txbSearch 
         BackColor       =   &H00FFFFC0&
         Height          =   345
         Left            =   4320
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "txtSearch"
         Top             =   0
         Width           =   1665
      End
      Begin VB.TextBox txbGoTo 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   6780
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "txtGoTo"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdSearchFromStart 
         Height          =   345
         Left            =   5970
         Picture         =   "BinFileEdit2.frx":10AE
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Search from start"
         Top             =   0
         Width           =   345
      End
      Begin VB.CheckBox ckbCaseSensitive 
         Caption         =   "Case"
         Height          =   195
         Left            =   3630
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Case sensitive"
         Top             =   60
         Value           =   1  'Checked
         Width           =   675
      End
      Begin VB.CommandButton cmdFileOpen 
         Height          =   345
         Left            =   0
         Picture         =   "BinFileEdit2.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Open file"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdExit 
         Height          =   345
         Left            =   1350
         Picture         =   "BinFileEdit2.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Exit"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdHelp 
         Height          =   345
         Left            =   1020
         Picture         =   "BinFileEdit2.frx":150C
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdSearchOrFindNext 
         Height          =   345
         Left            =   6330
         Picture         =   "BinFileEdit2.frx":1656
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Search / find next"
         Top             =   0
         Width           =   345
      End
      Begin VB.CommandButton cmdPgUp 
         Height          =   345
         Left            =   8310
         Picture         =   "BinFileEdit2.frx":17A0
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Page Up"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdDn 
         Height          =   345
         Left            =   8580
         Picture         =   "BinFileEdit2.frx":1AE2
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "If in view mode: down one line.  If in edit mode: down one character."
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdPgDn 
         Height          =   345
         Left            =   8010
         Picture         =   "BinFileEdit2.frx":1E24
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Page Down"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdUp 
         Height          =   345
         Left            =   8880
         Picture         =   "BinFileEdit2.frx":2166
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "If in view mode: up one line.  If in edit mode: up one character."
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   345
         Left            =   9780
         Picture         =   "BinFileEdit2.frx":24A8
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "First page"
         Top             =   0
         Width           =   285
      End
      Begin VB.CommandButton cmdLast 
         Height          =   345
         Left            =   10080
         Picture         =   "BinFileEdit2.frx":27EA
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Last page"
         Top             =   0
         Width           =   285
      End
      Begin VB.OptionButton OptSearch 
         Caption         =   "Hex"
         Height          =   195
         Index           =   0
         Left            =   2430
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   60
         Width           =   615
      End
      Begin VB.OptionButton OptSearch 
         Caption         =   "Chr"
         Height          =   195
         Index           =   1
         Left            =   3030
         TabIndex        =   12
         Top             =   60
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.Image imgOverWriteOn 
         Height          =   240
         Left            =   2100
         Picture         =   "BinFileEdit2.frx":2B2C
         ToolTipText     =   "Overwrite On (Use popup menu to change)"
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgOverwriteOff 
         Height          =   240
         Left            =   2100
         Picture         =   "BinFileEdit2.frx":2E6E
         ToolTipText     =   "Overwrite Off (Use popup menu to change)"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      Height          =   6255
      Left            =   330
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   653
      TabIndex        =   1
      Top             =   510
      Width           =   9855
      Begin VB.PictureBox picHexDisp 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5835
         Left            =   1500
         ScaleHeight     =   385
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   399
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   6045
         Begin VB.TextBox txbEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Terminal"
               Size            =   9
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   225
            Left            =   0
            MaxLength       =   2
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "txbEdit"
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.PictureBox picOffset2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1500
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   401
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   6045
      End
      Begin VB.PictureBox picChrDisp 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5835
         Left            =   7740
         ScaleHeight     =   385
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   3
         Top             =   330
         Width           =   2055
      End
      Begin VB.PictureBox picOffSet1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   5835
         Left            =   0
         ScaleHeight     =   387
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   85
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1305
      End
      Begin RichTextLib.RichTextBox rtbChr 
         Height          =   5835
         Left            =   30
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Visible         =   0   'False
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   10292
         _Version        =   393217
         HideSelection   =   0   'False
         ScrollBars      =   3
         TextRTF         =   $"BinFileEdit2.frx":31B0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFileSize 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblFileSize"
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "File size"
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label lblAscii 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblAscii"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   7770
         TabIndex        =   8
         ToolTipText     =   "ASCII"
         Top             =   0
         Width           =   705
      End
      Begin VB.Label lblBinary 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblBinary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   8490
         TabIndex        =   7
         ToolTipText     =   "Binary"
         Top             =   0
         Width           =   1305
      End
   End
   Begin VB.Image imgNoMode 
      Height          =   240
      Left            =   60
      Picture         =   "BinFileEdit2.frx":3238
      Top             =   4710
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgEditMode 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   30
      Picture         =   "BinFileEdit2.frx":357A
      Top             =   5040
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgViewMode 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Left            =   30
      Picture         =   "BinFileEdit2.frx":38BC
      Top             =   5370
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblFileSpec 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblFileSpec"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   330
      TabIndex        =   0
      Top             =   6810
      Width           =   9885
   End
   Begin VB.Image imgHerman 
      Height          =   1305
      Left            =   30
      Picture         =   "BinFileEdit2.frx":3BC6
      Top             =   5820
      Width           =   240
   End
   Begin VB.Menu mnuPopOverWrite 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu PopUpOverWrite 
         Caption         =   "Overwrite in edit"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmBinFileEdit2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' BinFileEdit2.frm
' (This version is for a faster machine.  If you have a slower CPU, use the other version:
' BinFileEdit1)
'
' By Herman Liu
'
' View/edit binary and text files, with both hex and character search facilities
' fully functional, and you can print any displayed page (upto 512 bytes, showing
' byte positions, hex and characters).
'
' Note: Terminal font is used because it can display most of the characters.  Others
' such as Courier or FixedSys cannot print tens of characters after ASCII 127.
'
Option Explicit

Private Const EM_GETSEL = &HB0

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    
Private Const CharsInRow = 16
Private Const CharsInCol = 32
Private Const mPageSize = CharsInRow * CharsInCol
Dim mFileSize As Long
Dim arrByte() As Byte
Dim arrSearchByte() As Byte
Dim pageStart As Long
Dim pageEnd As Long
Dim origHex As String

Dim StdW1 As Long
Dim StdH1 As Long
Dim StdW2 As Long
Dim StdH2 As Long
Dim ChrW As Long

Dim mSuspend As Boolean
Dim prevFoundPos As Long
Dim mDirty As Boolean
Dim gCancel As Boolean
Dim gcdg As Object



Private Sub Form_Load()
    Me.Show
    Me.KeyPreview = True

    rtbChr.Move 0, 0
    rtbChr.Width = Me.Width - 10
    rtbChr.Height = Me.Height - 10

    StdW1 = picHexDisp.ScaleWidth / CharsInRow
    StdH1 = picHexDisp.ScaleHeight / CharsInCol
    StdW2 = picChrDisp.ScaleWidth / CharsInRow
    StdH2 = picChrDisp.ScaleHeight / CharsInCol

    ChrW = picHexDisp.TextWidth("X")

   ' txbEdit.Visible = False
   ' txbEdit.Text = ""
    txbEdit.Width = picHexDisp.TextWidth("XX")
    txbEdit.Height = picHexDisp.TextHeight("X")

    'imgOverWriteOn.Visible = False
    'imgOverwriteOff.Visible = False
    'txbEdit.Visible = False
    txbEdit.Move 0, 0
    imgEdit.ToolTipText = ""
    mDirty = False
    setButtons False
End Sub



Private Sub cmdHelp_Click()
   Dim tmp As String
   tmp = tmp & "(1)  View or Edit Mode and Overwrite setting:" & vbCrLf
   tmp = tmp & "      Default screen is View Mode.  To edit, switch to Edit Mode first.   In Edit Mode, the" & vbCrLf
   tmp = tmp & "      default is Overwrite; but you may right-click the mouse to invoke a popup to toggle the" & vbCrLf
   tmp = tmp & "      setting.  Change of byte is effected after moving the cursor away from its current hex." & vbCrLf & vbCrLf
   tmp = tmp & "(2)  Navigation:" & vbCrLf
   tmp = tmp & "      Click buttons, or use PgDn, PgUp, Up, Dn, Home and End keys.  If in Edit Mode, Left," & vbCrLf
   tmp = tmp & "      Right, Shift+PgDn, Shift+PgUp or a mouse-click, also moves the hex edit position." & vbCrLf & vbCrLf
   tmp = tmp & "(3)  Character display:" & vbCrLf
   tmp = tmp & "      Characters of ASCII value less than 32 are displayed as a rectangle." & vbCrLf & vbCrLf
   tmp = tmp & "(4)  Find corresponding character/hex (In View Mode only):" & vbCrLf
   tmp = tmp & "      Click on the hex/character, the corresponding character/hex will be highlighted.   Its" & vbCrLf
   tmp = tmp & "      ASCII and binary values will also be displayed." & vbCrLf & vbCrLf
   tmp = tmp & "(5)  Search facilities:" & vbCrLf
   tmp = tmp & "      Search the whole file, Hex or Character search, Search from start or Search/Find Next." & vbCrLf & vbCrLf
   tmp = tmp & "(Note: For a slower machine, use BinFileEdit1 instead)."
   MsgBox tmp
   If lblFileSpec.Caption <> "" Then picHexDisp.SetFocus
End Sub



Private Sub setButtons(ByVal OnOff As Boolean)
    cmdSave.Enabled = OnOff
    cmdPrintPage.Enabled = OnOff
    cmdSearchOrFindNext.Enabled = OnOff
    cmdSearchFromStart.Enabled = OnOff
    If lblFileSpec.Caption <> "" Then
        imgEdit.Picture = imgViewMode.Picture
    Else
        imgEdit.Picture = imgNoMode.Picture
    End If
    If mFileSize <= mPageSize Then
         OnOff = False
    End If
      ' These always false when not in Edit Mode
    If imgEdit.Appearance = 0 Then
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
    End If
End Sub



Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdFileOpen_Click()
    On Error GoTo errHandler
     
    If mDirty = True Then
        Dim tmp
        tmp = MsgBox("Byte(s) changed; save to file", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
             Exit Sub
        ElseIf tmp = vbYes Then
             cmdSave_Click
             If gCancel Then
                 Exit Sub
             End If
        End If
    End If
    Dim mHandle
    gcdg.flags = cdlOFNFileMustExist
    gcdg.Filename = ""
    gcdg.CancelError = True
    gcdg.ShowOpen
    If gcdg.Filename = "" Then
        Exit Sub
    End If
      ' Read file.
    mHandle = FreeFile
    Open gcdg.Filename For Binary As #mHandle
    mFileSize = LOF(mHandle)
    If mFileSize = 0 Then
         Close mHandle
         MsgBox "Empty file"
         Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ReDim arrByte(1 To mFileSize)
    Get #mHandle, , arrByte
    Close mHandle
       ' We load the file into a hidden richtextboxes to facilitate Search
       ' if required
    rtbChr.Text = ""
    rtbChr.LoadFile gcdg.Filename
     
    lblFileSize.Caption = CStr(mFileSize) & " bytes"
    lblFileSpec.Caption = Space(2) & gcdg.Filename
     
    txbEdit.Move 0, 0
    
    pageStart = 1
    pageEnd = mPageSize
    ShowPage False
     
    mDirty = False
    
      'Ensure to start from View mode again
    imgEdit.Appearance = 0
    imgEdit.ToolTipText = "View Mode is on.  Toggle View/Edit Mode."
    imgOverWriteOn.Visible = False
    imgOverwriteOff.Visible = False
    txbEdit.Visible = False
    
    setButtons True
    
    picHexDisp.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If Err.Number <> 32755 Then
         Screen.MousePointer = vbDefault
         lblFileSize.Caption = ""
         lblFileSpec.Caption = ""
         rtbChr.Text = ""
         picHexDisp.Picture = LoadPicture()
         picChrDisp.Picture = LoadPicture()
         picOffSet1.Picture = LoadPicture()
         picOffset2.Picture = LoadPicture()
         setButtons False
         ErrMsgProc "cmdFileOpen_Click"
    End If
End Sub
Public Sub OpenAtOffset(Filename As String, Offset As Long)
    On Error GoTo errHandler
'___________________________________
'_____________Open File_____________
    Dim mHandle
    mHandle = FreeFile
    Open Filename For Binary As #mHandle
    mFileSize = LOF(mHandle)
    If mFileSize = 0 Then
         Close mHandle
         MsgBox "Empty file"
         Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    ReDim arrByte(1 To mFileSize)
    Get #mHandle, , arrByte
    Close mHandle
       ' We load the file into a hidden richtextboxes to facilitate Search
       ' if required
    rtbChr.Text = ""
    rtbChr.LoadFile Filename
     
    lblFileSize.Caption = CStr(mFileSize) & " bytes"
    lblFileSpec.Caption = Space(2) & Filename
     
    txbEdit.Move 0, 0
    
    pageStart = 1
    pageEnd = mPageSize
    ShowPage False
     
    mDirty = False
    
      'Ensure to start from View mode again
    imgEdit.Appearance = 0
    imgEdit.ToolTipText = "View Mode is on.  Toggle View/Edit Mode."
    imgOverWriteOn.Visible = False
    imgOverwriteOff.Visible = False
    txbEdit.Visible = False
    
    setButtons True
    
    picHexDisp.SetFocus
    Screen.MousePointer = vbDefault
'_____________________________________
'_____________Goto Offset_____________
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    If Offset > mFileSize Then
        MsgBox "Entry exceeds file size"
        txbGoTo.SetFocus
        Exit Sub
    End If
    Dim k
    Dim i As Long
    i = Offset
    If i > mPageSize Then
        k = (i + 1) / CLng(mPageSize)
        k = NoFraction(k)
        pageStart = k * mPageSize + 1
    Else
        pageStart = 1
    End If
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then
        pageEnd = mFileSize
    End If
       ' Just in case in Edit Mode, check if byte needs to be updated first
    updEditByte
    ShowPage True, i, i, vbYellow, vbBlack
    txbGoTo = Offset
    Exit Sub
errHandler:
    If Err.Number <> 32755 Then
         Screen.MousePointer = vbDefault
         lblFileSize.Caption = ""
         lblFileSpec.Caption = ""
         rtbChr.Text = ""
         picHexDisp.Picture = LoadPicture()
         picChrDisp.Picture = LoadPicture()
         picOffSet1.Picture = LoadPicture()
         picOffset2.Picture = LoadPicture()
         setButtons False
         ErrMsgProc "FileOpen_atOffset"
    End If

End Sub


Private Sub cmdSave_Click()
    On Error GoTo errHandler
    gCancel = True
    gcdg.Filename = LTrim(Trim(lblFileSpec.Caption))
    gcdg.flags = cdlOFNHideReadOnly
    gcdg.CancelError = True
    gcdg.Filter = "(*.*)|*.*|"
    gcdg.FilterIndex = 1
    gcdg.flags = cdlOFNOverwritePrompt
    gcdg.ShowSave
    Open gcdg.Filename For Binary Access Write As #1
    Put #1, , arrByte()
    Close #1
    lblFileSpec.Caption = Space(2) & gcdg.Filename
    picHexDisp.SetFocus
    mDirty = False
    gCancel = False
    Exit Sub
errHandler:
    If Err.Number <> 32755 Then
        ErrMsgProc "CmdSave"
    End If
End Sub



Private Sub ShowPage(ByVal Hilit As Boolean, Optional ByVal inStart As Long = 0, _
      Optional ByVal inEnd As Long = 0, Optional ByVal inPaint1 As Long, _
      Optional ByVal inPaint2 As Long)
    On Error Resume Next
    If mSuspend Or lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    Dim strContent As String
    Dim offSetPos As String
    Dim unDispChar As String
    Dim mAscii As Integer
    Dim mHex As String
    Dim x As Integer, y As Integer
    Dim tmp As String
    Dim i As Long
    Dim j As Long
    Dim k
    Dim origX
    Dim origY

    picHexDisp.Picture = LoadPicture()
    picChrDisp.Picture = LoadPicture()
    picOffSet1.Picture = LoadPicture()
    picOffset2.Picture = LoadPicture()
    
      ' Since we repaint, any values in ASCII & Binary labels are no longer valid
    lblAscii.Caption = ""
    lblBinary.Caption = ""
    
      ' Adjust if required - safety
    If mFileSize <= mPageSize Then
         pageStart = 1
         pageEnd = mFileSize
    Else
         If pageStart < 1 Then
             pageStart = 1
             pageEnd = mPageSize
             If pageEnd > mFileSize Then pageEnd = mFileSize
         End If
         If pageEnd > mFileSize Then
             k = (mFileSize - 1) / mPageSize
             k = NoFraction(k)
             pageStart = k * mPageSize + 1
             pageEnd = pageStart + mPageSize - 1
             If pageEnd > mPageSize Then pageEnd = mPageSize
         End If
    End If
    
      ' Also adjust if required - safety
    If (inStart > 0 And inStart < pageStart) Then inStart = 0
    If (inEnd > 0 And inEnd > pageEnd) Then inEnd = 0
    
      ' Display offset subhead
    picOffset2.CurrentY = 3
    For x = 0 To 15
        tmp = Format$(x, "@@")
        picOffset2.CurrentX = x * StdW1
        picOffset2.Print tmp;
    Next x
    
      ' Restart from top
    picHexDisp.CurrentX = 0
    picChrDisp.CurrentX = 0
    picOffSet1.CurrentY = 0
    picHexDisp.CurrentY = 0
    picChrDisp.CurrentY = 0
    
    unDispChar = Chr$(1)
    i = pageStart
    Do While i <= pageEnd
        offSetPos = Format$(i, " @@@@@@@")
        picOffSet1.Print offSetPos
        
        For j = 0 To 15
            If (i + j) > pageEnd Or (i + j) > mFileSize Then
                Exit For
            Else
                mAscii = arrByte(i + j)
                   ' For Hex area
                mHex = Hex(mAscii)
                If Len(mHex) < 2 Then
                    mHex = "0" & mHex
                End If
                
                picHexDisp.CurrentX = j * StdW1
                
                If Hilit = True And (inStart > 0 And inEnd > 0) Then
                    If (i + j) >= inStart And (i + j) <= inEnd Then
                        origX = picHexDisp.CurrentX
                        origY = picHexDisp.CurrentY
                        x = j * StdW1 - ChrW * 0.4
                        y = picHexDisp.CurrentY
                        picHexDisp.ForeColor = inPaint1
                        picHexDisp.Line (x, y)-(x + ChrW * 2.8, _
                              y + picHexDisp.TextHeight("X")), , BF    '"XX" + 0.4*2
                        picHexDisp.ForeColor = inPaint2
                        picHexDisp.CurrentX = origX
                        picHexDisp.CurrentY = origY
                        picHexDisp.Print mHex;
                        picHexDisp.ForeColor = vbBlack
                    Else
                        picHexDisp.Print mHex;
                    End If
                Else
                    picHexDisp.Print mHex;
                End If
                
                x = j * StdW2
                picChrDisp.CurrentX = x
                If (mAscii >= 31) And (mAscii <= 127) Then
                 picChrDisp.ForeColor = vbRed
             Else
                 picChrDisp.ForeColor = vbBlack
             End If
    ' For Chr area
                If Hilit = True And (inStart > 0 And inEnd > 0) Then
                    If (i + j) >= inStart And (i + j) <= inEnd Then
                        origX = picChrDisp.CurrentX
                        origY = picChrDisp.CurrentY
                        
                        y = picChrDisp.CurrentY
                        picChrDisp.ForeColor = inPaint1
                        picChrDisp.Line (x, y)-(x + ChrW, _
                           y + picChrDisp.TextHeight("X")), inPaint1, BF      ' "X"
                        picChrDisp.ForeColor = inPaint2
                        picChrDisp.CurrentX = origX
                        picChrDisp.CurrentY = origY
                        If mAscii > 31 Then
                            picChrDisp.Print Chr(mAscii);
                        Else
                            picChrDisp.Print unDispChar;
                        End If
                        picChrDisp.ForeColor = vbBlack
                    Else
                        If mAscii > 31 Then
                            picChrDisp.Print Chr(mAscii);
                        Else
                            picChrDisp.Print unDispChar;
                        End If
                    End If
                Else
                    If mAscii > 31 Then
                        picChrDisp.Print Chr(mAscii);
                    Else
                        picChrDisp.Print unDispChar;
                    End If
                End If
            End If
        Next j
        i = i + CharsInRow
        picHexDisp.Print               ' Force picHexDisp change row after earlier ";"
        picHexDisp.CurrentX = 0
        picChrDisp.Print
        picChrDisp.CurrentX = 0        ' Force picChrDisp change row after earlier ";"
    Loop
    For i = 1 To 3
        picHexDisp.Line (StdW1 * 4 * i - ChrW * 0.6, 0)-(StdW1 * 4 * i - ChrW * 0.6, 3), vbBlue, BF
    Next i
    For i = 1 To 3
        picHexDisp.Line (StdW1 * 4 * i - ChrW * 0.6, picHexDisp.ScaleHeight - 4)- _
                (StdW1 * 4 * i - ChrW * 0.6, picHexDisp.ScaleHeight), vbBlue, BF
    Next i
      ' If Edit Mode
    If imgEdit.Appearance = 1 Then
        k = GetByteIndex(txbEdit.Left, txbEdit.Top)
        If k > pageEnd Then
              ' Put txbEdit to pageEnd position
            i = pageEnd - pageStart
            txbEdit.Left = NoFraction(i Mod CharsInRow) * StdW1
            txbEdit.Top = NoFraction(i / CharsInRow) * StdH1
        End If
          ' Should user clicks PgDn....Search or GoTo
        txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
        origHex = txbEdit.Text
        txbEdit.SetFocus
    End If
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mDirty = True Then
        Dim tmp
        tmp = MsgBox("Byte(s) changed; save to file", vbYesNoCancel + vbQuestion)
        If tmp = vbCancel Then
             Cancel = True
        ElseIf tmp = vbYes Then
             cmdSave_Click
             If gCancel Then
                 Cancel = True
             End If
        End If
    End If
End Sub



Private Sub ImgEdit_Click()
    If lblFileSpec.Caption = "" Then Exit Sub
    If imgEdit.Appearance = 1 Then
        imgEdit.Appearance = 0
        imgEdit.Picture = imgViewMode
        imgEdit.ToolTipText = "View Mode is on.  Toggle View/Edit Mode."
        txbEdit.Visible = False
        imgOverWriteOn.Visible = False
        imgOverwriteOff.Visible = False
           ' Disable left and right
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
        ShowPage False                            ' Clear highlight if any
    Else
        imgEdit.Appearance = 1
        imgEdit.Picture = imgEditMode
        imgEdit.ToolTipText = "Edit Mode is on.  Toggle View/Edit Mode."
        txbEdit.Visible = True
        txbEdit.Move 0, 0
        ShowPage True, 1, 1, vbRed, vbYellow
        txbEdit.Text = GetByteHex(2, 2)
        imgOverWriteOn.Visible = (PopUpOverWrite.Checked = True)
        imgOverwriteOff.Visible = (PopUpOverWrite.Checked = False)
           ' Enable left and right
        cmdLeft.Enabled = True
        cmdRight.Enabled = True
        txbEdit.SetFocus
        origHex = txbEdit.Text
    End If
End Sub



Private Sub OptSearch_Click(Index As Integer)
    rtbChr.SelStart = 0                    ' For chr search
    rtbChr.SelLength = 0
    prevFoundPos = 0                       ' For hex search
    If OptSearch(0).Value = True Then
          ckbCaseSensitive.Value = 0
          If Len(txbSearch.Text) > 0 Then
               Dim tmp1 As String, tmp2 As String
               Dim i As Integer
                  ' Purposely not to use Replace()
               tmp1 = txbSearch.Text
               tmp2 = ""
               For i = 1 To Len(tmp1)
                    If Mid(tmp1, i, 1) <> " " Then
                       tmp2 = tmp2 & Mid(tmp1, i, 1)
                    End If
               Next i
               txbSearch.Text = tmp2
               txbSearch.Text = UCase(txbSearch.Text)
          End If
    Else
          ckbCaseSensitive.Value = 1            ' Default
    End If
End Sub



Private Sub txbSearch_KeyPress(KeyAscii As Integer)
    If OptSearch(0).Value = True Then
         KeyAscii = FilterHexKey(KeyAscii)
    End If
End Sub



Private Sub txbGoTo_KeyPress(KeyAscii As Integer)
    KeyAscii = FilterNumericKey(KeyAscii)
End Sub



Private Sub txbEdit_KeyPress(KeyAscii As Integer)
    KeyAscii = FilterHexKey(KeyAscii)
    If PopUpOverWrite.Checked = True Then
         Dim i
         If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
             (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Then
             i = SendMessageLong(txbEdit.hwnd, EM_GETSEL, 0, 0&) \ &H10000
             If i > 1 Then i = 1
             txbEdit.SelStart = i
             txbEdit.SelLength = 1
         End If
    End If
End Sub




Private Sub txbEdit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbRightButton Then
         Exit Sub
    End If
    txbEdit.Enabled = False
    txbEdit.Enabled = True
    PopupMenu mnuPopOverWrite
End Sub



Private Sub picHexDisp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopOverWrite
        Exit Sub
    End If
    
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    Dim k As Long
    Dim i, j
    Dim mHex As String
    k = GetByteIndex(x, y)
    If k > pageEnd Then                          ' Outside displayed area
        Exit Sub
    End If
    If imgEdit.Appearance = 0 Then
        ShowPage True, k, k, vbYellow, vbBlue        ' So to result in yellow, same as ForeColor of lblAscii & lblBinary
        mHex = Hex$(arrByte(k))
        If Len(mHex) < 2 Then mHex = "0" & mHex
        lblAscii.Caption = Trim(CStr(CInt("&h" & mHex)))
        lblBinary.Caption = HexToBinStr(mHex)
    Else
        updEditByte
        i = NoFraction(x / StdW1) * StdW1
        j = NoFraction(y / StdH1) * StdH1
        txbEdit.Move i, j
        k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
        ShowPage True, k, k, vbRed, vbYellow
        txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
        origHex = txbEdit.Text
    End If
End Sub



Private Sub picChrDisp_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
       ' Since we cannot do the same in picHexDisp, so don't allow
       ' it in picChrDisp
    If imgEdit.Appearance = 1 Then
         Exit Sub
    End If
    Dim i, j
    Dim k As Long
    Dim mHex As String
    i = NoFraction(x / StdW2)
    j = NoFraction(y / StdH2) * CharsInRow
    k = pageStart + j + i
    If k > pageEnd Then                          ' Outside displayed area
         Exit Sub
    End If
    ShowPage True, k, k, vbYellow, vbBlue
    mHex = Hex$(arrByte(k))
    If Len(mHex) < 2 Then mHex = "0" & mHex
    lblAscii.Caption = Trim(CStr(CInt("&h" & mHex)))
    lblBinary.Caption = HexToBinStr(mHex)
End Sub




Private Sub Form_KeyUp(keycode As Integer, Shift As Integer)
    Dim k, i
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    Select Case keycode
        Case 34           ' PgDn
             If imgEdit.Appearance = 1 Then
                  If Shift = 0 Then
                      cmdPgDn_Click
                  Else
                      k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
                      If k = pageEnd Then Exit Sub
                      updEditByte
                      i = pageEnd - pageStart
                      txbEdit.Left = NoFraction(i Mod CharsInRow) * StdW1
                      txbEdit.Top = NoFraction(i / CharsInRow) * StdH1
                      ShowPage True, pageEnd, pageEnd, vbRed, vbYellow
                      txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
                      origHex = txbEdit.Text
                  End If
             Else
                  cmdPgDn_Click
             End If
        Case 33           ' PgUp
             If imgEdit.Appearance = 1 Then
                  If Shift = 0 Then
                      cmdPgUp_Click
                  Else
                      k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
                      If k = pageStart Then Exit Sub
                      updEditByte
                      txbEdit.Move 0, 0
                      txbEdit.Text = GetByteHex(2, 2)
                      ShowPage True, pageStart, pageStart, vbRed, vbYellow
                      origHex = txbEdit.Text
                  End If
             Else
                 cmdPgUp_Click
             End If
        Case 40           ' Dn
            cmdDn_Click
        Case 38           ' Up
            cmdUp_Click
        Case 36           ' Home
            cmdFirst_Click
        Case 35           ' End
            cmdLast_Click
        Case 37           ' Left
            If imgEdit.Appearance = 1 Then
                cmdLeft_Click
            End If
        Case 39           ' Right
            If imgEdit.Appearance = 1 Then
                cmdRight_Click
            End If
        Case vbKeyDelete
            If imgEdit.Appearance = 1 Then
                txbEdit.SetFocus
            End If
    End Select
End Sub



Private Function GetByteIndex(ByVal x As Single, ByVal y As Single) As Long
    Dim i, j
    Dim k As Long
    i = NoFraction(x / StdW1)
    j = NoFraction(y / StdH1) * CharsInRow
    k = pageStart + j + i
    GetByteIndex = k
End Function




Private Function GetByteHex(ByVal x As Single, ByVal y As Single) As String
    Dim i, j
    Dim k As Long
    Dim mHex As String
    i = NoFraction(x / StdW1)
    j = NoFraction(y / StdH1) * CharsInRow
    k = pageStart + j + i
    If k > pageEnd Then                          ' Outside displayed area
         mHex = ""
    Else
         mHex = Hex$(arrByte(k))
         If Len(mHex) = 1 Then mHex = "0" & mHex
    End If
    GetByteHex = mHex
End Function



Private Sub UpdateByte(ByVal x As Single, ByVal y As Single)
    Dim k As Long
    Dim mHex As String
    k = GetByteIndex(x, y)
    mHex = arrByte(k)
    If Len(mHex) = 1 Then mHex = "0" & mHex
    If mHex <> txbEdit.Text Then
        arrByte(k) = CByte("&h" & txbEdit.Text)
        mDirty = True
    End If
End Sub



Private Function updEditByte() As Boolean
    updEditByte = False
    If imgEdit.Appearance = 0 Then
        Exit Function
    End If
    If Len(LTrim(Trim(txbEdit.Text))) = 2 Then
        If txbEdit.Text <> origHex Then
             UpdateByte txbEdit.Left + 2, txbEdit.Top + 2
             updEditByte = True
        End If
    End If
End Function



Private Sub cmdPgDn_Click()
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    Dim k, i
    If mFileSize <= mPageSize Then
         Exit Sub
    End If
    If pageEnd = mFileSize Then
         Exit Sub
    End If
    updEditByte
    picHexDisp.SetFocus
    pageStart = pageStart + mPageSize
    If pageStart > mFileSize Then pageStart = pageStart - mPageSize
    pageEnd = pageEnd + mPageSize
    If pageEnd > mFileSize Then pageEnd = mFileSize
    
    If imgEdit.Appearance = 1 Then
        k = GetByteIndex(txbEdit.Left, txbEdit.Top)
          ' Safety
        If k > pageEnd Then
              ' Put txbEdit to pageEnd position
            i = pageEnd - pageStart
            txbEdit.Left = NoFraction(i Mod CharsInRow) * StdW1
            txbEdit.Top = NoFraction(i / CharsInRow) * StdH1
            k = GetByteIndex(txbEdit.Left, txbEdit.Top)
        End If
        ShowPage True, k, k, vbRed, vbYellow
        txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
        origHex = txbEdit.Text
    Else
        ShowPage False
    End If
End Sub



Private Sub cmdPgUp_Click()
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    Dim k
    If mFileSize <= mPageSize Then
         Exit Sub
    End If
    If pageStart = 1 Then
         Exit Sub
    End If
    updEditByte
    picHexDisp.SetFocus
    pageStart = pageStart - mPageSize
    If pageStart < 1 Then pageStart = 1
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then pageEnd = mFileSize
    If imgEdit.Appearance = 1 Then
        k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
        ShowPage True, k, k, vbRed, vbYellow
        txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
        origHex = txbEdit.Text
    Else
        ShowPage False
    End If
End Sub



Private Sub cmdDn_Click()
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    Dim k
    If imgEdit.Appearance = 0 Then
        If mFileSize <= mPageSize Then
            Exit Sub
        End If
        picHexDisp.SetFocus
        pageStart = pageStart + CharsInRow
        If pageStart > mFileSize Then pageStart = pageStart - CharsInRow
        pageEnd = pageStart + mPageSize - 1
        If pageEnd > mFileSize Then pageEnd = mFileSize
        ShowPage False
    Else
        k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
        If k = mFileSize Then Exit Sub
        updEditByte
        If k + CharsInRow > mFileSize Then        ' One line below is empty
               ' If last line not visible, show it, but don't call updEditByte
              pageStart = pageStart + CharsInRow
              If pageStart > mFileSize Then
                    pageStart = pageStart - CharsInRow      ' Restore
                    Exit Sub
              End If
              pageEnd = pageStart + mPageSize - 1
              If pageEnd > mFileSize Then pageEnd = mFileSize
         Else
              If txbEdit.Top + StdH1 > StdH1 * (CharsInCol - 1) Then
                   pageStart = pageStart + CharsInRow
                   If pageStart > mFileSize Then
                        pageStart = pageStart - CharsInRow
                        Exit Sub
                   End If
                   pageEnd = pageStart + mPageSize - 1
                   If pageEnd > mFileSize Then pageEnd = mFileSize
              Else
                   txbEdit.Move txbEdit.Left, txbEdit.Top + StdH1
              End If
         End If
         k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
         ShowPage True, k, k, vbRed, vbYellow
         txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
         origHex = txbEdit.Text
    End If
End Sub



Private Sub cmdUp_Click()
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    Dim k
    If imgEdit.Appearance = 0 Then
         If mFileSize <= mPageSize Then
             Exit Sub
         End If
         If pageStart <= CharsInRow Then
             Exit Sub
         End If
         picHexDisp.SetFocus
         pageStart = pageStart - CharsInRow
         If pageStart < 1 Then pageStart = pageStart + CharsInRow
         pageEnd = pageStart + mPageSize - 1
         If pageEnd > mFileSize Then pageEnd = mFileSize
         ShowPage False
    Else
         k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
         If k = 1 Then Exit Sub
         updEditByte
         If txbEdit.Top - StdH1 < 0 Then
              pageStart = pageStart - CharsInRow
              If pageStart < 1 Then
                   pageStart = pageStart + CharsInRow
                   Exit Sub
              End If
              pageEnd = pageStart + mPageSize - 1
              If pageEnd > mFileSize Then pageEnd = mFileSize
          Else
              txbEdit.Move txbEdit.Left, txbEdit.Top - StdH1
          End If
          k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
          ShowPage True, k, k, vbRed, vbYellow
          txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
          origHex = txbEdit.Text
    End If
End Sub



Private Sub cmdLeft_Click()
    Dim k
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    If imgEdit.Appearance <> 1 Then           ' Not edit mode, for safety
        Exit Sub
    End If
    
    k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
    If k = 1 Then Exit Sub
    updEditByte
    If (txbEdit.Left - StdH1 < 0) Then
        If k = pageStart Then
            pageStart = pageStart - CharsInRow
            If pageStart > mFileSize Then
                 pageStart = pageStart + CharsInRow
                 Exit Sub
            End If
            pageEnd = pageStart + mPageSize - 1
            If pageEnd > mFileSize Then
                 pageEnd = mFileSize
            End If
               ' txbEdit at last col of prev row
            txbEdit.Left = StdW1 * (CharsInRow - 1)
            updEditByte
        Else
               ' txbEdit at last col of prev row
            txbEdit.Left = StdW1 * (CharsInRow - 1)
            txbEdit.Top = txbEdit.Top - StdH1
        End If
    Else
        txbEdit.Move txbEdit.Left - StdW1, txbEdit.Top
    End If
    k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
    ShowPage True, k, k, vbRed, vbYellow
    txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
    origHex = txbEdit.Text
End Sub



Private Sub cmdRight_Click()
    Dim k
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    If imgEdit.Appearance <> 1 Then           ' Not edit mode, for safety
        Exit Sub
    End If
    
    k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
    If k = mFileSize Then Exit Sub
    updEditByte
    If (txbEdit.Left + StdW1) > (picHexDisp.Width - StdW1) Then
        If k = pageEnd Then
             pageStart = pageStart + CharsInRow
             If pageStart < 1 Then
                  pageStart = pageStart - CharsInRow
                  Exit Sub
             End If
             pageEnd = pageStart + mPageSize - 1
             If pageEnd > mFileSize Then
                  pageEnd = mFileSize
             End If
                ' txbEdit at first col of next row
             txbEdit.Left = 0
             updEditByte
        Else
               ' txbEdit at first col of next row
             txbEdit.Left = 0
             txbEdit.Top = txbEdit.Top + StdH1
        End If
    Else
        txbEdit.Move txbEdit.Left + StdW1, txbEdit.Top
    End If
    k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
    ShowPage True, k, k, vbRed, vbYellow
    txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
    origHex = txbEdit.Text
End Sub




Private Sub cmdFirst_Click()
    If lblFileSpec.Caption = "" Then
        Exit Sub
    End If
    Dim k
    picHexDisp.SetFocus
    If imgEdit.Appearance = 1 Then
          updEditByte
          pageStart = 1
          pageEnd = mPageSize
          If pageEnd > mFileSize Then pageEnd = mFileSize
          txbEdit.Move 0, 0
          ShowPage True, 1, 1, vbRed, vbYellow
          txbEdit.Text = GetByteHex(2, 2)
          origHex = txbEdit.Text
    Else
          If mFileSize <= mPageSize Then
              Exit Sub
          End If
          pageStart = 1
          pageEnd = mPageSize
          If pageEnd > mFileSize Then pageEnd = mFileSize
          ShowPage False
    End If
End Sub



Private Sub cmdLast_Click()
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    Dim k, i
    If imgEdit.Appearance = 1 Then
         updEditByte
         If pageEnd = mFileSize Then
             k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
             If k < pageEnd Then
                  i = pageEnd - pageStart
                  txbEdit.Left = NoFraction(i Mod CharsInRow) * StdW1
                  txbEdit.Top = NoFraction(i / CharsInRow) * StdH1
                  k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
                  ShowPage True, k, k, vbRed, vbYellow
                  txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
                  origHex = txbEdit.Text
                  Exit Sub
             End If
         End If
         k = (mFileSize - 1) / mPageSize
         k = NoFraction(k)
         pageStart = k * mPageSize + 1
         pageEnd = pageStart + mPageSize - 1
         If pageEnd > mFileSize Then pageEnd = mFileSize
         ShowPage False
         k = GetByteIndex(txbEdit.Left + 2, txbEdit.Top + 2)
         ShowPage True, k, k, vbRed, vbYellow
         txbEdit.Text = GetByteHex(txbEdit.Left + 2, txbEdit.Top + 2)
         origHex = txbEdit.Text
    Else
         If mFileSize <= mPageSize Then
             Exit Sub
         End If
         k = (mFileSize - 1) / mPageSize
         k = NoFraction(k)
         pageStart = k * mPageSize + 1
         pageEnd = pageStart + mPageSize - 1
         If pageEnd > mFileSize Then pageEnd = mFileSize
         ShowPage False
    End If
End Sub




Private Sub cmdSearchFromStart_Click()
    rtbChr.SelStart = 0                    ' For chr search
    rtbChr.SelLength = 0
    prevFoundPos = 0                       ' For hex search
    cmdSearchOrFindNext_Click
End Sub



Private Sub cmdSearchOrFindNext_Click()
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    If Len(txbSearch.Text) = 0 Then
         MsgBox "No search text entered yet"
         txbSearch.SetFocus
         Exit Sub
    End If
    picHexDisp.SetFocus
    If OptSearch(1).Value = True Then
         doChrSearch
    Else
         If (Len(txbSearch.Text) Mod 2) > 0 Then
              MsgBox "Incorrect hex value entered"
              Exit Sub
         End If
         doHexSearch
    End If
End Sub




Private Sub doChrSearch()
    On Error Resume Next
    Dim foundStartPos As Long
    Screen.MousePointer = vbHourglass
    foundStartPos = rtbChr.SelStart
    If foundStartPos = 0 Then
        If ckbCaseSensitive.Value = 0 Then              ' Not checked
             foundStartPos = InStr(foundStartPos + 1, UCase(rtbChr.Text), UCase(txbSearch.Text))
        Else                                            ' Checked
             foundStartPos = InStr(foundStartPos + 1, rtbChr.Text, txbSearch.Text)
        End If
    Else
        If ckbCaseSensitive.Value = 0 Then              ' Not checked
             foundStartPos = InStr(foundStartPos + 2, UCase(rtbChr.Text), UCase(txbSearch.Text))
        Else                                            ' Checked
             foundStartPos = InStr(foundStartPos + 2, rtbChr.Text, txbSearch.Text)
        End If
    End If
    If foundStartPos = prevFoundPos Then
        prevFoundPos = 0
    Else
        prevFoundPos = foundStartPos
    End If
    If foundStartPos = 0 Then
          ' Start the search from beginning
        rtbChr.SelStart = 0
        rtbChr.SelLength = 0
    Else
          ' Start the search from this point
        rtbChr.SelStart = foundStartPos - 1
        rtbChr.SelLength = Len(txbSearch.Text)
    End If
    Screen.MousePointer = vbDefault
    If prevFoundPos > 0 And rtbChr.SelLength > 0 Then
        Dim k
        k = (foundStartPos + 1) / CLng(mPageSize)
        k = NoFraction(k)
        pageStart = k * mPageSize + 1
        pageEnd = pageStart + mPageSize - 1
        If pageEnd > mFileSize Then pageEnd = mFileSize
        k = foundStartPos + (Len(txbSearch.Text) - 1)
        If k > pageEnd Then k = pageEnd
          ' So to result in Aqua, similar to txbSearch BgColor
          ' Just in case in Edit Mode, check if byte value needs to be updated first
        updEditByte
        ShowPage True, foundStartPos, k, &HFFFF00, vbRed
        Exit Sub
    End If
    MsgBox txbSearch.Text & vbCrLf & vbCrLf & "Searched to end."
End Sub



Private Sub doHexSearch()
    On Error Resume Next
    Dim HexCtn As Integer
    Dim i, j
    Dim mMatch As Boolean
    Dim foundStartPos As Long
    Screen.MousePointer = vbHourglass
    HexCtn = Len(txbSearch.Text) / 2
    ReDim arrHexByte(1 To HexCtn)
    For i = 1 To HexCtn
         arrHexByte(i) = CByte("&h" & (Mid(txbSearch.Text, (i * 2 - 1), 2)))
    Next i
    foundStartPos = prevFoundPos + 1
    For i = foundStartPos To (UBound(arrByte) - (HexCtn - 1))
         If arrByte(i) = arrHexByte(1) Then
              mMatch = True
                ' Compare rest bytes
              For j = 1 To (HexCtn - 1)
                   If arrByte(i + j) <> arrHexByte(1 + j) Then
                       mMatch = False
                       Exit For
                   End If
              Next j
              If mMatch = True Then
                   Dim k
                   foundStartPos = i
                   prevFoundPos = i
                   k = (foundStartPos + 1) / CLng(mPageSize)
                   k = NoFraction(k)
                   pageStart = k * mPageSize + 1
                   pageEnd = pageStart + mPageSize - 1
                   If pageEnd > mFileSize Then pageEnd = mFileSize
                   k = foundStartPos + (HexCtn - 1)
                   If k > pageEnd Then k = pageEnd
                   updEditByte
                   ShowPage True, foundStartPos, k, &HFFFF00, vbRed
                   Screen.MousePointer = vbDefault
                   Exit Sub
              End If
         End If
    Next i
    Screen.MousePointer = vbDefault
    prevFoundPos = 0
    MsgBox txbSearch.Text & vbCrLf & vbCrLf & "Searched to end."
End Sub



Private Sub cmdGoTo_Click()
    If lblFileSpec.Caption = "" Then
         Exit Sub
    End If
    If Len(txbGoTo.Text) = 0 Then
        MsgBox "No byte position entered yet"
        txbGoTo.SetFocus
        Exit Sub
    ElseIf Val(txbGoTo.Text) > mFileSize Then
        MsgBox "Entry exceeds file size"
        txbGoTo.SetFocus
        Exit Sub
    End If
    Dim k
    Dim i As Long
    i = Val(txbGoTo.Text)
    If i > mPageSize Then
        k = (i + 1) / CLng(mPageSize)
        k = NoFraction(k)
        pageStart = k * mPageSize + 1
    Else
        pageStart = 1
    End If
    pageEnd = pageStart + mPageSize - 1
    If pageEnd > mFileSize Then
        pageEnd = mFileSize
    End If
       ' Just in case in Edit Mode, check if byte needs to be updated first
    updEditByte
    ShowPage True, i, i, vbYellow, vbBlue
End Sub



Function NoFraction(ByVal inVal As Variant) As Long
    Dim x As Integer
    Dim tmp As String
    Dim k As Long
    tmp = CStr(inVal)
    x = InStr(tmp, ".")
    If x > 0 Then
        tmp = Left(tmp, x - 1)
    End If
    k = Val(tmp)
    NoFraction = k
End Function




Private Sub cmdPrintPage_Click()
    On Error GoTo errHandler
    gcdg.CancelError = True
    gcdg.flags = cdlPDReturnDC + cdlPDNoPageNums + cdlPDNoSelection
    gcdg.ShowPrinter
    picContainer.Picture = LoadPicture()
    BitBlt picContainer.hdc, picOffSet1.Left, picOffSet1.Top, picOffSet1.ScaleWidth, _
       picOffSet1.ScaleHeight, picOffSet1.hdc, 0, 0, vbSrcCopy
    BitBlt picContainer.hdc, picOffset2.Left, picOffset2.Top, picOffset2.ScaleWidth, _
       picOffset2.ScaleHeight, picOffset2.hdc, 0, 0, vbSrcCopy
    BitBlt picContainer.hdc, picHexDisp.Left, picHexDisp.Top, picHexDisp.ScaleWidth, _
       picHexDisp.ScaleHeight, picHexDisp.hdc, 0, 0, vbSrcCopy
    BitBlt picContainer.hdc, picChrDisp.Left, picChrDisp.Top, picChrDisp.ScaleWidth, _
       picChrDisp.ScaleHeight, picChrDisp.hdc, 0, 0, vbSrcCopy
    picContainer.Picture = picContainer.Image
    Printer.Print ""
    Printer.CurrentX = 1440
    Printer.CurrentY = 1440
    Printer.Print gcdg.Filename
    Printer.PaintPicture picContainer.Picture, 1440, 2880
    Printer.EndDoc
    picContainer.Picture = LoadPicture()
    picHexDisp.SetFocus
    Exit Sub
errHandler:
If Err <> 32755 Then
    ErrMsgProc "cmdPrintPage_Click"
End If
End Sub



Private Function HexToBinStr(ByVal inHex As String) As String
    Dim mDec As Integer
    Dim s As String
    Dim i
    mDec = CInt("&h" & inHex)
    s = Trim(CStr(mDec Mod 2))
    i = mDec \ 2
    Do While i <> 0
        s = Trim(CStr(i Mod 2)) & s
        i = i \ 2
    Loop
    Do While Len(s) < 8
        s = "0" & s
    Loop
    HexToBinStr = s
    Exit Function
End Function



Private Sub popUpOverWrite_click()
    PopUpOverWrite.Checked = Not PopUpOverWrite.Checked
    If txbEdit.Visible = True And txbEdit.Enabled = True Then
        imgOverWriteOn.Visible = (PopUpOverWrite.Checked = True)
        imgOverwriteOff.Visible = (PopUpOverWrite.Checked = False)
        txbEdit.SetFocus
    Else
        imgOverWriteOn.Visible = False
        imgOverwriteOff.Visible = False
    End If
End Sub




Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description
End Sub



Function FilterHexKey(mInkey) As Integer
    If mInkey < Asc("0") Or mInkey > Asc("9") Then
        If Not (mInkey >= Asc("A") And mInkey <= Asc("F")) Then
            If Not (mInkey >= Asc("a") And mInkey <= Asc("f")) Then
                 If mInkey <> 8 Then
                      mInkey = 0
                 End If
            End If
        End If
    End If
    If mInkey >= Asc("a") And mInkey <= Asc("f") Then
        mInkey = mInkey - 32
    End If
    FilterHexKey = mInkey
End Function



Function FilterNumericKey(inkey) As Integer
    If inkey < Asc("0") Or inkey > Asc("9") Then
        If inkey <> 8 Then
              inkey = 0
        End If
    End If
    FilterNumericKey = inkey
End Function


