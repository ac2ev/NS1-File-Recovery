VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Ns1KFRaC"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12660
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":2CCA
   ScaleHeight     =   7590
   ScaleWidth      =   12660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbListViewEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6945
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3660
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ListView lstView 
      Height          =   6315
      Left            =   3150
      TabIndex        =   16
      Top             =   960
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11139
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ColHdrIcons     =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView treFiles 
      Height          =   1620
      Left            =   36
      TabIndex        =   14
      Top             =   960
      Width           =   2856
      _ExtentX        =   5027
      _ExtentY        =   2858
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   870
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   12600
      Begin VB.CommandButton cmdAbort 
         BackColor       =   &H000000FF&
         Caption         =   "Abort"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   9435
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   60
         Visible         =   0   'False
         Width           =   2595
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2184
         Top             =   444
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   20
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":300C
               Key             =   "ap"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5CE6
               Key             =   "fldrcls"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6080
               Key             =   "fldropn"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":655E
               Key             =   "file"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6878
               Key             =   "NotDone"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6ACF
               Key             =   "Done"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6D80
               Key             =   "fault"
            EndProperty
         EndProperty
      End
      Begin VB.Label lblHeaderValue 
         Height          =   240
         Index           =   3
         Left            =   5205
         TabIndex        =   13
         Top             =   30
         Width           =   510
      End
      Begin VB.Label lblheader 
         Alignment       =   1  'Right Justify
         Caption         =   "Bad Records Removed"
         Height          =   240
         Index           =   3
         Left            =   3288
         TabIndex        =   12
         Top             =   24
         Width           =   1800
      End
      Begin VB.Label lblFilename 
         Alignment       =   1  'Right Justify
         Height          =   225
         Left            =   3165
         TabIndex        =   11
         Top             =   660
         Width           =   9255
      End
      Begin VB.Label lblheader 
         Alignment       =   1  'Right Justify
         Caption         =   "AP Count"
         Height          =   240
         Index           =   2
         Left            =   495
         TabIndex        =   10
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label lblheader 
         Alignment       =   1  'Right Justify
         Caption         =   "File Format Version"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   322
         Width           =   1395
      End
      Begin VB.Label lblHeaderValue 
         Height          =   240
         Index           =   1
         Left            =   1590
         TabIndex        =   8
         Top             =   322
         Width           =   510
      End
      Begin VB.Label lblheader 
         Alignment       =   1  'Right Justify
         Caption         =   "Header"
         Height          =   240
         Index           =   0
         Left            =   495
         TabIndex        =   7
         Top             =   60
         Width           =   1050
      End
      Begin VB.Label lblHeaderValue 
         Height          =   240
         Index           =   0
         Left            =   1590
         TabIndex        =   6
         Top             =   45
         Width           =   510
      End
      Begin VB.Label lblHeaderValue 
         Height          =   240
         Index           =   2
         Left            =   1590
         TabIndex        =   5
         Top             =   585
         Width           =   510
      End
   End
   Begin VB.PictureBox Splitter1 
      BorderStyle     =   0  'None
      Height          =   6345
      Left            =   2955
      ScaleHeight     =   6345
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   855
      Width           =   75
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   11040
      Top             =   285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EDA
            Key             =   "ap"
            Object.Tag             =   "ap"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BB4
            Key             =   "connect"
            Object.Tag             =   "connect"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D0E
            Key             =   "key"
            Object.Tag             =   "key"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E68
            Key             =   "lock"
            Object.Tag             =   "lock"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9FC2
            Key             =   "lap"
            Object.Tag             =   "lap"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A11C
            Key             =   "nowep"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A276
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A6C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B3A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C07C
            Key             =   "DownArrow"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C14E
            Key             =   "UpArrow"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C220
            Key             =   "fault"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treView 
      Height          =   4590
      Left            =   30
      TabIndex        =   2
      Top             =   2625
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   8096
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "imgList"
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin VB.PictureBox ProgressBar1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   30
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   1
      Top             =   -120
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5910
      Top             =   5205
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7335
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11536
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuMerge 
         Caption         =   "Merge"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu zzbatch 
         Caption         =   "Batch"
         Begin VB.Menu mnuLoadList 
            Caption         =   "Load List"
         End
         Begin VB.Menu mnuBatch 
            Caption         =   "Run Batch"
         End
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save as ns1"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuOpenNet 
         Caption         =   "Open Recoverd file in NetStumbler"
         Enabled         =   0   'False
      End
      Begin VB.Menu zzsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenSavens1 
         Caption         =   "Open/Save ns1"
      End
      Begin VB.Menu mnuopnwiscan 
         Caption         =   "Open/Export Wi-Scan Summary"
      End
      Begin VB.Menu zzsep1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu zzTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu zzExport 
         Caption         =   "Export"
         Enabled         =   0   'False
         Begin VB.Menu mnuEachToOwn 
            Caption         =   "Each to it's own ns1"
         End
         Begin VB.Menu mnuExport 
            Caption         =   "To Text (.csv)"
         End
         Begin VB.Menu mnuWiscan 
            Caption         =   "Wi-Scan Summary"
         End
         Begin VB.Menu mnuKML 
            Caption         =   "Google .KML"
         End
         Begin VB.Menu mnuExcel 
            Caption         =   "To Excel"
         End
      End
   End
   Begin VB.Menu zzExplore 
      Caption         =   "Record Recovery"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuExportsingle 
         Caption         =   "Export To Ns1"
      End
      Begin VB.Menu mnuRemoveRecord 
         Caption         =   "Remove ApData for entry"
      End
   End
   Begin VB.Menu mnuFixes 
      Caption         =   "Fix"
      Enabled         =   0   'False
      Begin VB.Menu mnuclearAPData 
         Caption         =   "ClearAll ApData/GPSData"
      End
   End
   Begin VB.Menu mnuLiveUpdate 
      Caption         =   "Live Update"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInfo 
         Caption         =   "Info"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu lvVSSMenu 
      Caption         =   "VSSMENU"
      Visible         =   0   'False
      Begin VB.Menu lvVSSMenuEdit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu lvVSSMenuPaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0


Dim TotalRecords  As Long

Dim bAuto As Boolean
Public bAbort As Boolean
Public bLoading As Boolean

Dim prevKey As String

Public BatchIndex As Long
Private m_ClsAddProgToSbar As cAddProgToSBar
'//picturebox splitter Coded By Travis John Close - Ides of March, 2000
 

'//variable to hold the width of the Splitter bar
'//medium rare please - size does matter - wider is better - etc., etc.
Private Const SPLT_WDTH As Integer = 75
 
'//an arbitrary value set that will most likely be out of all range
'//used to see if the Splitter has changed position when dragged
Private Const SPLT_DEFAULT_POS As Long = -2000000
 
'//variable to hold the last-sized postion
Private currSplitPosX As Long
 
 
'//variables that hold the offsets from the forms edge
'//I use these in case we have to add stuff above, below, and
'//to the side of the tree & list views at a later time
'//ex: the coolbar is added to the top (see form load)
Private intLvwOffsetRight As Integer
Private intTvwOffsetLeft As Integer
Private intOffsetToptree As Integer
Private intOffsetToplist As Integer
Private intOffsetBottom As Integer

Dim LatestMajorver As String
Dim LatestMinorver As String

 
'//variable to hold the Splitter bar color
Private SplitColor As Long
Dim tHt As LVHITTESTINFO


Private Sub cmdAbort_Click()
    bAbort = True
End Sub

Private Sub SortListView(ByRef List As ListView, ColHeadIndex As Integer)
    
    Dim lcv As Long     'Loop Control Variable
  
    With List
        ' Make sure the Sorted property is set to true
        .Sorted = True
        
        ' Sort according to the colum that was clicked (off by one)
        .SortKey = ColHeadIndex - 1
       
        ' Does the column already have an icon?
        If .ColumnHeaders(ColHeadIndex).Icon = 0 Then
            'No, So we will assume this column is not sorted
            
            ' Set to Ascending order
            .SortOrder = lvwAscending
            
            ' Set the ColumnHeader to be the Up Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "UpArrow"
            
        ' Does the column have an UpArrow icon?
        ElseIf .ColumnHeaders(ColHeadIndex).Icon = "UpArrow" Then
            ' Yes, So the column is in Ascending order, switch to descending
            
            ' Set the Column Icon to the Down Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "DownArrow"
            
            ' Set the sort order to descending
            .SortOrder = lvwDescending
        
        Else
            ' Otherwise sort into ascending order
        
            ' Set to Ascending order
            .SortOrder = lvwAscending
            
            ' Set the ColumnHeader to be the Up Arrow
            .ColumnHeaders(ColHeadIndex).Icon = "UpArrow"
        End If
       
        ' Remove any icon (presumably an arrow icon) from all other columns
        ' For every Column in the ListView Control...
        For lcv = 1 To List.ColumnHeaders.count
            ' Is the current column the clicked column?
            If Not (lcv = ColHeadIndex) Then
                ' No, remove any icon it may have
                .ColumnHeaders(lcv).Icon = 0
            End If
        Next lcv
    
        ' Refresh the display of the ListView Control
        .Refresh
    End With
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ret As VbMsgBoxResult
    If KeyCode = 27 And bLoading Then
     ret = MsgBox("Are you sure you want to Abort Loading?", vbYesNo, "Abort Load")
        If ret = vbYes Then
            bAbort = True
        End If
    End If
End Sub

Private Sub Form_Load()
    
    Dim bLiveUpdate As Boolean
    ' note: declare gDebugMode in any module --> Global gDebugMode as
    Dim BlankArray(16) As Byte
    Set creg = New cRegistry
    creg.hKey = HKEY_LOCAL_MACHINE
    creg.KeyPath = "Software\NS1"
    NoSSID = creg.GetRegistryValue("NoSSID", "", , , , False)
    NoPrint = creg.GetRegistryValue("NoPrint", 0, , , , False)
    Descriptor = creg.GetRegistryValue("KMLDescriptor", 0, , , , False)
    ExportItem = creg.GetRegistryValue("ExportItems", BlankArray(), , , , False)
    Use3D = creg.GetRegistryValue("3D", 1, , , , False)
    
    creg.SetRegistryValue "Version", App.Major & "." & App.Minor & "." & App.Revision, REG_SZ, , , , False
    bLiveUpdate = creg.GetRegistryValue("LiveUpdate", False, , , , False)
    GroupBy = creg.GetRegistryValue("GroupBy", 0, , , , False)
    If bLiveUpdate Then
        mnuLiveUpdate_Click
    End If
    gDebugMode = InIDE(hwnd)
   If gDebugMode Then
      Debug.Print "IDE detected"
   End If
    Set m_ClsAddProgToSbar = New cAddProgToSBar
    Set_Up_Panels
    Set treView.ImageList = imgList
    

    
    'Initialize listview
    tHt.lItem = -1
    
    ' set lvVSS to set nodes for project.
    Call ListView_FullRowSelect(lstView)
    Call ListView_GridLines(lstView)
    
    With lstView
        .ColumnHeaders.Add , "SSID", "SSID"
        .ColumnHeaders.Add , "BSSID", "BSSID"
        .ColumnHeaders.Add , "MaxSignal", "MaxSignal"
        .ColumnHeaders.Add , "MinNoise", "MinNoise"
        .ColumnHeaders.Add , "MaxSNR", "MaxSNR"
        .ColumnHeaders.Add , "flags", "flags"
        .ColumnHeaders.Add , "BeaconInterval", "BeaconInterval"
        .ColumnHeaders.Add , "FirstSeen", "FirstSeen"
        .ColumnHeaders.Add , "LastSeen", "LastSeen"
        .ColumnHeaders.Add , "BestLat", "BestLat"
        .ColumnHeaders.Add , "BestLong", "BestLong"
        .ColumnHeaders.Add , "Name", "Name"
        .ColumnHeaders.Add , "Channels", "Channels"
        .ColumnHeaders.Add , "LastChannel", "LastChannel"
        .ColumnHeaders.Add , "IPAddress", "IPAddress"
        .ColumnHeaders.Add , "MinSignal", "MinSignal"
        .ColumnHeaders.Add , "MaxNoise", "MaxNoise"
        .ColumnHeaders.Add , "DataRate", "DataRate"
        .ColumnHeaders.Add , "IPSubnet", "IPSubnet"
        .ColumnHeaders.Add , "IPMask", "IPMask"
        .ColumnHeaders.Add , "ApFlags", "ApFlags"
        .ColumnHeaders.Add , "InformationElements", "InformationElements"
        .ColumnHeaders.Add , "Begin Offset", "Begin Offset"
        .ColumnHeaders.Add , "End Offset", "End Offset"
End With

'===============================

 '//set the startup variables
    '//control offset values from edge of form
    Let intOffsetToptree = 45 + (Me.treFiles.Top + Me.treFiles.Height)
    Let intOffsetToplist = 45 + (Me.Frame1.Top + Me.Frame1.Height)
    Let intLvwOffsetRight = 45
    Let intTvwOffsetLeft = 45
    Let intOffsetBottom = 45 + Me.StatusBar1.Height

    '//dark gray color used when dragging Splitter
    Let SplitColor = &H808080

    '//set the current Splitter bar position to an arbitrary value that will always be outside
    '//the possibe range. This allows us to check for movement of the spltter bar in subsequent
    '//mousexxx subs.

    Let currSplitPosX = SPLT_DEFAULT_POS
End Sub

Private Sub Form_Resize()
   Dim x1 As Integer 'new left of treeview
    Dim x2 As Integer 'new left of listview
    Dim heighttree As Integer 'new height of controls
    Dim heightlist As Integer 'new height of controls
    Dim width1 As Integer 'new width of treeview
    Dim width2 As Integer 'new width of listview
    
    On Error Resume Next

    '//move the frame into position
    Me.Frame1.Move 0, 0, Me.ScaleWidth, Me.treFiles.Height

    'since we just moved the coolbar, re-calculate the top offset
    Let intOffsetToptree = 60 + (Me.treFiles.Top + Me.treFiles.Height)
    Let intOffsetToplist = 60 + Me.treFiles.Top '(Me.Frame1.Top + Me.Frame1.Height)
    
    '//calculate some positions
    Let heighttree = ScaleHeight - intOffsetBottom - intOffsetToptree
    Let heightlist = ScaleHeight - intOffsetBottom - intOffsetToplist
    
    Let x1 = intTvwOffsetLeft
    Let width1 = treView.Width
    
    Let x2 = x1 + treView.Width + SPLT_WDTH - 1
    Let width2 = ScaleWidth - x2 - intLvwOffsetRight

    '//move the treeview into position
    treView.Move x1% - 1, intOffsetToptree, width1, heighttree
    
    '//move the listview into position
    lstView.Move x2, intOffsetToplist - 15, width2 + 1, heightlist
    
    treFiles.Move treView.Left, treFiles.Top, treView.Width
    '//move the Splitter bar into position
    Splitter1.Move x1 + treView.Width - 1, intOffsetToplist, SPLT_WDTH, heightlist
End Sub



Private Sub Form_Unload(Cancel As Integer)
 Set creg = Nothing

End Sub

Private Sub lstView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call SortListView(lstView, ColumnHeader.Index)
End Sub



Private Sub lstView_DblClick()
    Call lvVSSMenuEdit_Click
End Sub

Private Sub lstView_ItemCheck(ByVal item As MSComctlLib.IListItem)
Static DataCount As Long
Static count As Long
Static blnset As Boolean
  On Error GoTo ErrHandler
  'StatusBar1.Panels(1).Text = item.Index
  If item.Checked = True Then
      If prevKey <> "" Then
        treView.Nodes(prevKey).parent.Expanded = False
        treView.Nodes(prevKey).BackColor = vbWhite
        treView.Nodes(prevKey).ForeColor = vbBlack
      End If
      treView.Nodes(item.SubItems(1) & "|" & item.Index).EnsureVisible
      treView.Nodes(item.SubItems(1) & "|" & item.Index).BackColor = vbBlue
      treView.Nodes(item.SubItems(1) & "|" & item.Index).ForeColor = vbYellow
      prevKey = item.SubItems(1) & "|" & item.Index
      If ns1.apinfo(item.Index).BSSID = item.SubItems(1) Then
'        If blnset = False Then
            DataCount = ns1.apinfo(item.Index).DataCount
'            blnset = True
'        End If
'        If Count < DataCount Then
            ns1.apinfo(item.Index).DataCount = 0
'            Count = Count + 1
'        End If

      mnuSave_Click
      mnuOpenNet_Click
      ns1.apinfo(item.Index).DataCount = DataCount
      End If
Else
        treView.Nodes(item.SubItems(1) & "|" & item.Index).parent.Expanded = False
        treView.Nodes(item.SubItems(1) & "|" & item.Index).BackColor = vbWhite
        treView.Nodes(item.SubItems(1) & "|" & item.Index).ForeColor = vbBlack
End If
  Exit Sub
ErrHandler:
End Sub



Private Sub lstView_ItemClick(ByVal item As MSComctlLib.ListItem)
  On Error GoTo ErrHandler
  StatusBar1.Panels(1).Text = item.Index
  If prevKey <> "" Then
    treView.Nodes(prevKey).parent.Expanded = False
    treView.Nodes(prevKey).BackColor = vbWhite
    treView.Nodes(prevKey).ForeColor = vbBlack
  End If
  treView.Nodes(item.SubItems(1) & "|" & item.Index).EnsureVisible
  treView.Nodes(item.SubItems(1) & "|" & item.Index).BackColor = vbBlue
  treView.Nodes(item.SubItems(1) & "|" & item.Index).ForeColor = vbYellow
  prevKey = item.SubItems(1) & "|" & item.Index
  Exit Sub
ErrHandler:
End Sub

Private Sub lstView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then bAbort = True
End Sub

Private Sub lstView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      
    Call ListView_AfterEdit(lstView, tHt, tbListViewEdit)
      
    tHt = ListView_HitTest(lstView, x, y)
        
    If Button <> 2 Then Exit Sub
    
    If tHt.lItem = -1 Then
        lvVSSMenuEdit.Enabled = False
    Else
        lvVSSMenuEdit.Enabled = True
        lstView.ListItems(tHt.lItem + 1).Selected = True
    End If
    
    PopupMenu lvVSSMenu
End Sub

Private Sub lvVSSMenuEdit_Click()

        Call ListView_ScaleEdit(lstView, tHt, tbListViewEdit)
        
        Call ListView_BeforeEdit(lstView, tHt, tbListViewEdit)
End Sub

Private Sub mnuAbout_Click()
    dontcheckreg = True
    frmSplash.Show 1
    dontcheckreg = False
End Sub

Private Sub mnuBatch_Click()
    BatchJob = True
    
    On Error Resume Next
      treFiles.Enabled = False
     For Each BatchNode In treFiles.Nodes
      BatchNode.EnsureVisible
      If BatchNode <> BatchNode.Root Then OpenFile BatchNode.Root.Key & BatchNode.Key, BatchIndex, True
        
     Next
     treFiles.Enabled = True
End Sub

Private Sub mnuclearAPData_Click()
Dim lIndex As Long
Dim lData As Long
Dim lcount As Long
    For lIndex = 1 To UBound(ns1.apinfo) - BadRecords.Items.apcount
         ns1.apinfo(lIndex).DataCount = 0
         lcount = lcount + 1
    Next lIndex
    StatusBar1.Panels(1).Text = "Fixed:" & lcount & " Corrupt Data Rates "
    lstView.ListItems.Clear
    treView.Refresh
End Sub

Private Sub mnuEachToOwn_Click()
 Dim selNode As Node
 For Each selNode In treView.Nodes
    If Left(selNode.Key, 1) = "_" Then
    Debug.Print Right(selNode.Key, Len(selNode.Key) - InStrRev(selNode.Key, "_"))
    
    SaveFile ns1, False, Right(selNode.Key, Len(selNode.Key) - InStrRev(selNode.Key, "_"))
    End If
 Next
End Sub

Private Sub mnuFixAltitude_Click()
Dim lIndex As Long
Dim lData As Long
Dim lcount As Long
Dim bytes(7) As Byte

    For lIndex = 1 To UBound(ns1.apinfo) - BadRecords.Items.apcount
        For lData = LBound(ns1.apinfo(lIndex).APData) To UBound(ns1.apinfo(lIndex).APData)
             If ns1.apinfo(lIndex).APData(lData).Location_Source <> 0 Then
                 If ns1.apinfo(lIndex).APData(lData).GPSDATA.Altitude.dbl < 0 Then
                    ns1.apinfo(lIndex).APData(lData).GPSDATA.Altitude.dbl = 0
                    ns1.apinfo(lIndex).APData(lData).GPSDATA.Altitude.bytes = bytes
                    lcount = lcount + 1
                 End If
             End If
        Next lData
    Next lIndex
    StatusBar1.Panels(1).Text = "Fixed:" & lcount & " Bad Altitudes"
    lstView.ListItems.Clear
    treView.Refresh
End Sub

Private Sub mnuFixDataRates_Click()
Dim lIndex As Long
Dim lData As Long
Dim lcount As Long
    For lIndex = 1 To UBound(ns1.apinfo) - BadRecords.Items.apcount
             If ns1.apinfo(lIndex).DataRate = 0 Then
                ns1.apinfo(lIndex).DataRate = 110
                lcount = lcount + 1
             End If
    Next lIndex
    StatusBar1.Panels(1).Text = "Fixed:" & lcount & " Corrupt Data Rates "
    lstView.ListItems.Clear
    treView.Refresh
End Sub

Private Sub mnuFixLocSource_Click()
Dim lIndex As Long
Dim lData As Long
Dim lcount As Long
    For lIndex = 1 To UBound(ns1.apinfo) - BadRecords.Items.apcount
        For lData = LBound(ns1.apinfo(lIndex).APData) To UBound(ns1.apinfo(lIndex).APData)
             If ns1.apinfo(lIndex).APData(lData).Location_Source = 2081750912 Then
                ns1.apinfo(lIndex).APData(lData).Location_Source = 3
                lcount = lcount + 1
             End If
        Next lData
    Next lIndex
    StatusBar1.Panels(1).Text = "Fixed:" & lcount & " Corrupt Location Sources "
    lstView.ListItems.Clear
    treView.Refresh
End Sub

Private Sub mnuExportsingle_Click()
 Dim Index As Long
 Dim i As Long
 Dim bfound As Boolean
 Dim ret As VbMsgBoxResult
 
 Index = Right(LastNode.Key, Len(LastNode.Key) - InStrRev(LastNode.Key, "_"))
    If BadRecords.Items.apcount <> 0 And Index >= BadRecords.indexes(1) Then
       For i = 1 To UBound(BadRecords.indexes)
        If Index = BadRecords.indexes(i) Then
            ret = MsgBox("You have selected to export the corrupted AP to a ns1 file" & vbCrLf & _
                   "The file may not be readable in Netstumbler due to the corruption" & vbCrLf & _
                   "It is suggested that you click on the apinfo for this entry and" & vbCrLf & _
                   "export the data to Excel instead" & vbCrLf & _
                   "Do you still want to export it to a ns1 file?" _
                   , vbYesNo, "File may not be readable")
            If ret = vbYes Then
                bfound = True
                SaveFile BadRecords.Items, False, i
            End If
            Exit For
        End If
       Next i
       If bfound = False Then
       'Wasn't the bad one so offset index
        SaveFile ns1, False, Index - BadRecords.Items.apcount
       End If
    Else
    SaveFile ns1, False, Index
    End If
End Sub

Private Sub mnuInfo_Click()
    Dialog.Show 1
End Sub

Private Sub mnuKML_Click()
    On Error Resume Next
    Dim Obj As cdlg, locfname As String
    '------------------------------------------------------------
    ' This will export the contents of the listview
    ' to an Excel 97 workbook.  You must provide a
    ' full path and filename.  The second argument
    ' is wheather or not you want the workbook to open
    ' after export or not.
    ' Note: even if the user has Office 2000 this will
    ' still work.
    '------------------------------------------------------------
    ' PLEASE NOTE: If the file passed already exists,
    ' it will be overwritten.  It will be your job
    ' to make sure the user wants to.
    '------------------------------------------------------------
    Set Obj = New cdlg
    locfname = Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & ".kml"
    Obj.VBGetSaveFileName locfname, , , "*.kml (Google KML)|*.kml", , CurDir, "Export to", "*.kml", Me.hwnd
    If locfname <> "" And InStr(locfname, "*") = 0 Then
        If UCase(Right(locfname, 4)) <> ".KML" Then locfname = locfname & ".kml"
        Screen.MousePointer = vbHourglass
        ExportToKML locfname

        Screen.MousePointer = vbDefault
    End If



End Sub

Private Sub mnuLiveUpdate_Click()
On Error GoTo ErrHandler
    Shell App.Path & "\Live Update.exe", vbNormalFocus
  Exit Sub
ErrHandler:
MsgBox "Error trying to run " & vbCrLf & App.Path & "\Live Update.exe", vbOKOnly, Err.Description

End Sub

Private Sub mnuLoadList_Click()
Dim commondialog As cdlg
Dim filecount As Long
Dim RootNode    As Node
Dim sRet As String
Dim i As Long
Dim sDir As String
Dim sFiles() As String
Dim sFile As String
Dim blankns1 As ns1
On Error GoTo ErrHandler
Set commondialog = New cdlg

treFiles.Nodes.Clear
treView.Nodes.Clear
lstView.ListItems.Clear
   
   commondialog.VBGetOpenFileName sFile, "Select Files To Open", True, True, False, False, FileFilter, , , , , , OFN_ALLOWMULTISELECT Or OFN_READONLY Or OFN_EXPLORER
   DoEvents
   commondialog.ParseMultiFileName sDir, sFiles(), filecount
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    Set RootNode = treFiles.Nodes.Add(, , sDir, "Stumbles", "fldropn", "fldropn")
    For i = 0 To filecount - 1
        frmMain.treFiles.Nodes.Add RootNode, tvwChild, sFiles(i), sFiles(i), "file", "ap"
    Next
    treFiles.Nodes(1).Expanded = True
    treFiles.Enabled = True
    
Cleanup:
     Set RootNode = Nothing
     Set commondialog = Nothing
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, , "Select Files"
        CloseFile
        Resume
    End If
    Resume Cleanup

End Sub

Private Sub mnuMerge_Click()
Dim commondialog As cdlg
Dim filecount As Long
Dim RootNode    As Node
Dim sRet As String
Dim i As Long
Dim sDir As String
Dim sFiles() As String
Dim sFile As String
Dim blankns1 As ns1
Dim prvcount As Long
Dim loopcnt As Long
Dim lcount As Long
Dim Index As Long

On Error GoTo ErrHandler

Set commondialog = New cdlg
Erase MergedNs1.apinfo()
treFiles.Enabled = False

treFiles.Nodes.Clear
treView.Nodes.Clear
lstView.ListItems.Clear
   bAbort = False
   cmdAbort.Visible = True
   
   commondialog.VBGetOpenFileName sFile, "Select Files To Merge", True, True, False, False, "NetStumbler Files (*.ns1)|*.ns1", , , , , , OFN_ALLOWMULTISELECT Or OFN_READONLY Or OFN_EXPLORER
   DoEvents
   commondialog.ParseMultiFileName sDir, sFiles(), filecount
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    Set RootNode = treFiles.Nodes.Add(, , sDir, "Stumbles", "fldrcls", "fldropn")
    For i = 0 To filecount - 1
        frmMain.treFiles.Nodes.Add RootNode, tvwChild, sFiles(i), sFiles(i), "file", "ap"
    Next
    treFiles.Nodes(1).Expanded = True
    If IsNull(sFiles) Then Exit Sub
    
For i = LBound(sFiles) To UBound(sFiles)
        If bAbort Then Exit For
 
     ns1 = blankns1
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
        treFiles.Nodes(sFiles(i)).Selected = True
        fname = sDir & sFiles(i)
        RecoverFile True
        DoEvents
        Call AddStringArrays(ns1.apinfo, MergedNs1.apinfo)
        DoEvents
Next i
        
        MergedNs1.apcount = UBound(MergedNs1.apinfo)
        MergedNs1.dwFileVer = 12
        MergedNs1.dwSignature = "NetS"
        
        SaveFile ns1, True, , MergedNs1.apcount
Cleanup:
     Set RootNode = Nothing
     Set commondialog = Nothing
     cmdAbort.Visible = False
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        Resume Next
        MsgBox Err.Description, , "Select Files"
        CloseFile
        
        'Resume Cleanup
    End If
    Resume Cleanup
End Sub
Private Sub AddStringArrays(ByRef arSrc() As apinfo, ByRef arDest() As apinfo)
    'Add arSrc Array to arDest Array
    
    On Error GoTo ErrHandler
    
    Dim lngMaxDest As Long
    Dim lngMaxSrc As Long
    Dim lngMax As Long
    Dim lngCnt As Long
    Dim lngStart As Long
    Dim lngCurrent As Long
    
    lngMaxSrc = UBound(arSrc)
    
    If Not IsArrayInit(arDest) Then
        lngMaxDest = lngMaxSrc
        lngMax = lngMaxSrc
        lngStart = 0
    Else
        lngMaxDest = UBound(arDest)
        lngMax = lngMaxDest + lngMaxSrc + 1
        lngStart = lngMaxDest + 1
    End If
    
    ReDim Preserve arDest(lngMax)
    
    lngCurrent = LBound(arSrc)
    For lngCnt = lngStart To lngMax
        arDest(lngCnt) = arSrc(lngCurrent)
        lngCurrent = lngCurrent + 1
    Next
    
exitHandler:
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbOKOnly, "Error: " & Err.Number
    Resume exitHandler
    
End Sub


Private Sub mnuOpenNet_Click()
    Shell "C:\Program Files\Network Stumbler\NetStumbler.exe " & Chr(34) & Left(fname, Len(fname) - 4) & "_Recovered.ns1" & Chr(34)
End Sub

Private Sub mnuOpenSavens1_Click()
Dim commondialog As cdlg
Dim filecount As Long
Dim RootNode    As Node
Dim sRet As String
Dim i As Long
Dim sDir As String
Dim sFiles() As String
Dim sFile As String
Dim blankns1 As ns1
On Error GoTo ErrHandler
Set commondialog = New cdlg
treFiles.Enabled = False

treFiles.Nodes.Clear
treView.Nodes.Clear
lstView.ListItems.Clear
   bAbort = False
   cmdAbort.Visible = True
   
   commondialog.VBGetOpenFileName sFile, "Select Files To Open", True, True, False, False, "NetStumbler Files (*.ns1)|*.ns1", , , , , , OFN_ALLOWMULTISELECT Or OFN_READONLY Or OFN_EXPLORER
   DoEvents
   commondialog.ParseMultiFileName sDir, sFiles(), filecount
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    Set RootNode = treFiles.Nodes.Add(, , sDir, "Stumbles", "fldrcls", "fldropn")
    For i = 0 To filecount - 1
        frmMain.treFiles.Nodes.Add RootNode, tvwChild, sFiles(i), sFiles(i), "file", "ap"
    Next
    treFiles.Nodes(1).Expanded = True
    If IsNull(sFiles) Then Exit Sub
    For i = LBound(sFiles) To UBound(sFiles)
        If bAbort Then Exit For
 
     ns1 = blankns1
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
        
        treFiles.Nodes(sFiles(i)).Selected = True
        fname = sDir & sFiles(i)
        RecoverFile True
        DoEvents
        SaveFile ns1
        DoEvents
    Next i
    
Cleanup:
     Set RootNode = Nothing
     Set commondialog = Nothing
     cmdAbort.Visible = False
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, , "Select Files"
        CloseFile
        Resume Cleanup
    End If
    Resume Cleanup
End Sub
Private Sub mnuopnwiscan_Click()
Dim commondialog As cdlg
Dim filecount As Long
Dim RootNode    As Node
Dim sRet As String
Dim i As Long
Dim sDir As String
Dim sFiles() As String
Dim sFile As String
Dim blankns1 As ns1
On Error GoTo ErrHandler
Set commondialog = New cdlg
treFiles.Enabled = False

treFiles.Nodes.Clear
treView.Nodes.Clear
lstView.ListItems.Clear
   bAbort = False
   cmdAbort.Visible = True
   
   commondialog.VBGetOpenFileName sFile, "Select Files To Open", True, True, False, False, "NetStumbler Files (*.ns1)|*.ns1", , , , , , OFN_ALLOWMULTISELECT Or OFN_READONLY Or OFN_EXPLORER
   DoEvents
   commondialog.ParseMultiFileName sDir, sFiles(), filecount
    If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
    Set RootNode = treFiles.Nodes.Add(, , sDir, "Stumbles", "fldrcls", "fldropn")
    For i = 0 To filecount - 1
        frmMain.treFiles.Nodes.Add RootNode, tvwChild, sFiles(i), sFiles(i), "file", "ap"
    Next
    treFiles.Nodes(1).Expanded = True
    If IsNull(sFiles) Then Exit Sub
    For i = LBound(sFiles) To UBound(sFiles)
        If bAbort Then Exit For
 
     ns1 = blankns1
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
        
        treFiles.Nodes(sFiles(i)).Selected = True
        fname = sDir & sFiles(i)
        RecoverFile True
        DoEvents
        If InStr(1, fname, ".") Then
            SaveWiscan Left(fname, InStr(1, fname, ".") - 1)
        Else
            SaveWiscan fname
        End If
        DoEvents
    Next i
    
Cleanup:
     Set RootNode = Nothing
     Set commondialog = Nothing
     cmdAbort.Visible = False
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
        Exit Sub
    Else
        MsgBox Err.Description, , "Select Files"
        CloseFile
        Resume Cleanup
    End If
    Resume Cleanup
End Sub

Private Sub mnuOptions_Click()
 frmOptions.Show 1
End Sub

Private Sub mnuPrint_Click()
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
 
PrintListView lstView, PaperSize, Orientation, LeftMargin, RightMargin, TopMargin, BottomMargin
End Sub

Private Sub mnuRemoveRecord_Click()
 ns1.apinfo(Right(LastNode.Key, Len(LastNode.Key) - InStr(1, LastNode.Key, "|", vbTextCompare))).DataCount = 0
End Sub

Private Sub mnuSave_Click()
    SaveFile ns1
End Sub


Private Sub mnuwififofum_Click()
    OpenFile , , True
End Sub

Private Sub mnuWiscan_Click()
    On Error Resume Next
    Dim Obj As cdlg, locfname As String
    '------------------------------------------------------------
    ' This will export the contents of the listview
    ' to an Excel 97 workbook.  You must provide a
    ' full path and filename.  The second argument
    ' is wheather or not you want the workbook to open
    ' after export or not.
    ' Note: even if the user has Office 2000 this will
    ' still work.
    '------------------------------------------------------------
    ' PLEASE NOTE: If the file passed already exists,
    ' it will be overwritten.  It will be your job
    ' to make sure the user wants to.
    '------------------------------------------------------------
    Set Obj = New cdlg
    locfname = Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1)
    Obj.VBGetSaveFileName locfname, , , , , CurDir, "Export to", , Me.hwnd
    If fname <> "" And InStr(fname, "*") = 0 Then
        Screen.MousePointer = vbHourglass
        SaveWiscan locfname
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Sub Splitter1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        '//change the Splitter color
        Let Splitter1.BackColor = SplitColor
        
        '//set the current position to x
        Let currSplitPosX = CLng(x)
    Else
        '//not the left button, so... if the current position <> default, cause a mouseup
        If currSplitPosX <> SPLT_DEFAULT_POS Then Splitter1_MouseUp Button, Shift, x, y
        
        '//set the current position to the default value
        Let currSplitPosX = SPLT_DEFAULT_POS
    End If
End Sub

 Private Sub Splitter1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//if the Splitter has been moved...
    If currSplitPosX& <> SPLT_DEFAULT_POS Then
        '//if the current position <> default,
        '//reposition the Splitter and set this as the current value
        If CLng(x) <> currSplitPosX Then
            Splitter1.Move Splitter1.Left + x, intOffsetToplist, SPLT_WDTH + 15, _
                                ScaleHeight - intOffsetToplist - intOffsetBottom
            Let currSplitPosX = CLng(x)
        End If
    End If
End Sub

Private Sub Splitter1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    '//if the Splitter has been moved...
    If currSplitPosX <> SPLT_DEFAULT_POS Then
        '//if the current postition <> the last position do a final move of the Splitter
        If CLng(x) <> currSplitPosX Then
            Splitter1.Move Splitter1.Left + x, intOffsetToplist, SPLT_WDTH, _
                            ScaleHeight - intOffsetToplist - intOffsetBottom
        End If
        
        '//call this the default position
        Let currSplitPosX = SPLT_DEFAULT_POS
        
        '//restore the normal Splitter color
        Let Splitter1.BackColor = Me.BackColor '//&H8000000F
        
        '//and check for valid sizings.
        '//Either enforce the default minimum & maximum widths for the left list,

        '//or, if within range, set the width
        If Splitter1.Left > 60 And Splitter1.Left < (ScaleWidth - 60) Then
            '//the pane is within range
            treView.Width = Splitter1.Left - treView.Left
            ElseIf Splitter1.Left < 60 Then '//the pane is too small
                treView.Width = 60
                treFiles.Width = 60
            Else
                treView.Width = ScaleWidth - 60 '//the pane is too wide
                treFiles.Width = ScaleWidth - 60 '//the pane is too wide
        End If
            '//reposition both lists, and the Splitter bar
            Call Form_Resize
    End If
End Sub
Private Sub Frame1_HeightChanged(ByVal NewHeight As Single)
    '//when the height of the coolbar changes, we have to reposition the controls
    Call Form_Resize
End Sub
Private Sub mnuExcel_Click()
'    ExportToExcel lstView, True
    On Error Resume Next
    Dim Obj As cdlg, locfname As String
    '------------------------------------------------------------
    ' This will export the contents of the listview
    ' to an Excel 97 workbook.  You must provide a
    ' full path and filename.  The second argument
    ' is wheather or not you want the workbook to open
    ' after export or not.
    ' Note: even if the user has Office 2000 this will
    ' still work.
    '------------------------------------------------------------
    ' PLEASE NOTE: If the file passed already exists,
    ' it will be overwritten.  It will be your job
    ' to make sure the user wants to.
    '------------------------------------------------------------
    Set Obj = New cdlg
    locfname = Mid(fname, InStrRev(fname, "\") + 1, InStrRev(fname, ".") - InStrRev(fname, "\") - 1) & ".xls"
    Obj.VBGetSaveFileName locfname, , , "*.xls (Excel Workbook)|*.xls", , CurDir, "Export to", "*.xls", Me.hwnd
    If locfname <> "" And InStr(locfname, "*") = 0 Then
        If UCase(Right(locfname, 4)) <> ".XLS" Then locfname = locfname & ".xls"
        Screen.MousePointer = vbHourglass
        ExportToExcel locfname, lstView, True
        
        'lstView.ExportToExcel fname, True
        '------------------------------------------------------------
        '------------------------------------------------------------
        '------------------------------------------------------------
        ' Also go see my ExportToExcelComplete Event
        '------------------------------------------------------------
        '------------------------------------------------------------
        '------------------------------------------------------------
        Screen.MousePointer = vbDefault
    End If


End Sub

Private Sub mnuExport_Click()
   SaveLW lstView, fname & ".csv"
End Sub

Private Sub mnuOpen_Click()
 OpenFile
End Sub
Public Sub OpenFile(Optional FileName As String, Optional ByRef Index As Long = 0, Optional isBatch As Boolean = False)
Dim ret As VbMsgBoxResult
Dim step As Long
On Error GoTo File_Open_Errorhandler
 'Dim blankns1 As ns1
 'ns1 = blankns1
     
 If Index = 0 Then
    treView.Enabled = False
    treView.Nodes.Clear
    lstView.ListItems.Clear
    BadRecords.Items.apcount = 0
 
    mnuOpenNet.Enabled = False
    Erase BadRecords.indexes
    ReDim BadRecords.indexes(0)
 End If
 treFiles.Enabled = False
 
fname = FileName
If fname <> "" Then
  Select Case LCase$(Right$(fname, 3))
    Case "ns1" 'NetStumbler ns1
        CommonDialog1.FilterIndex = 1
    Case "csv" 'Kismet csv
        CommonDialog1.FilterIndex = 3
    Case "wpt" 'OziExplorer file
        CommonDialog1.FilterIndex = 5
    Case Else 'Other Wiscan
        CommonDialog1.FilterIndex = 4
End Select
End If
    CommonDialog1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNNoReadOnlyReturn
    CommonDialog1.Filter = FileFilter
    CommonDialog1.DialogTitle = "Open File"
    CommonDialog1.InitDir = App.Path
    CommonDialog1.CancelError = True
Select Case True
    Case fname = ""
        CommonDialog1.ShowOpen
        fname = CommonDialog1.FileName   ' Set file.
        If fname = "" Then Exit Sub
    Case Dir(fname) = ""
        MsgBox fname
        ret = MsgBox("Would you like to open a different file?", vbQuestion + vbYesNo, "File Does not Exist")
        fname = ""
        If ret = vbNo Then Exit Sub
        CommonDialog1.ShowOpen
        fname = CommonDialog1.FileName   ' Set  file.
        If fname = "" Then Exit Sub
End Select
    lblFilename.Caption = fname
    zzExport.Enabled = True
    mnuSave.Enabled = True
    mnuFixes.Enabled = True
    treFiles.Enabled = True
  
  Select Case True
    Case CommonDialog1.FilterIndex = 1 Or LCase$(Right$(fname, 3)) = "ns1" 'NetStumbler ns1
        RecoverFile , False, isBatch
    Case CommonDialog1.FilterIndex = 2 'Wififofum ns1 with bad header
        RecoverFile , True, isBatch
    Case CommonDialog1.FilterIndex = 3 Or LCase$(Right$(fname, 3)) = "csv" 'Kismet csv
        Read_Kismet fname, isBatch
    Case CommonDialog1.FilterIndex = 6 Or LCase$(Right$(fname, 3)) = "wpt" 'Ozi Explorer
        Read_OZI fname, isBatch
     Case CommonDialog1.FilterIndex = 7 Or LCase$(Right$(fname, 3)) = "txt" 'Pontisoft Sniffi
        Read_Sniffi fname, isBatch
    Case Else 'Other Wiscan summary
        Read_WiScan fname, isBatch
  End Select
  treView.Enabled = True
    
  Index = Index + 1
  StatusBar1.Panels(1).Text = ns1.apcount & " records retrieved/verified successfully... Ready"
  lblHeaderValue(2).Caption = ns1.apcount
  Exit Sub
File_Open_Errorhandler:
    CloseFile
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox Err.Description, vbExclamation + vbOKOnly, "File Open Error"
        Resume
    End If
End Sub
Sub SaveWiscan(rfname As String)
    
    'Save all your hard work!
    Dim FNamebak As String
    Dim Counter As Long
    Dim apcount As Long
    Dim strtext As String
    Dim ret As Long
    Dim l As Long
    'Set filenames

'    rfname = Left(rfname, Len(rfname) - 4)
    
   
    'if the file is there, get rid of it
    'this is a good spot to add some code to backup the original file
    If Dir(rfname) <> "" Then
        Kill rfname
    End If
    
    StatusBar1.Panels(1).Text = "Exporting Records..."
    
    'Output all data to the file
    
    OpentxtFile rfname
    PutText "# $Creator: NS1 Recovery " & App.Major & "." & App.Minor & "." & App.Revision
    PutText "# $Format: wi-scan summary with extensions"
    PutText "# Latitude  Longitude   ( SSID )    Type    ( BSSID )   Time (GMT)  [ SNR Sig Noise ]   # ( Name )   Flags Channelbits BcnIntvl    DataRate    LastChannel"
    PutText "# $DateGMT: " & Format(UtcFromLocalTime(ns1.apinfo(1).firstseen.Time), "YYYY-MM-DD")

      For apcount = 1 To UBound(ns1.apinfo) - BadRecords.Items.apcount
        StatusBar1.Panels(1).Text = "Exporting Record: " & apcount
        With ns1.apinfo(apcount)
        strtext = ValToDms(.BestLat.dbl, True, True) & vbTab & _
                  ValToDms(.BestLong.dbl, False, True) & vbTab & _
                  "( " & .SSID & " )" & vbTab & _
                  IIf(Right(.flags, 1) = 2, "ad-hoc", "BSS") & vbTab & _
                  "( " & LCase(.BSSID) & " )" & vbTab & _
                  Format(UtcFromLocalTime(.firstseen.Time), "HH:MM:SS") & " (GMT)" & vbTab & _
                  "[ " & .MaxSNR & " " & .MaxSignal + 149 & " " & (.MaxSignal + 149) - .MaxSNR & " ]" & vbTab & _
                  "# ( " & .Name & " )" & vbTab & _
                  Format(CDToH(.flags), "00##") & vbTab & _
                  Channelbits(.Channels.str) & vbTab & _
                  .BeaconInterval & vbTab & _
                  .DataRate & vbTab & _
                  .LastChannel
        'Debug.Print .Channels.str
        
        End With
        PutText strtext
    Next apcount
    CloseFile
    mnuOpenNet.Enabled = True
    StatusBar1.Panels(1).Text = ns1.apcount - 1 & " records saved successfully... Ready"

End Sub


Sub Set_Up_Panels()
'***************************************************************************************************
'Set up the panels that are on the bottom of the screen
'***************************************************************************************************
Dim i As Integer
Dim p_IntPanel As Integer
    With StatusBar1.Panels
        .item(1).Style = sbrText
        .item(1).Text = "Current File  " & fname
        .item(2).Style = sbrTime
        .item(3).Style = sbrDate
    End With
'Progress bar to Statusbar
    p_IntPanel = 4
    m_ClsAddProgToSbar.AddPBtoSB StatusBar1.hwnd, ProgressBar1.hwnd, hwnd, p_IntPanel
    ' This text will be displayed when the StatusBar is in Simple style.
    StatusBar1.SimpleText = "Date and Time: " & Now
    'StatusBar1.Style = sbrSimple
End Sub

Public Function FillTreeview(Index As Long, Optional icoindex As Integer = 5, Optional bIgnoreBRC As Boolean = False)
    On Error GoTo ErrHandler
    Dim badrecs As Long
    If Not bIgnoreBRC Then badrecs = BadRecords.Items.apcount
    
    Set ParentNode = treView.Nodes("_" & ns1.apinfo(Index - badrecs).SSID & "_" & Index)
   ' ParentNode.EnsureVisible
    Set ChildNode = treView.Nodes.Add(ParentNode, tvwChild, ns1.apinfo(Index - badrecs).BSSID & "|" & Index, ns1.apinfo(Index - badrecs).BSSID, 6, 6)
     treView.Nodes.Add ChildNode, tvwChild, "APData" & "|" & Index, "ApData", 7, 7
    
    Exit Function
ErrHandler:
        If Err.Number = 35601 Then
            'Element not found: Parent doesn 't exist yet so create it
            Set ParentNode = treView.Nodes.Add(RootNode, tvwChild, "_" & ns1.apinfo(Index - badrecs).SSID & "_" & Index, ns1.apinfo(Index - badrecs).SSID, icoindex, icoindex)
        
            Set ChildNode = treView.Nodes.Add(ParentNode, tvwChild, ns1.apinfo(Index - badrecs).BSSID & "|" & Index, ns1.apinfo(Index - badrecs).BSSID, 6, 6)
            treView.Nodes.Add ChildNode, tvwChild, "APData" & "|" & Index, "ApData", 7, 7
        End If
End Function


Private Sub RecoverFile(Optional AutoMode As Boolean = False, Optional wififofum As Boolean = False, Optional isBatch As Boolean = False)
On Error GoTo ErrHandler
Dim ProMax As Long
Dim Counter As Long
Dim bytes() As Byte
Dim Offset As Long
Dim blankns1 As ns1
Dim apcount As Long
Dim step As Long
Dim apdone As Boolean
Dim ret As Long
Dim bcorrupt As Boolean

    bLoading = True
    If (Not isBatch) Or (frmMain.BatchIndex = 0 And isBatch) Then
        ns1 = blankns1
        PrevCount = 0
        frmMain.treView.Nodes.Clear
        Set RootNode = frmMain.treView.Nodes.Add(, , "Root", "SSID", 2, 2)
        frmMain.lstView.ListItems.Clear
    End If
    
    OpenBinFile
    ret = Read_ns1Header(Offset, wififofum)
    If ret <> 0 Then Exit Sub
    prevKey = ""
    lblHeaderValue(0).Caption = ns1.dwSignature
    lblHeaderValue(1).Caption = ns1.dwFileVer
    lblHeaderValue(2).Caption = ns1.apcount
    

    
    step = 1
    ReDim Preserve ns1.apinfo(ns1.apcount + PrevCount)
    If ns1.dwFileVer < 6 Then
            MsgBox "Your Ns1 File version: " & ns1.dwFileVer & " is too old for this program. " & vbCrLf & _
                "It's not coded to handle vesions older then File Version 6", vbCritical, "Unsupported Version"
            GoTo Cleanup
    Else
        For apcount = 1 To ns1.apcount
            StatusBar1.Panels(1).Text = "Verifying Records: " & apcount
            DrawProgress ns1.apcount, apcount
            apdone = False
            bcorrupt = False
            Do While apdone = False
               ret = Read_ApInfo((apcount - BadRecords.Items.apcount) + PrevCount, Offset, step, apdone, bcorrupt, wififofum)
               If ret <> 0 And ret <> 9999 Then Exit For
            Loop
            If Not AutoMode Then
                If Not bcorrupt Then
                    If ret <> 9999 Then FillTreeview apcount + PrevCount
                Else
                    Debug.Print BadRecords.Items.apcount
                    
                    If ret <> 9999 Then FillTreeview apcount + PrevCount, 12, True
                End If
            End If
            DoEvents
            If bAbort Then GoTo Cleanup
        Next apcount

    End If
    
    ns1.apcount = ns1.apcount + PrevCount
    PrevCount = ns1.apcount
    lblHeaderValue(3).Caption = BadRecords.Items.apcount

Cleanup:
    bLoading = False
    bAbort = False
    Me.MousePointer = vbDefault
    CloseFile
    Exit Sub
    
ErrHandler:
    Dim free As Long
    free = FreeFile
    MsgBox "Error Number: " & Err.Number & vbCrLf & Err.Description & vbCrLf & _
           "Please email author ns1 file and " & App.Path & "\Errorlog.txt", vbCritical, "Error"
'    On Error Resume Next
    Open App.Path & "\Errorlog.txt" For Output As #free
    Print #free, "----------------------------------------"
    Print #free, "File: " & fname
    Print #free, "Signature: " & ns1.dwSignature
    Print #free, "Version: " & ns1.dwFileVer
    Print #free, "Ap Count: " & ns1.apcount
    Print #free, Err.Description & vbCrLf & "At Step " & step
    Print #free, "Index: " & 9999
    Print #free, "Offset: " & Offset
    Print #free, "SSID: " & ns1.apinfo(0).SSID
    Print #free, "+++++++++++++++++++++++++++++++++++++++++"
    Close #free
    If gDebugMode Then
        Debug.Print Err.Description
        Resume
    Else
        Resume Cleanup
    End If

End Sub


Public Sub DrawProgress(Maxvalue As Long, step As Long, Optional DrawColor As Long = &HECBB68)
    Dim sngPercentWidth As Single
    Dim sngDrawWidth As Single
    Dim PercentDone As Single
    If Not m_ClsAddProgToSbar Is Nothing Then
        m_ClsAddProgToSbar.RefreshProgressBar
    End If
    PercentDone = (step / Maxvalue) * 100
    ' Check for valid input
    If PercentDone < 0 Then PercentDone = 0
    If PercentDone > 100 Then PercentDone = 100
    If PercentDone > 0 Then
        StatusBar1.Panels(5).Text = Format(PercentDone, "###") & "%"
    Else
        StatusBar1.Panels(5).Text = ""
    End If
    ' Allow for decimal representation of percent
    'If PercentDone < 1 Then PercentDone = PercentDone * 100

    ' Determine the width of one percent
    sngPercentWidth = ProgressBar1.ScaleWidth / 100

    ' Determine the width to draw
    sngDrawWidth = PercentDone * sngPercentWidth

    ' Fill the picturebox
    If PercentDone = 0 Then
        ProgressBar1.Line (0, 0)-(ProgressBar1.ScaleWidth, ProgressBar1.Height), &H80000004, BF
    Else
        ProgressBar1.Line (0, 0)-(sngDrawWidth, ProgressBar1.Height), DrawColor, BF
    End If
    
    ' Without Refresh, the progress is not displayed until done
    ProgressBar1.parent.Refresh
End Sub

Private Sub tbListViewEdit_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case vbKeyEscape
        tbListViewEdit.Visible = False
    Case vbKeyReturn
        Call ListView_AfterEdit(lstView, tHt, tbListViewEdit)
    End Select
    
End Sub

Private Sub tbListViewEdit_LostFocus()

    Dim bNextItem As Boolean
    bNextItem = False
    
    If tbListViewEdit.Visible = True Then
        bNextItem = True
    End If
    
    Call ListView_AfterEdit(lstView, tHt, tbListViewEdit)
        
    If bNextItem = True Then
            lstView.ListItems(tHt.lItem + 1).Selected = True
            Call lvVSSMenuEdit_Click
    End If
    
End Sub



Private Sub treFiles_NodeClick(ByVal Node As MSComctlLib.Node)
     Dim blankns1 As ns1
     If Node.Index = 1 Then Exit Sub
     ns1 = blankns1
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
     OpenFile treFiles.Nodes(1).Key & Node.Key
End Sub

Private Sub treFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim blankns1 As ns1
     Dim XY As Long
     Dim TreeRoot As Node
     ns1 = blankns1
     treFiles.Nodes.Clear
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
     PrevCount = 0
     On Error Resume Next
    ' check to see source of drop event.
    Select Case Effect
        Case 7
        If Data.Files.count = 1 Then
            OpenFile Data.Files(1)
        Else
            If treFiles.Nodes.count = 0 Then Set TreeRoot = treFiles.Nodes.Add(, , Mid(Data.Files(1), 1, InStrRev(Data.Files(1), "\")), "Stumbles", "fldropn", "fldropn")
            For XY = 1 To Data.Files.count
                frmMain.treFiles.Nodes.Add TreeRoot, tvwChild, Right(Data.Files(XY), Len(Data.Files(XY)) - InStrRev(Data.Files(XY), "\")), Right(Data.Files(XY), Len(Data.Files(XY)) - InStrRev(Data.Files(XY), "\")), "file", "ap"
            Next
            treFiles.Nodes(1).Expanded = True
            treFiles.Enabled = True
            mnuBatch_Click
        End If
    End Select
    Set TreeRoot = Nothing
'    treFiles.OLEDropMode = ccOLEDropNone
'    treView.OLEDropMode = ccOLEDropNone
End Sub

Private Sub treView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim blankns1 As ns1
     Dim XY As Long
     Dim TreeRoot As Node
     ns1 = blankns1
     treFiles.Nodes.Clear
     treView.Nodes.Clear
     lstView.ListItems.Clear
     BadRecords.Items.apcount = 0
     mnuOpenNet.Enabled = False
     Erase BadRecords.indexes
     ReDim BadRecords.indexes(0)
     PrevCount = 0
     On Error Resume Next
    ' check to see source of drop event.
    Select Case Effect
        Case 7
        If Data.Files.count = 1 Then
            OpenFile Data.Files(1)
        Else
            If treFiles.Nodes.count = 0 Then Set TreeRoot = treFiles.Nodes.Add(, , Mid(Data.Files(1), 1, InStrRev(Data.Files(1), "\")), "Stumbles", "fldropn", "fldropn")
            For XY = 1 To Data.Files.count
                frmMain.treFiles.Nodes.Add TreeRoot, tvwChild, Right(Data.Files(XY), Len(Data.Files(XY)) - InStrRev(Data.Files(XY), "\")), Right(Data.Files(XY), Len(Data.Files(XY)) - InStrRev(Data.Files(XY), "\")), "file", "ap"
            Next
            treFiles.Nodes(1).Expanded = True
            treFiles.Enabled = True
            mnuBatch_Click
        End If
    End Select
    Set TreeRoot = Nothing
'    treFiles.OLEDropMode = ccOLEDropNone
'    treView.OLEDropMode = ccOLEDropNone
End Sub

Private Sub treView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = vbRightButton Then
        If lstView.ListItems.count = 0 Or LastNode.Key = "Root" Then
           
            mnuExportsingle.Visible = False
        End If
         If Left(LastNode.Key, 1) = "_" Then
            mnuExportsingle.Visible = True
            mnuRemoveRecord.Visible = False
         Else
            mnuExportsingle.Visible = False
            mnuRemoveRecord.Visible = True

         End If
         Me.PopupMenu zzExplore
        End If
            
    
End Sub

Private Sub TreView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim icolumn As Integer
Dim iRow As Long
Dim ret As VbMsgBoxResult
Dim Index As Long
Dim i As Long
Dim bfound As Boolean
 
Set LastNode = Node

Dim liTag As ListItem
If Node.Key = "" Then Exit Sub
lstView.ListItems.Clear
If Node.Key = "Root" Then
    For iRow = LBound(ns1.apinfo) To UBound(ns1.apinfo) - BadRecords.Items.apcount
        RefillListBox ns1, lstView, iRow, True
    Next iRow
Else
    If InStr(1, Node.Key, "|", vbTextCompare) <> 0 Then
        Select Case Left(Node.Key, 6)
         Case "APData"
            gphIndex = Right(Node.parent.Key, Len(Node.parent.Key) - InStr(1, Node.parent.Key, "|", vbTextCompare))
            RefillListBox ns1, lstView, gphIndex, False, 1
            If Not IsNothing(ns1.apinfo(gphIndex).APData()) Then
            
                frmGraph.Show
                frmGraph.Command1_Click
            End If

         Case Else
           Index = Right(Node.Key, Len(Node.Key) - InStr(1, Node.Key, "|", vbTextCompare))
           If UBound(BadRecords.indexes) <> 0 Then
           If BadRecords.Items.apcount <> 0 And Index >= BadRecords.indexes(1) Then
             For i = 1 To UBound(BadRecords.indexes)
                If Index = BadRecords.indexes(i) Then
                    bfound = True
                    RefillListBox BadRecords.Items, lstView, i
                    Exit For
                End If
             Next i
             If bfound = False Then
               'Wasn't the bad one so offset index
               RefillListBox ns1, lstView, Index - BadRecords.Items.apcount
             End If
           Else
             RefillListBox ns1, lstView, Index
           End If
           End If
        End Select
    End If
End If

AutosizeColumns lstView
Set liTag = Nothing
End Sub



Private Function IsNothing(ary() As APData) As Boolean
    On Error GoTo ErrHandler
    Dim l As Long
        l = UBound(ary)
        
Exit Function
ErrHandler:
    If Err.Number = 9 Then
        IsNothing = True
    End If
End Function
'checks to see if we've got an image file
Function FileFormatCheck(strFileName As String, Default_Ext As String) As Boolean
    Dim strExtention As String
    
    'grab the file's extention
    strExtention = Right(strFileName, 3)
     
    'check the extention for an image type
    If UCase(strExtention) = UCase(Default_Ext) Then
        FileFormatCheck = True
    Else
        FileFormatCheck = False
    End If
        
End Function


Private Sub AutosizeColumns(ByVal TargetListView As ListView)

  Const SET_COLUMN_WIDTH    As Long = 4126
  Const AUTOSIZE_USEHEADER  As Long = -2

  Dim lngColumn As Long

  For lngColumn = 0 To (TargetListView.ColumnHeaders.count - 1)
   
    Call SendMessage(TargetListView.hwnd, _
                     SET_COLUMN_WIDTH, _
                     lngColumn, _
                     ByVal AUTOSIZE_USEHEADER)
        
  Next lngColumn

End Sub


