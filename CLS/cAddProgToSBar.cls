VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAddProgToSBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const SW_HIDE = 0
Private Const SW_SHOW = 5
Private Const m_WM_USER = &H400
Private Const m_SB_GETRECT = (m_WM_USER + 10)

'Also used w/ SB_GETRECT below
Private Type m_RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
 End Type

'Api Decs
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As m_RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As m_RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

'Handles
Private m_HwndSBar As Long
Private m_HwndProgBar As Long
Private m_HwndParent As Long

'User defined variables
Private m_BlnSbarSet As Boolean
Private m_IntPanal As Integer
Private m_RC As m_RECT

Public Function AddPBtoSB(p_LngHwndSBar As Long, _
            p_LngHwndProgBar As Long, _
            p_LngHWndOriginalParent As Long, _
            p_IntPanel As Integer)

    m_HwndProgBar = p_LngHwndProgBar
    m_HwndSBar = p_LngHwndSBar
    m_HwndParent = p_LngHWndOriginalParent
    m_IntPanal = p_IntPanel - 1
    'Set the parent
    SetParent m_HwndProgBar, m_HwndSBar
    'Flag to show that the progress bar is set
    m_BlnSbarSet = True
    'Redraw the bar
    RefreshProgressBar
    
End Function

Public Sub RefreshProgressBar()
    'Reposition the progbar if required

    'this function is best called from the statusbars parent paint event
    'That way the status bar has recalculated the panal sizes

    If m_BlnSbarSet = False Then Exit Sub
    Call SendMessage(m_HwndSBar, m_SB_GETRECT, m_IntPanal, m_RC)
    InflateRect m_RC, -1, -1
    MoveWindow m_HwndProgBar, m_RC.Left, m_RC.Top, m_RC.Right - m_RC.Left, m_RC.Bottom - m_RC.Top, 1
   'Display Window if not visible
    ShowWindow m_HwndProgBar, SW_SHOW

End Sub
Public Function RemoveProgBarFromSBar()
    SetParent m_HwndProgBar, m_HwndParent
    'Hide the Progress Bar
    ShowWindow m_HwndProgBar, SW_HIDE
    'Reset the flag
    m_BlnSbarSet = False
End Function

Public Property Get ItemLoaded() As Boolean
    ItemLoaded = m_BlnSbarSet
End Property

Private Sub Class_Terminate()
    RemoveProgBarFromSBar
End Sub
