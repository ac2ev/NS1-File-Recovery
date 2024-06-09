VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListViewPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'+==+==+==+==+==+==+==+==+==+==+==+==+==+
'  ListView Printer Class
'  Author: Opal Raj Ghimire, buna48@hotmail.com  http://geocities.com/opalraj/vb
'  You are free to use this class in your projects.
'  Credit me where it deserves
'  A notification is appreciated.
'+==+==+==+==+==+==+==+==+==+==+==+==+==+
'  This class comes with absolutely
'  NO WAR ranty ! Use it at your own risk !!!
'  May posess errors
'+==+==+==+==+==+==+==+==+==+==+==+==+==+

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public HasPicture As Boolean: Public PicHeight As Long: Public PicWidth As Long
Public RowFrom As Long: Public RowTo As Long: Public RowHeight As Long: Public LastRowPrinted As Boolean
Public NumOfRowsPerPage As Integer: Dim lngLastRow As Long
Public DrawHorizontalLines As Boolean: Public DrawVerticalLines As Boolean: Dim bHeadPrinted As Boolean
Public BorderDistance As Long: Public DrawBorder As Boolean
Public PosX As Long: Public PosY As Long:
' Millimeters per twip
Const mmPerTwip As Single = 25.4 / 1440#

Dim objListView As Object:
Dim PRN As Object
Dim ColPerPage As Long
Dim lngWidthHor As Long, lngWidthVer As Long, lngWidthBor As Long
Dim lngColorHor As Long, lngColorVer As Long, lngColorBor As Long
' The useful width of the page (page width minus margins)
Dim TotalWidth As Single
' The page's margins
Dim TopMargin As Single, LeftMargin As Single
Dim PageHeight As Single, PageWidth As Single
' Top of page at YMin, bottom of page at YMax
Dim YMin As Single, YMax As Single
Private Const DT_END_ELLIPSIS = 32768: Const DT_MODIFYSTRING = 65536
Dim MorePages As Boolean
' The index on the ListView control
' of the row being printed
Dim currentRow As Integer
Enum LineID
        VerticalLines = 1
        HorizontalLines = 2
        BorderLines = 3
End Enum
'local variable(s) to hold property value(s)
Private mvarCurrentY As Long 'local copy
Private mvarCurrentX As Long 'local copy
' The widths of the columns on the printout
' Each column of the printout takes on the page
' the same percentage as the corresponding column
' on the control. If the second column's width on
' the ListView control is 22% of the control's width,
' the 2nd column of the printout takes up 22% of the
' available width on the page
Dim colWidths() As Single
Public Property Let CurrentY(ByVal vData As Long)
    mvarCurrentY = vData
End Property

Public Property Get CurrentY() As Long
    CurrentY = mvarCurrentY
End Property

Public Property Let CurrentX(ByVal vData As Long)
    mvarCurrentX = vData
End Property


Public Property Get CurrentX() As Long
    CurrentX = mvarCurrentX
End Property

Private Sub Class_Initialize()
PicHeight = 16: PicWidth = 16
BorderDistance = 5
DrawHorizontalLines = True: DrawVerticalLines = True: DrawBorder = True
lngWidthBor = 1: lngWidthHor = 1: lngWidthVer = 1
lngColorHor = vbBlack: lngColorVer = vbBlack: lngColorBor = vbBlack
NumOfRowsPerPage = 10
lngLastRow = 1
RowHeight = 250

End Sub
' The following subroutine sets up the printer's properties
' and then calls the PrintListView subroutine to generate
' the actual printout.
Public Function PrintList(ByVal preview As Boolean) As Boolean
    Dim frm As New frmPreview
    ' An instance of the CommonDialog control
    ' that's be displayed by the LVPrint class's code
    ' THIS CLASS MUST BE INSTALLED AT THE CLIENT
    ' ALONG WITH THE APPLICATION THAT USES IT !!!
    Dim CD As CDLG
    Set CD = New CDLG
    ' Set up the Print dialog box ...
    CD.CancelError = True
    On Error Resume Next
    ' ... and display it
    CD.VBPageSetupDlg PRN.hwnd

    CD.VBPrintDlg PRN.hwnd
     
    'CD.ShowPrinter
    ' If user has clicked Cancel, exit
    If Err.Number = 32755 Then Exit Function
    ' reset default error handling
    On Error GoTo 0
    ' Set up the printout's orientation
'    If CD.Orientation = cdlLandscape Then
'        Printer.Orientation = cdlLandscape
'    Else
'        Printer.Orientation = cdlPortrait
'    End If
    ' Set up the PRN object for the appropriate operation
    ' (preview or printout)
    If preview Then
        Set PRN = frmPreview
    Else
        Set PRN = Printer
        If PRN Is Nothing Then
            Exit Function
        End If
    End If
    ' Set printer's (or preview pane's) font
    ' so that report is printed in
    ' same font as the ListView control
    SetPrinterFont
    ' Clear preview pane if needed
    If PRN Is frmPreview Then
        PRN.Show
        PRN.Cls
    End If
    ' Set the page's margins
    SetMargins
    ' Calculate the widths of the printout's columns
    CalculateColumnWidths
    ' Initialize a couple of local variables
    MorePages = True
    currentRow = 0
    ' and start printing:
    ' Keep printing
    ' while there are more pages to be printed and
    ' we haven't exhausted all rows on the control
'    While MorePages And currentRow < objListView.ListItems.Count
'        If PRN Is frmPreview Then
'           ShowMargins          ' Show the margins on the form
'        Else
'            ' Any special actions that must be performed
'            ' when the output is sent to the printer
'        End If
'        ' Initialize variables for first page
'        mvarCurrentX = LeftMargin: mvarCurrentY = TopMargin
'        PRN.CurrentX = mvarCurrentX: PRN.CurrentY = mvarCurrentY
'        ' Print the column headers
'        ' If the PrintListViewHeaders fails to print,
'        ' (most likely due to a printer error), the
'        ' class terminates the job and returns a False value
'        If PrintHead(PRN) = False Then
'            PrintList = False
'            Exit Function
'        End If
'        ' Print the items
'        PrintBody PRN
'    Wend
'    If Not PRN Is frmPreview Then
'        Printer.EndDoc
'    Else
'        ' code to terminate the preview process
'    End If
'    PrintList = True
End Function
'  Prints the Head(Column header of ListView Control)

Public Function PrintHead() As Boolean

Dim lngKount As Long, rectBox As RECT, strTmp As String, lngBUpScalemode As Long
Dim intAlign As Integer, strIconName As String, intSpace_Btn_Txt_n_Pic As Integer
Dim lngPosY As Long, lngPosX As Long, DX As Long, DY As Long
Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
Dim sngXFac As Single, sngYFac As Single, intPad As Integer

Dim bNewPage As Boolean

Dim lngcnt As Long
sngXFac = 1: sngYFac = 1
intPad = 3
PRN.ForeColor = objListView.ForeColor
lngNumOfColumns = objListView.ColumnHeaders.Count
lngNumOfSubItems = lngNumOfColumns - 1
lngRowHeight = PRN.ScaleY(RowHeight, vbTwips, vbPixels)
If TypeOf PRN Is Printer Then
            PRN.Print
            sngXFac = Screen.TwipsPerPixelX / PRN.TwipsPerPixelX
            sngYFac = Screen.TwipsPerPixelY / PRN.TwipsPerPixelY
            intPad = 3 * sngXFac
End If  'TypeOf PRN Is
lngBUpScalemode = PRN.ScaleMode
PRN.ScaleMode = vbPixels
intSpace_Btn_Txt_n_Pic = PRN.ScaleX(400, vbTwips, vbPixels)

lngPosX = PRN.ScaleX(PosX, vbTwips, vbPixels)
lngPosY = PRN.ScaleY(PosY, vbTwips, vbPixels)

If DrawBorder Then
lngPosX = lngPosX + (BorderDistance * sngXFac)
lngPosY = lngPosY + (BorderDistance * sngYFac)
End If
'SetMargins
' Calculate the widths of the printout's columns
'CalculateColumnWidths
'For lngcount = LBound(colWidths) To UBound(colWidths) - 1
'
'Next

With rectBox
           .Left = lngPosX
           .Top = lngPosY
           
          .Bottom = lngRowHeight + .Top
          
           For lngKount = 1 To lngNumOfColumns
                        .Right = (objListView.ColumnHeaders(lngKount).width * sngXFac) + .Left
                If (lngKount + 1 <= lngNumOfColumns) Then
                If .Right + (objListView.ColumnHeaders(lngKount + 1).width * sngXFac) + intPad >= PageWidth Then
                    'The next one printed won't fit
                 bNewPage = True
                 ColPerPage = lngKount
                End If
                End If
'                 Exit For
'                Else
                     strTmp = objListView.ColumnHeaders(lngKount).Text
                     intAlign = objListView.ColumnHeaders(lngKount).Alignment
                     If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
                     .Left = .Left + intPad: .Right = .Right - intPad
                     .Top = .Top + (2 * sngYFac)
                     If HasPicture = True And intAlign = 0 Then
                         strIconName = objListView.ColumnHeaders(lngKount).Icon
                         If strIconName <> 0 And Sgn(.Right - .Left) <> -1 Then
                                      DX = (.Right - .Left) + (4 * sngXFac): DY = PicWidth
                                      If DX < DY Then DY = DX
                                      PRN.PaintPicture objListView.ColumnHeaderIcons.ListImages(strIconName).Picture, .Left, .Top, (DY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                                     .Left = .Left + intSpace_Btn_Txt_n_Pic
                         End If 'strIconName
                     End If  'HasPicture
                     DrawText PRN.hdc, strTmp, Len(strTmp), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                    .Left = .Left - intPad: .Right = .Right + intPad
                    .Top = .Top - (2 * sngYFac)
                    .Left = .Right
                'End If
            Next  'lngKount = 1 To lngNumOfColumns

            If DrawHorizontalLines Then
                        PRN.DrawWidth = lngWidthHor
                        PRN.Line (lngPosX, .Top)-(.Left, .Bottom - lngRowHeight), lngColorHor 'first line
                        PRN.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor 'second line
            End If  'DrawHorizontalLines
            If DrawBorder Then .Right = .Right + BorderDistance * sngXFac: .Bottom = .Bottom + BorderDistance * sngYFac

           If TypeOf PRN Is Printer Then
                        mvarCurrentX = .Left * PRN.TwipsPerPixelX
                        mvarCurrentY = .Top * PRN.TwipsPerPixelY
            Else
                        mvarCurrentX = .Right * Screen.TwipsPerPixelX
                        mvarCurrentY = .Bottom * Screen.TwipsPerPixelY
            End If

End With  'RectBox
bHeadPrinted = True
PRN.ScaleMode = lngBUpScalemode
End Function


'  Prints the Body (Listitems and SubItems of the Listview Control)

Public Function PrintBody() As Boolean
Dim intAlign As Integer, lngCol As Long
Dim lngRow As Long, rectBox As RECT, lngSubItems As Long, strText As String, lngBottom As Long
Dim lngColWidth As Long, lngBUpScalemode As Long, strIconName As String, intSpace_Btn_Txt_n_Pic As Integer
Dim DX As Long, DY As Long, lngMaxRow As Long
Dim lngPosY As Long, lngPosX As Long
Dim lngNumOfColumns As Long, lngNumOfSubItems As Long, lngRowHeight As Long
Dim sngXFac As Single, sngYFac As Single, intPad As Integer
On Error GoTo ErrHandler

PRN.ForeColor = objListView.ForeColor
intPad = 3: sngXFac = 1: sngYFac = 1
lngNumOfColumns = objListView.ColumnHeaders.Count
lngNumOfSubItems = lngNumOfColumns - 1
lngRowHeight = PRN.ScaleY(RowHeight, vbTwips, vbPixels)

If TypeOf PRN Is Printer Then
            PRN.Print
            sngXFac = Screen.TwipsPerPixelX / PRN.TwipsPerPixelX
            sngYFac = Screen.TwipsPerPixelY / PRN.TwipsPerPixelY
            intPad = 3 * sngXFac
End If

lngBUpScalemode = PRN.ScaleMode
PRN.ScaleMode = vbPixels
intSpace_Btn_Txt_n_Pic = PRN.ScaleX(260, vbTwips, vbPixels)
PRN.DrawWidth = lngWidthHor
lngMaxRow = objListView.ListItems.Count
LastRowPrinted = False

lngPosX = PRN.ScaleX(PosX, vbTwips, vbPixels)
lngPosY = PRN.ScaleY(PosY, vbTwips, vbPixels)
       
 If DrawBorder Then
                lngPosX = lngPosX + (BorderDistance * sngXFac)
                lngPosY = lngPosY + (BorderDistance * sngYFac)
 End If

With rectBox
           
           .Left = lngPosX: .Top = lngPosY
            
            If bHeadPrinted Then .Top = lngPosY + lngRowHeight
                       
            For lngRow = RowFrom To RowTo
                        For lngCol = 1 To lngNumOfColumns 'objListView.ListItems(lngRow).ListSubItems.Count + 1
                                If lngCol = 1 Then
                                        strText = objListView.ListItems(lngRow).Text
                                        PRN.FontBold = objListView.ListItems(lngRow).Bold
                                        PRN.ForeColor = objListView.ListItems(lngRow).ForeColor          '
                                Else
                                        strText = objListView.ListItems(lngRow).SubItems(lngCol - 1)
                                        PRN.FontBold = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).Bold
                                        PRN.ForeColor = objListView.ListItems(lngRow).ListSubItems(lngCol - 1).ForeColor
                                End If  'lngCol = 1
                                intAlign = objListView.ColumnHeaders(lngCol).Alignment
                                If intAlign = 1 Then intAlign = 2 Else If intAlign = 2 Then intAlign = 1
                                .Right = .Left + (objListView.ColumnHeaders(lngCol).width * sngXFac)
                                .Bottom = .Top + lngRowHeight
                                .Left = .Left + intPad: .Right = .Right - intPad
                                .Top = .Top + (2 * sngYFac)
                                
                                If HasPicture Then
                                                                           
                                            If lngCol = 1 Then strIconName = objListView.ListItems(lngRow).SmallIcon Else strIconName = "" ' objListView.ListItems(lngRow).ListSubItems(lngCol - 1).SmallIcon
                                                    If strIconName <> "" Then
                                                                  DX = (.Right - .Left) + (4 * sngYFac): DY = PicWidth
                                                                  If DX < DY Then DY = DX
                                                                  PRN.PaintPicture objListView.SmallIcons.ListImages(strIconName).Picture, .Left, .Top, (DY - 1) * sngXFac, (PicHeight - 1) * sngYFac
                                                                 .Left = .Left + intSpace_Btn_Txt_n_Pic
                                                    End If  ' strIconName <> 0
                                                    
                                End If  'HasPicture
                                
                                DrawText PRN.hdc, strText, Len(strText), rectBox, intAlign Or DT_END_ELLIPSIS Or DT_MODIFYSTRING
                                PRN.FontBold = False: PRN.ForeColor = objListView.ForeColor
                               .Left = .Left - intPad: .Right = .Right + intPad
                               .Top = .Top - (2 * sngYFac)
                                
                               .Left = .Right
                    
                        Next  'lngCol = 1 To lngNumOfColumns
                        If DrawHorizontalLines Then PRN.Line (lngPosX, .Top + lngRowHeight)-(.Left, .Bottom), lngColorHor
                       
                       .Left = lngPosX: lngBottom = .Bottom
                       .Top = .Bottom
                        If lngRow = lngMaxRow Then LastRowPrinted = True: Exit For
            Next  'lngRow = RowFrom To RowTo
            
            If DrawHorizontalLines And (Not bHeadPrinted) Then PRN.Line (lngPosX, lngPosY)-(.Right, lngPosY), lngColorHor
            
            If DrawVerticalLines Then
                        PRN.DrawWidth = lngWidthVer
                        PRN.Line (lngPosX, lngPosY)-(lngPosX, lngBottom), lngColorVer
                        For lngCol = 1 To lngNumOfColumns
                                lngColWidth = lngColWidth + (objListView.ColumnHeaders(lngCol).width * sngXFac)
                                PRN.Line (lngPosX + lngColWidth, lngPosY)-(lngPosX + lngColWidth, lngBottom), lngColorVer
                        Next  'lngCol = 1 To lngNumOfColumns
            End If  'DrawVerticalLines

            If DrawBorder Then
                       PRN.DrawWidth = lngWidthBor
                        lngPosY = lngPosY - (BorderDistance * sngYFac)
                        lngPosX = lngPosX - (BorderDistance * sngXFac)
                        DX = .Right + (BorderDistance * sngXFac)
                        DY = .Bottom + (BorderDistance * sngYFac)
                        
                        PRN.Line (DX, DY)-(DX, lngPosY), lngColorBor
                        PRN.Line -(lngPosX, lngPosY), lngColorBor
                        PRN.Line -(lngPosX, DY), lngColorBor
                        PRN.Line -(DX, DY), lngColorBor
                        .Right = .Right + (BorderDistance * sngXFac): .Bottom = .Bottom + (BorderDistance * sngYFac)
            End If  'DrawBorder
            If TypeOf PRN Is Printer Then
                    mvarCurrentX = .Right * PRN.TwipsPerPixelX
                    mvarCurrentY = .Bottom * PRN.TwipsPerPixelY
            Else
                    mvarCurrentX = .Right * Screen.TwipsPerPixelX
                    mvarCurrentY = .Bottom * Screen.TwipsPerPixelY
            End If
End With  'rectbox
bHeadPrinted = False
PRN.ScaleMode = lngBUpScalemode
Exit Function
ErrHandler:
MsgBox Err.Description
Resume
End Function

'Sets the ListViewControl
Public Property Set ListViewName(ByVal oNewValue As Object)
Set objListView = oNewValue
End Property



'LineWidth and Color setting, Printer can have non integer value such as 1.2, 1.5, 2.1 etc.
Public Sub SetLine(ID As LineID, Optional width As Single = 1#, Optional color As Long = vbBlack)
Dim intWidth As Integer
Dim sngMul As Single
If width = 1 Then intWidth = 1: GoTo DownThere
If TypeOf PRN Is Printer Then
            sngMul = 15 / Printer.TwipsPerPixelX  'Trying to make drawwidth similar in different resolution printer
            intWidth = Abs(sngMul * width)
Else
            intWidth = Abs(Int(width))
            If intWidth = 0 Then intWidth = 1
End If
DownThere:
Select Case ID
            Case HorizontalLines
                       lngWidthHor = intWidth
                       lngColorHor = color
            Case VerticalLines
                       lngWidthVer = intWidth
                       lngColorVer = color
            Case BorderLines
                       lngWidthBor = intWidth
                       lngColorBor = color
End Select
End Sub

'Set RowTo and RowFrom, to make printing several page easier.
Public Sub SetRows()
RowFrom = lngLastRow
RowTo = lngLastRow + (NumOfRowsPerPage - 1)
lngLastRow = RowTo + 1
If lngLastRow > objListView.ListItems.Count Then lngLastRow = 1
End Sub

Private Sub SetMargins()
    ' The margins in the sample application are hardcoded.
    ' Design a small form to prompt user for the page margins
    ' (or switch to VB.NET and use the PageSetup dialog box control)
    LeftMargin = 15
    TopMargin = 20
    ' The useful area on the page is the are of the page
    ' minus the margins
    PageHeight = Printer.ScaleHeight - 2 * TopMargin
    PageWidth = Printer.ScaleWidth - 2 * LeftMargin
    RowHeight = Printer.TextHeight("A")
    YMin = TopMargin
End Sub
Private Sub CalculateColumnWidths()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Calculates the widths of the ListView control's columns
'''     on the output device. The printed columns' widths are
'''     proportional to the width of the control's column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim colWidths(objListView.ColumnHeaders.Count)
    Dim iCol As Integer
    TotalWidth = 0
    ' Calculate width of all columns
    Dim TotalColumnWidth As Single
    For iCol = 0 To objListView.ColumnHeaders.Count - 1
        TotalColumnWidth = TotalColumnWidth + objListView.ColumnHeaders(iCol + 1).width * mmPerTwip
    Next
    ' and now calculate the width of each column
    ' on the printout given the relative widths
    ' of the columns on the ListView control
    ' and the width of the preview/print area.
    For iCol = 0 To objListView.ColumnHeaders.Count - 1
        colWidths(iCol) = (objListView.ColumnHeaders(iCol + 1).width / TotalColumnWidth) * PageWidth
        colWidths(iCol) = colWidths(iCol) * mmPerTwip
        TotalWidth = TotalWidth + colWidths(iCol)
    Next
End Sub

Private Sub SetPrinterFont()
    Dim fnt As New StdFont
    Set fnt = objListView.Font
    PRN.Font.Name = fnt.Name
    PRN.Font.Size = fnt.Size
    PRN.Font.Italic = fnt.Italic
    PRN.Font.Underline = fnt.Underline
    PRN.Font.Strikethrough = fnt.Strikethrough
End Sub
Private Sub ShowMargins()
    PRN.Line (0, 0)-(Printer.ScaleWidth, Printer.ScaleHeight), vbWhite, BF
End Sub
