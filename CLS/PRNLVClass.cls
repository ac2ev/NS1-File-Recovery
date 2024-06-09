VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LVPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Millimeters per twip
Const mmPerTwip As Single = 25.4 / 1440#
' Column alignment constants for the ListView control
Const lvwColumnleft As Integer = 0
Const lvwColumnright As Integer = 1
Const lvwColumncenter As Integer = 2

' A reference to the ListView control we're printing
Public LV As Object
' The report's title
Public ReportTitle As String

' Printout Parameters & Margins
Dim GridColor
Dim PageNo As Integer           '   current page number
' The page's margins
Dim TopMargin As Single, LeftMargin As Single
Dim PageHeight As Single, PageWidth As Single
' The useful width of the page (page width minus margins)
Dim TotalWidth As Single
' Padding around each cell
Dim extraHSpace As Single, extraVSpace As Integer

' Work variables
Dim PRN As Object
' The current position on the page
Dim XPos As Single, YPos As Single
' The default height of each row
Dim RowHeight As Single
' The maximum cell height in each row
' If a cell is broken into multiple lines
' the CellHeight's value is a multiple of
' the RowHeight's value
Dim CellHeight As Single
' Top of page at YMin, bottom of page at YMax
Dim YMin As Single, YMax As Single
' More pages to print?
Dim MorePages As Boolean
' The index on the ListView control
' of the row being printed
Dim currentRow As Integer
' The widths of the columns on the printout
' Each column of the printout takes on the page
' the same percentage as the corresponding column
' on the control. If the second column's width on
' the ListView control is 22% of the control's width,
' the 2nd column of the printout takes up 22% of the
' available width on the page
Dim colWidths() As Single

Public Sub Class_Initialize()
    ReportTitle = "ListView Control Sample Printout"
    ' All distances measured in millimeters
    frmPreview.ScaleMode = 6
    Printer.ScaleMode = 6
    ' The default grid color
    GridColor = RGB(0, 0, 0)
    ' The default orientation
    ' A little extra space between each cell's last
    ' text line and the following separator
    extraVSpace = 1
    ' A little extra space left to the left of a left-aligned string
    ' or to the right of a right-aligned string.
    ' Without this extra space, the first character in each cell will
    ' touch the vertical line that separates it from the previous
    ' cell on the same column.
    extraHSpace = 1
    ' Initialize page counter
    PageNo = 1
    ' The preview form's caption
    frmPreview.Caption = "Printout Preview"
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
    Dim CD As mscomdlg.commondialog
    Set CD = frmPreview.CommonDialog1
    ' Set up the Print dialog box ...
    CD.CancelError = True
    On Error Resume Next
    ' ... and display it
    CD.ShowPrinter
    ' If user has clicked Cancel, exit
    If Err.Number = 32755 Then Exit Function
    ' reset default error handling
    On Error GoTo 0
    ' Set up the printout's orientation
    If CD.Orientation = cdlLandscape Then
        Printer.Orientation = cdlLandscape
    Else
        Printer.Orientation = cdlPortrait
    End If
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
    While MorePages And currentRow < LV.ListItems.Count
        If PRN Is frmPreview Then
           ShowMargins          ' Show the margins on the form
        Else
            ' Any special actions that must be performed
            ' when the output is sent to the printer
        End If
        ' Initialize variables for first page
        XPos = LeftMargin: YPos = TopMargin
        PRN.CurrentX = XPos: PRN.CurrentY = YPos
        ' Print the column headers
        ' If the PrintListViewHeaders fails to print,
        ' (most likely due to a printer error), the
        ' class terminates the job and returns a False value
        If PrintListViewHeaders = False Then
            PrintList = False
            Exit Function
        End If
        ' Print the items
        PrintListView
    Wend
    If Not PRN Is frmPreview Then
        Printer.EndDoc
    Else
        ' code to terminate the preview process
    End If
    PrintList = True
End Function

' The following subroutine prints the control's rows
Private Sub PrintListView()
    ' set up the preview pane's font
    frmPreview.Font = LV.Font
    ' Use row and col variables to keep track of the
    ' current row and column (the coordinates of the
    ' current cell on the control).
    Dim row As Integer, col As Integer
    PRN.CurrentX = XPos
    PRN.CurrentY = YPos
    ' Each cells text is copied into the txtLines array
    ' Long cells are split into multiple text lines, which
    ' are stored in sepate elements of hte txtLines array.
    Dim txtLines() As String
    ' Iterate through the ListView control's ListItems collection
    For row = currentRow To LV.ListItems.Count - 1
        CellHeight = RowHeight
        ' Break each item into multiple lines ...
        txtLines = Split(BreakItemText(LV.ListItems(row + 1).Text, _
                                        colWidths(0) - 2), vbCrLf)
        PRN.CurrentY = YPos
        ' ... and print each line of the item's text in its corresponding cell
        PrintSubItems txtLines, 0
        XPos = XPos + colWidths(0)
        PRN.CurrentY = YPos
        ' Then iterate through the current item's subitems ...
        For col = 1 To LV.ColumnHeaders.Count - 1
            ' break each subitem into multiple text lines (if needed) ...
            If col <= LV.ListItems(row + 1).ListSubItems.Count Then    ' skip column if it doesn't contain subitem
                txtLines = Split(BreakItemText(LV.ListItems(row + 1).ListSubItems(col), _
                                            colWidths(col) - 2), vbCrLf)
                ' ... and then print the lines
                PrintSubItems txtLines, col
                ' Advance the X coordinate for the next cell's
                ' (XPos is the next cell's horizontal coordinate)
                XPos = XPos + colWidths(col)
                PRN.CurrentX = XPos
                PRN.CurrentY = YPos
            End If
        Next
        ' The coordinates of the first cell on the following row
        XPos = LeftMargin
        YPos = YPos + CellHeight
        ' Print a horizontal line between cells
        YPos = YPos + extraVSpace
        PRN.Line (LeftMargin, YPos + 0)-(LeftMargin + TotalWidth, YPos + 0), GridColor
        ' Add padding to vertical coordinate
        YPos = YPos + extraVSpace
        PRN.CurrentX = XPos
        PRN.CurrentY = YPos
        ' Is there room for another item?
        ' (some cells may take more than one line, so we make sure there's
        ' room for two text lines and we hope the last item on the page
        ' doesn't contain a very long string!)
        ' An improved version of this utility should be able to break
        ' a single item between two pages. This version of the utility assumes
        ' that the items of the ListView control are short strings that will
        ' fit in the available width and some items may have to be broken
        ' into 2 or 3 lines of text.
        ' If there's no room for another row, print the last horizontal line
        ' at the bottom of the current page, then print the vertical lines
        ' between columns, increase page counter and start a new page
        If PRN.CurrentY > TopMargin + PageHeight - 2 * RowHeight Then
            YMax = YPos - extraVSpace
            PRN.Line (LeftMargin, YMin)-(LeftMargin, YMax), GridColor
            DrawVerticalLines
            PageNo = PageNo + 1
            ' In Preview mode we prompt the user before
            ' displaying the next page. In Printout mode
            ' we print all pages.
            If PRN Is frmPreview Then
                If MsgBox("View Next Page", vbYesNo) = vbYes Then
                    MorePages = True
                    PRN.Cls
                    Exit Sub
                Else
                    MorePages = False
                    Exit Sub
                End If
            Else
                ' Emit current page and get ready to
                ' print the following item on the next page
                PRN.NewPage
                Exit Sub
            End If
        End If
        currentRow = row + 1
    Next
    YMax = YPos - extraVSpace
    PRN.Line (LeftMargin, YMin)-(LeftMargin, YMax), GridColor
    ' Draw the vertical lines on the last page
    DrawVerticalLines
    MorePages = False
    If PRN Is Printer Then PRN.EndDoc
End Sub

Private Function PrintListViewHeaders() As Boolean
    Dim col As Integer
    ' Print date and page number
    If PrintDatePageNo = False Then
        PrintListViewHeaders = False
        Exit Function
    End If
    ' Use a thicker line for the header separators
    ' and bold typeface for headers
    If PRN Is Printer Then
        PRN.DrawWidth = CInt(0.2 / mmPerTwip)
    Else
        PRN.DrawWidth = 2
    End If
    PRN.Font.Bold = True
    PRN.Line (LeftMargin, YPos)-(LeftMargin + TotalWidth, YPos), GridColor
    ' Move pen at the location where the first
    ' column header will be printed
    XPos = LeftMargin
    YPos = TopMargin + extraVSpace
    PRN.CurrentX = XPos
    PRN.CurrentY = YPos
    Dim txtLines() As String
    ' The minimum height of the row with the headers
    CellHeight = RowHeight
    ' Print column headers
    For col = 1 To LV.ColumnHeaders.Count
        txtLines = Split(BreakItemText(LV.ColumnHeaders(col), _
                                    colWidths(col - 1)), vbCrLf)
        ' If cell contains a single line of text
        If UBound(txtLines) = 0 Then
            PRN.CurrentX = XPos + extraHSpace
            PRN.CurrentY = YPos
            PRN.Print txtLines(0)
        Else
        ' If cell contains multiple lines of text
            Dim i As Integer
            For i = 0 To UBound(txtLines)
                PRN.CurrentX = XPos + extraHSpace
                PRN.Print txtLines(i)
                PRN.CurrentY = YPos + RowHeight * (i + 1)
            Next
            CellHeight = RowHeight * i
        End If
        XPos = XPos + colWidths(col - 1)
        PRN.CurrentX = XPos
        PRN.CurrentY = YPos
    Next
    ' leave a little space below the headers
    YPos = PRN.CurrentY + CellHeight * 1.25
    ' and draw a horizontal line
    PRN.Line (LeftMargin, YPos)-(LeftMargin + TotalWidth, YPos), GridColor
    '''PRN.DrawWidth = 1
    ' Advance a little more after the header separator
    ' before printing the items
    YPos = YPos + extraVSpace
    PRN.CurrentY = YPos
    XPos = LeftMargin
    PRN.CurrentX = XPos
    If PRN Is Printer Then
        PRN.DrawWidth = CInt(0.1 / mmPerTwip)
    Else
        PRN.DrawWidth = 1
    End If
    PRN.Font.Bold = False
    PrintListViewHeaders = True
End Function

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
    RowHeight = PRN.TextHeight("A")
    YMin = TopMargin
End Sub

Private Sub ShowMargins()
    PRN.Line (0, 0)-(Printer.ScaleWidth, Printer.ScaleHeight), vbWhite, BF
End Sub

Private Sub PrintSubItems(lines() As String, ByVal col As Integer)
    Dim asterisks As String
    asterisks = "* * *"
    Dim i As Integer
    ' Iterate through each string in the lines array and print it
    ' with the proper alignment on the output device
    For i = 0 To UBound(lines)
        If PRN.CurrentY < TopMargin + PageHeight Then
            lines(i) = Trim(lines(i))
            If Len(lines(i)) > 0 Then
                ' LEFT ALIGNED COLUMN
                If LV.ColumnHeaders(col + 1).Alignment = lvwColumnleft Then
                    PRN.CurrentX = XPos + extraHSpace
                    PRN.Print lines(i)
                    ' RIGHT-ALIGNED COLUMN
                    ElseIf LV.ColumnHeaders(col + 1).Alignment = lvwColumnright Then
                        PRN.CurrentX = XPos + (colWidths(col) - PRN.TextWidth(lines(i)) - 1)
                        PRN.Print lines(i)
                        Else
                            ' CENTERED COLUMN
                            PRN.CurrentX = XPos + (colWidths(col) - PRN.TextWidth(lines(i)) - 2) / 2 + 1
                            PRN.Print lines(i)
                End If
            End If
        Else
            ' Ran out of room on the page for the current cell!
            ' Insert additional statements to handle this case.
            ' The current version of the software simply skips the
            ' segment of the cell that doesn't fit on the page.
            PRN.CurrentX = XPos + (colWidths(col) - PRN.TextWidth(asterisks) - 2) / 2 + 1
            PRN.Print asterisks
            Exit For
        End If
    Next
    ' CellHeight is the height of the tallest cell in the current row.
    ' After printing the current row's items we advance by the amount
    ' specified by CellHeight, and this point becomes the Y coordinate
    ' of the following line of the printout.
    If CellHeight < PRN.CurrentY - YPos Then
        CellHeight = PRN.CurrentY - YPos
    End If
End Sub

Private Sub DrawVerticalLines()
    Dim j As Integer
    Dim lineX As Single
    ' The X-coordinate of the current vertical line
    lineX = LeftMargin
    ' Iterate through the list's columns and draw a vertical
    ' line between consecutive columns.
    ' At each iteration, lineX is the X-coordinate of the
    ' current vertical line.
    For j = 0 To UBound(colWidths) - 1
        lineX = lineX + colWidths(j)
        PRN.Line (lineX, YMin)-(lineX, YMax), GridColor
    Next
End Sub

Private Function PrintDatePageNo() As Boolean
    PRN.CurrentX = LeftMargin
    PRN.CurrentY = TopMargin / 1.5
    PRN.ForeColor = RGB(0, 0, 255)
    ' THIS IS THE FIRST PRINTING STATEMENT IN THE CODE
    ' AND IT'S A GOOD PLACE TO DETECT ANY PROBLEM IN
    ' THE COMPUTER'S COMMUNICATION WITH THE PRINTER
    On Error Resume Next
    PRN.Print FormatDateTime(Now, vbLongDate)
    If Err.Number <> 0 Then
        PrintDatePageNo = False
        Exit Function
    End If
    On Error GoTo 0
    PRN.CurrentX = Printer.ScaleWidth - LeftMargin - 10
    PRN.CurrentY = TopMargin / 1.5
    PRN.Print "PAGE " & PageNo
    PRN.Font.Size = PRN.Font.Size + 4
    PRN.CurrentX = (PRN.ScaleWidth - PRN.TextWidth(ReportTitle)) / 2
    PRN.CurrentY = TopMargin / 4
    PRN.Print ReportTitle
    PRN.Font.Size = PRN.Font.Size - 4
    PRN.ForeColor = RGB(0, 0, 0)
    PrintDatePageNo = True
End Function

Private Sub SetPrinterFont()
    Dim fnt As New StdFont
    Set fnt = LV.Font
    PRN.Font.Name = fnt.Name
    PRN.Font.Size = fnt.Size
    PRN.Font.Italic = fnt.Italic
    PRN.Font.Underline = fnt.Underline
    PRN.Font.Strikethrough = fnt.Strikethrough
End Sub

'   H E L P E R   F U N C T I O N S

Private Function BreakItemText(ByVal item As String, ByVal width As Single) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Breaks a string into multiple lines that will fit in the
'''     specified width when rendered on the output device (screen or printer).
'''     The item string is broken at work boundaries.
'''         *** YOU MUST PROVIDE AN ADEQUATE WIDTH TO FIT THE LONGEST WORD!
'''         *** THE FUNCTION WON'T BREAK A WORD THAT EXCEEDS THE SPECIFIED
'''         *** WIDTH AND, AS A RESULT, THE STRING WILL LEAK INTO ADJACENT
'''         *** CELLS WHEN PRINTED !!!
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If PRN.TextWidth(item) < width Then
        BreakItemText = item
    Else
        Dim iChar As Integer: iChar = 1
        Dim newitem As String: newitem = ""
        Dim moreWords As Boolean: moreWords = True
        Dim nextWord As String
        While moreWords
            nextWord = GetNextWord(item, iChar)
            iChar = iChar + Len(nextWord)
            If PRN.TextWidth(newitem & nextWord) < width Then
                newitem = newitem & nextWord
            Else
                newitem = newitem & vbCrLf & nextWord
            End If
            BreakItemText = newitem
            If iChar > Len(item) Then
                moreWords = False
            End If
        Wend
    End If
End Function

Private Function GetNextWord(ByVal str As String, ByVal pos As Integer)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Retrieves the following word in the specified string,
'''     starting at character pos in the string
'''     Spaces are added appended to the selected word
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim nextWord As String
    While pos <= Len(str) And Mid(str, pos, 1) <> " "
        nextWord = nextWord & Mid(str, pos, 1)
        pos = pos + 1
    Wend
    While pos <= Len(str) And Mid(str, pos, 1) = " "
        nextWord = nextWord & Mid(str, pos, 1)
        pos = pos + 1
    Wend
    GetNextWord = nextWord
End Function

Private Sub CalculateColumnWidths()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''     Calculates the widths of the ListView control's columns
'''     on the output device. The printed columns' widths are
'''     proportional to the width of the control's column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim colWidths(LV.ColumnHeaders.Count)
    Dim iCol As Integer
    TotalWidth = 0
    ' Calculate width of all columns
    Dim TotalColumnWidth As Single
    For iCol = 0 To LV.ColumnHeaders.Count - 1
        TotalColumnWidth = TotalColumnWidth + LV.ColumnHeaders(iCol + 1).width * mmPerTwip
    Next
    ' and now calculate the width of each column
    ' on the printout given the relative widths
    ' of the columns on the ListView control
    ' and the width of the preview/print area.
    For iCol = 0 To LV.ColumnHeaders.Count - 1
        colWidths(iCol) = (LV.ColumnHeaders(iCol + 1).width / TotalColumnWidth) * PageWidth
        colWidths(iCol) = colWidths(iCol) * mmPerTwip
        TotalWidth = TotalWidth + colWidths(iCol)
    Next
End Sub



