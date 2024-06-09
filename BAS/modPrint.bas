Attribute VB_Name = "modPrint"
Option Explicit
Public posX As Long: Public posy As Long:
Private Type RECT
        Left As Long
        top As Long
        Right As Long
        bottom As Long
End Type
Public Sub PrintListView(ListView As ListView, PaperSize As Long, Orientation As Long, LeftMargin As Long, RightMargin As Long, TopMargin As Long, BottomMargin As Long)
    On Error GoTo PRNror

    Dim RetVal As Long
    Dim PrevFont As String
    Dim PrevFontSize As Integer
    Dim Pages As Integer
    Dim PageNumber As Integer
    Dim RowHeight
    Dim RowsPerPage As Integer
    Dim HWidth As Long
    Dim MaxWidth As Long
    Dim HHeight As Long
    Dim ColText As String
    Dim P As Integer
    Dim R As Integer
    Dim C As Integer
    Dim savecursor
    Dim IconNum As Integer
    Dim sngXFac As Single, sngYFac As Single, ypos As Single
    Dim Offset As Long
    
    Dim PRN As Object
    Set PRN = Printer
    
'    LeftMargin = LeftMargin * 1440
'    RightMargin = RightMargin * 1440
'    TopMargin = TopMargin * 1440
'    BottomMargin = BottomMargin * 1440
    
    sngXFac = 1: sngYFac = 1
    'intPad = 3
    PRN.ForeColor = ListView.ForeColor
    'PRN.BackColor = ListView.BackColor
    
    ' Set the cursor to an hourglass...
    savecursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    PrevFont = PRN.Font.Name
    PRN.Font.Name = ListView.Font.Name
    PrevFontSize = PRN.Font.Size
    PRN.Font.Size = ListView.Font.Size
    If TypeOf PRN Is Printer Then
            'PRN.Print
            sngXFac = Screen.TwipsPerPixelX / PRN.TwipsPerPixelX
            sngYFac = Screen.TwipsPerPixelY / PRN.TwipsPerPixelY
            PRN.Orientation = Orientation
            PRN.PaperSize = PaperSize

    Else
        PRN.Show
        PRN.Cls
    End If  'TypeOf Obj Is

    PRN.ScaleMode = vbPixels
    Offset = PRN.ScaleX(50, vbTwips, vbPixels)
    RowHeight = PRN.TextHeight("Get Line Height At This Font Size...")
    RowHeight = RowHeight + (RowHeight \ 4) ' Add 25% of line height for row spacing
    RowsPerPage = PRN.ScaleHeight \ RowHeight - 7 ' Leave 2 rows for column header & 2 for footer & 3 for horiz line spacing
    
    ' Calculate the number of pages required to print the ListView...
    Pages = ListView.ListItems.Count \ RowsPerPage
    If ListView.ListItems.Count Mod RowsPerPage > 0 Then
        Pages = Pages + 1
    End If
    
    PageNumber = 1
    For P = 1 To Pages
        HWidth = Offset
        HHeight = TopMargin * sngYFac
        
        PRN.CurrentX = HWidth
        PRN.CurrentY = HHeight
        'MaxWidth = 0 'RightMargin + LeftMargin
        For C = 1 To ListView.ColumnHeaders.Count
            HHeight = TopMargin
            

            ' Print the column header...
            ColText = ListView.ColumnHeaders(C).Text
            ' Do a generic test to see if we think this column will fit on the page.
            ' If not, force a new page before we continue...

            If (HWidth + (2 * PRN.TextWidth(ColText) + (15 * sngXFac))) > PRN.ScaleWidth Then
                PRN.DrawWidth = 2
                ypos = PRN.CurrentY
                PRN.Line (PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), 0)-(PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), ypos), vbBlack  'last vertical
                PRN.Line (0, 0)-(0, ypos), vbBlack 'left vertical
                ' Add the page number to the page...
                PRN.Font.Bold = True
                PRN.CurrentX = 0
                PRN.CurrentY = (RowsPerPage + 5) * RowHeight
                PRN.Print fname & vbTab & vbTab; "Page " & PageNumber
                PRN.Font.Bold = False

                ' Force a new page for the rest of the columns...
                 If TypeOf PRN Is Printer Then
                    PRN.NewPage
                 Else
                    MsgBox "New page"
                    PRN.Cls
                 End If
                PageNumber = PageNumber + 1
                HWidth = 0
            End If

            PRN.Font.Bold = True
            PRN.Font.Underline = False

            PRN.DrawWidth = 2
            PRN.Line (0, HHeight)-(PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), HHeight), vbBlack
            PRN.CurrentX = HWidth + Offset
            If C = 1 Then PRN.CurrentX = PRN.CurrentX + PRN.ScaleX(300, vbTwips, vbPixels)

            PRN.CurrentY = PRN.CurrentY + 1
            PRN.Print ColText
            ypos = PRN.CurrentY + RowHeight * 0.1
            PRN.Line (0, ypos)-(PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), ypos), vbBlack
            PRN.Font.Bold = False
            MaxWidth = Max(MaxWidth, PRN.TextWidth(ColText))
            HHeight = HHeight + (1.2 * RowHeight)
            

            
            For R = ((P - 1) * RowsPerPage) + 1 To Min(ListView.ListItems.Count, (P * RowsPerPage))
                
                Select Case C
                    Case 1:
                            If C = 1 Then IconNum = ListView.ListItems(R).SmallIcon Else IconNum = ListView.ListItems(R).ListSubItems(C - 1).ReportIcon
                            If IconNum <> 0 Then
                                PRN.PaintPicture ListView.SmallIcons.ListImages(IconNum).Picture, HWidth + Offset, HHeight, 15 * sngXFac, 15 * sngYFac

                            End If
                              ColText = ListView.ListItems(R).Text
                            
                    Case Else:
                        ColText = ListView.ListItems(R).SubItems(C - 1)
                End Select

                  
                PRN.CurrentX = HWidth + Offset '+ PRN.ScaleX(300, vbTwips, vbPixels)
                If C = 1 Then PRN.CurrentX = PRN.CurrentX + PRN.ScaleX(300, vbTwips, vbPixels)
                PRN.CurrentY = HHeight + RowHeight * 0.1
                PRN.Print ColText
                posX = PRN.CurrentX + PRN.ScaleX(300, vbTwips, vbPixels)
                
                PRN.DrawWidth = 1
                PRN.Line (0, PRN.CurrentY - RowHeight)-(PRN.ScaleWidth - posX, PRN.CurrentY - RowHeight), vbBlack 'first line

                MaxWidth = Max(MaxWidth, PRN.TextWidth(ColText))
                HHeight = HHeight + RowHeight + RowHeight * 0.1
            Next R
            
            posX = HWidth - PRN.ScaleX(50, vbTwips, vbPixels)
            PRN.Line (posX, 0)-(posX, HHeight), vbBlack  'second line
            
            posX = PRN.ScaleX(300, vbTwips, vbPixels) 'PRN.CurrentX + PRN.ScaleX(290, vbTwips, vbPixels)
            PRN.DrawWidth = 2
            PRN.Line (0, PRN.CurrentY + (RowHeight * 0.1))-(PRN.ScaleWidth - posX, PRN.CurrentY + (RowHeight * 0.1)), vbBlack    'first line
                
            PRN.DrawWidth = 1
            posX = HWidth - PRN.ScaleX(50, vbTwips, vbPixels)
            PRN.Line (posX, 0)-(posX, PRN.CurrentY + RowHeight * 0.1), vbBlack 'second line
                           
            HWidth = HWidth + MaxWidth + PRN.ScaleX(400, vbTwips, vbPixels)

        Next C

        
        ' Add the page number to the page...
        PRN.Font.Bold = True
        PRN.CurrentX = 0
        ypos = PRN.CurrentY
        PRN.CurrentY = (RowsPerPage + 5) * RowHeight
        PRN.Print fname & vbTab & vbTab; "Page " & PageNumber
        PRN.Font.Bold = False

        ' Start a new page...
        If P <> Pages Then
                PRN.DrawWidth = 2
                PRN.Line (PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), 0)-(PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), ypos), vbBlack 'last vertical
                PRN.Line (0, 0)-(0, ypos), vbBlack 'left vertical
                 If TypeOf PRN Is Printer Then
                    PRN.NewPage
                 Else
                    MsgBox "New page2"
                    PRN.Cls
                 End If
            PageNumber = PageNumber + 1
        End If
    Next P
    PRN.DrawWidth = 2
    PRN.Line (PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), 0)-(PRN.ScaleWidth - PRN.ScaleX(300, vbTwips, vbPixels), ypos), vbBlack 'last vertical
    PRN.Line (0, 0)-(0, ypos), vbBlack 'left vertical
    ' End the print job...
    If TypeOf PRN Is Printer Then PRN.EndDoc

    ' Restore original PRN settings...
'    PRN.Font.Name = PrevFontName
'    PRN.Font.Size = PrevFontSize

    ' Reset the cursor...
    Screen.MousePointer = savecursor

    Exit Sub

PRNror:
    MsgBox "Unable to print contents of " & ListView.Name & " control.", vbOKOnly, App.Title
    Debug.Print Err.Description
    Printer.EndDoc
    'Resume
End Sub

Function Max(ByVal A As Variant, ByVal B As Variant) As Variant
    If (A > B) Then
        Max = A
    Else
        Max = B
    End If
End Function
Function Min(ByVal A As Variant, ByVal B As Variant) As Variant
    If (B > A) Then
        Min = A
    Else
        Min = B
    End If
End Function
